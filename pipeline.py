"""
Pipeline: Two-step orchestration with user checkpoint.

Step 1: Harvest deck state (Aspose, instant)
Step 2: Generate structural plan (LLM Pass 1, ~3 seconds)
  -- User reviews and approves --
Step 3: Execute plan deterministically
  A: Structural changes (Aspose, instant)
  B: Generate content (LLM Pass 2, ~8-30 seconds)
  C: Execute content updates (Aspose, instant)
  D: Validate (placeholder detection + data integrity)
  E: Save + smoke test
"""

import json
import os
import aspose.slides as slides

from state import harvest_deck, compact_state
from llm import generate_structure_plan, generate_content
from executor import execute_plan
from validation import check_placeholders, validate_data_integrity, smoke_test
from config import LLM_PROVIDER, DEFAULT_OUTPUT_DIR, MAX_LLM_RETRIES


_ACTION_TO_SHAPE_TYPE = {
    "fill_placeholder": "text", "edit_run": "text", "edit_paragraph": "text",
    "fill_table": "table", "edit_table_cell": "table", "edit_table_run": "table",
    "update_chart": "chart",
    "create_chart": "chart", "create_table": "table",
}


def _remap_content_shapes(content: dict, plan: dict, post_state: dict,
                          cloned_labels: set):
    """
    Fix content output shape names using the already-remapped manifest.

    The content LLM (especially weaker local models) often ignores the
    instruction to use document-state shape names and outputs layout
    placeholder names instead (e.g. "Holder 2" instead of "object 2").

    Strategy: for each content_update targeting a cloned slide, find the
    matching manifest entry (by slide_label + action + position index)
    and copy its remapped shape_name.
    """
    # Build label → actual shape names from post-structural state
    slide_lookup = {}
    label_list = post_state.get("label_list", [])
    slides_data = post_state.get("slides", [])
    for i, label in enumerate(label_list):
        if i < len(slides_data):
            slide_lookup[label] = {
                s["name"] for s in slides_data[i].get("shapes", [])
            }

    manifest = plan.get("content_manifest", [])

    for update in content.get("content_updates", []):
        slide_label = update.get("slide_label", "")
        if slide_label not in cloned_labels:
            continue

        actual_names = slide_lookup.get(slide_label, set())
        action = update.get("action", "")

        # Skip actions that create new shapes (no shape_name to fix)
        if action in ("create_chart", "create_table"):
            continue

        # If the shape name already exists on the slide, no fix needed
        if update.get("shape_name", "") in actual_names:
            continue

        # Find the matching manifest entry by slide_label + action type
        # + position index (nth action of this type for this slide)
        target_type = _ACTION_TO_SHAPE_TYPE.get(action)
        if not target_type:
            continue

        # Count this update's position index among same slide+type updates
        update_idx = 0
        for prev in content.get("content_updates", []):
            if prev is update:
                break
            if (prev.get("slide_label") == slide_label and
                    _ACTION_TO_SHAPE_TYPE.get(prev.get("action", "")) == target_type):
                update_idx += 1

        # Find the matching manifest entry at the same position index
        manifest_idx = 0
        for m_entry in manifest:
            if (m_entry.get("slide_label") == slide_label and
                    _ACTION_TO_SHAPE_TYPE.get(m_entry.get("action", "")) == target_type):
                if manifest_idx == update_idx:
                    remapped_name = m_entry.get("shape_name", "")
                    if remapped_name and remapped_name in actual_names:
                        update["shape_name"] = remapped_name
                    break
                manifest_idx += 1


def _remap_manifest_shapes(plan: dict, post_state: dict, cloned_labels: set):
    """
    Fix manifest shape names for cloned slides after donor cloning.

    Layout placeholder names (e.g. "Holder 2") don't match the donor slide's
    actual shape names (e.g. "object 2"). Remap by matching shape type and
    position order within the slide.
    """
    # Build label → slide data lookup using label_list (executor labels)
    # label_list[i] maps to slides[i] by position — harvest_deck labels
    # (slide_0, slide_1) differ from executor labels (new_exec_summary)
    slide_lookup = {}
    label_list = post_state.get("label_list", [])
    slides = post_state.get("slides", [])
    for i, label in enumerate(label_list):
        if i < len(slides):
            slide_lookup[label] = slides[i]

    for entry in plan.get("content_manifest", []):
        slide_label = entry.get("slide_label", "")
        if slide_label not in cloned_labels:
            continue
        slide_data = slide_lookup.get(slide_label)
        if not slide_data:
            continue

        # Check if the manifest shape name actually exists on the slide
        actual_names = {s["name"] for s in slide_data.get("shapes", [])}
        if entry.get("shape_name", "") in actual_names:
            continue  # Name already matches, no remap needed

        # Find shapes of matching type on the actual slide
        target_type = _ACTION_TO_SHAPE_TYPE.get(entry.get("action", ""))
        if not target_type:
            continue
        matching_shapes = [s["name"] for s in slide_data.get("shapes", [])
                          if s.get("type") == target_type]

        # Count how many previous manifest entries target the same slide+type
        # to determine position index
        idx = 0
        for prev in plan.get("content_manifest", []):
            if prev is entry:
                break
            if (prev.get("slide_label") == slide_label and
                    _ACTION_TO_SHAPE_TYPE.get(prev.get("action", "")) == target_type):
                idx += 1

        if idx < len(matching_shapes):
            entry["shape_name"] = matching_shapes[idx]


def step1_harvest(input_path: str) -> tuple:
    """
    Load the deck and harvest its state.
    Returns (Aspose Presentation object, state dict).
    Called once when the user uploads a file.
    """
    prs = slides.Presentation(input_path)
    deck_state = harvest_deck(prs)
    return prs, deck_state


def step2_plan(deck_state: dict, user_instruction: str,
               provider: str = None) -> dict | None:
    """
    Pass 1: Generate the structural plan + content manifest.
    Fast (~3 seconds). Returns the plan for user review.
    Returns None if JSON parsing fails after retries.
    """
    if provider is None:
        provider = LLM_PROVIDER
    # Use compact state for LLM context to stay within token limits
    compact = compact_state(deck_state)
    deck_state_json = json.dumps(compact, indent=2)
    return _call_with_retry(
        generate_structure_plan, deck_state_json, user_instruction, provider
    )


def step3_execute(plan: dict, deck_state: dict, prs,
                  provider: str = None, output_path: str = None) -> dict:
    """
    After user approves the plan:
    1. Execute structural changes (Aspose, instant)
    2. Pass 2: Generate all content (LLM, ~8-30 seconds)
    3. Execute content updates (Aspose, instant)
    4. Validate and save

    Returns execution result with log and warnings.
    """
    if provider is None:
        provider = LLM_PROVIDER
    if output_path is None:
        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(DEFAULT_OUTPUT_DIR, "output.pptx")

    label_list = deck_state["label_list"].copy()
    compact = compact_state(deck_state)
    deck_state_json = json.dumps(compact, indent=2)
    log = []

    # --- Fix misplaced actions: LLM sometimes puts structural ops in manifest ---
    _STRUCTURAL_ACTIONS = {"clone_slide", "delete_slides", "reorder_slides", "duplicate_slide"}
    structural_changes = list(plan.get("structural_changes", []))
    content_manifest = []
    for item in plan.get("content_manifest", []):
        if item.get("action") in _STRUCTURAL_ACTIONS:
            structural_changes.append(item)
        else:
            content_manifest.append(item)
    plan = {**plan, "structural_changes": structural_changes, "content_manifest": content_manifest}

    # --- Phase A: Execute structural changes immediately ---
    structural_plan = {
        "structural_changes": plan.get("structural_changes", []),
        "content_updates": []
    }
    struct_result = execute_plan(structural_plan, prs, label_list)

    if struct_result["status"] == "structural_failure":
        failed_at = struct_result.get("failed_at", "unknown")
        last_error = ""
        for entry in struct_result["log"]:
            if entry.get("status") == "error":
                last_error = entry.get("message", "")
        return {"status": "error",
                "message": f"Structural operation failed at '{failed_at}': {last_error}",
                "log": struct_result["log"]}
    log.extend(struct_result["log"])

    # --- Re-harvest after structural changes ---
    # The deck has changed (slides cloned/deleted/reordered).
    # Re-harvest so the content LLM sees actual shape names on new slides.
    post_struct_state = harvest_deck(prs)
    post_struct_state["label_list"] = label_list.copy()  # Use executor's label list

    # --- Fix manifest shape names for cloned slides ---
    # Pass 1 uses layout placeholder names (e.g. "Holder 2") but after
    # donor cloning, the actual shapes have different names (e.g. "object 2").
    # Remap by matching shape type and position order.
    cloned_labels = set()
    for step in plan.get("structural_changes", []):
        if step.get("action") == "clone_slide":
            cloned_labels.add(step.get("label", ""))

    if cloned_labels:
        _remap_manifest_shapes(plan, post_struct_state, cloned_labels)

    deck_state_json = json.dumps(compact_state(post_struct_state), indent=2)

    # --- Phase B: Generate content (Pass 2 LLM call, the slow part) ---
    plan_json = json.dumps(plan, indent=2)
    content = _call_with_retry(
        generate_content, plan_json, deck_state_json, provider
    )

    if content is None:
        return {"status": "error",
                "message": "LLM failed to generate content after retries"}

    # --- Fix content shape names for cloned slides ---
    # The content LLM (especially local models) often outputs layout
    # placeholder names instead of actual post-clone shape names.
    # Correct them using the already-remapped manifest as ground truth.
    if cloned_labels:
        _remap_content_shapes(content, plan, post_struct_state, cloned_labels)

    # --- Phase C: Execute content updates (Aspose, instant) ---
    content_plan = {
        "structural_changes": [],
        "content_updates": content.get("content_updates", [])
    }
    content_result = execute_plan(content_plan, prs, label_list)
    log.extend(content_result["log"])

    # --- Phase C.5: Auto-fit tables that overflow slide bottom ---
    # Post-write safety net: pre-write truncation is the primary defense
    # but some edits still push tables past slide bounds. Shrink fonts
    # in the tallest rows until the table fits.
    from tools import fit_tables_to_slide
    edited_slide_labels = set()
    for step in content.get("content_updates", []):
        lbl = step.get("slide_label")
        if lbl:
            edited_slide_labels.add(lbl)
        if step.get("action") == "swap_table_sections":
            for key in ("slide_label_a", "slide_label_b"):
                l2 = step.get(key)
                if l2:
                    edited_slide_labels.add(l2)
    fit_warnings = []
    for lbl in edited_slide_labels:
        try:
            s_idx = label_list.index(lbl)
        except ValueError:
            continue
        fit_result = fit_tables_to_slide(
            prs, s_idx,
            bottom_margin=5.0,
            max_iterations=50,
            min_font_height=7.0
        )
        if fit_result.get("shrunk"):
            for entry in fit_result["shrunk"]:
                fit_warnings.append(
                    f"{lbl} {entry['name']}: shrunk fonts "
                    f"({entry['iterations']} iterations, "
                    f"{entry['initial_bottom']:.0f}→{entry['final_bottom']:.0f}pt)"
                )
        if fit_result.get("overflow_remaining"):
            for entry in fit_result["overflow_remaining"]:
                fit_warnings.append(
                    f"{lbl} {entry['name']}: still overflows by "
                    f"{entry['overflow_pt']:.0f}pt "
                    f"({'min font hit' if entry.get('hit_floor') else 'unknown'})"
                )

    # --- Phase D: Validate ---
    # Placeholder detection
    placeholder_result = check_placeholders(prs)
    if placeholder_result["status"] == "placeholders_found":
        # Build a targeted fix manifest
        fix_manifest = []
        for f in placeholder_result["findings"]:
            fix_manifest.append({
                "action": "fill_placeholder",
                "slide_label": f"slide_{f['slide_idx']}" if "slide_label" not in f else f["slide_label"],
                "shape_name": f["shape_name"],
                "instruction": f"Replace placeholder text: {f['text'][:50]}"
            })

        fix_plan = json.dumps({"content_manifest": fix_manifest}, indent=2)
        fix_content = _call_with_retry(
            generate_content, fix_plan, deck_state_json, provider
        )
        if fix_content:
            execute_plan(
                {"structural_changes": [],
                 "content_updates": fix_content.get("content_updates", [])},
                prs, label_list
            )

    # Data integrity check for financial content
    financial_updates = [
        s for s in content.get("content_updates", [])
        if any(kw in json.dumps(s).lower()
               for kw in ["revenue", "ebitda", "$", "%", "margin"])
    ]
    data_warnings = []
    if financial_updates:
        integrity = validate_data_integrity(financial_updates, deck_state, provider)
        if not integrity.get("accurate", True):
            data_warnings = integrity.get("discrepancies", [])

    # --- Phase E: Save and smoke test ---
    prs.save(output_path, slides.export.SaveFormat.PPTX)
    smoke = smoke_test(output_path)
    if smoke["status"] != "ok":
        return {"status": "error", "message": "Smoke test failed",
                "detail": smoke.get("message", "")}

    return {
        "status": "complete",
        "output_path": output_path,
        "log": log,
        "data_warnings": data_warnings,
        "fit_warnings": fit_warnings,
        "placeholder_check": placeholder_result["status"]
    }


def _call_with_retry(fn, *args, max_retries: int = None):
    """
    Call an LLM function and retry on JSON parse failure.
    Works with any function that returns a dict.
    """
    if max_retries is None:
        max_retries = MAX_LLM_RETRIES
    last_error = None
    for attempt in range(max_retries):
        try:
            result = fn(*args)
            if isinstance(result, dict):
                return result
            return json.loads(result)
        except (json.JSONDecodeError, TypeError, KeyError) as e:
            last_error = f"Attempt {attempt + 1}: {type(e).__name__}: {e}"
            print(f"[RETRY] {last_error}")
            continue
        except Exception as e:
            last_error = f"Attempt {attempt + 1}: {type(e).__name__}: {e}"
            print(f"[RETRY] {last_error}")
            continue
    print(f"[FAILED] All {max_retries} attempts failed. Last error: {last_error}")
    return None
