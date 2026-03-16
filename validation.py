"""
Validation Layer: Post-execution checks.

Three layers:
1. Constraint enforcement (pre-generation, handled by char_limit in prompts)
2. Placeholder detection (post-execution, deterministic, zero LLM calls)
3. Data integrity check (post-generation):
   a. Deterministic for edit actions (verify values exist in source data)
   b. LLM call for fill actions (synthesized content needs judgment)

Plus a smoke test: save → re-open → verify file integrity.
"""

import json
import math
import re
import aspose.slides as slides
from config import PLACEHOLDER_PATTERNS


def _safe_text_frame(shape):
    """Get text_frame safely."""
    try:
        tf = shape.text_frame
        if tf is not None:
            return tf
    except Exception:
        pass
    return None


def check_placeholders(prs) -> dict:
    """
    Scan entire deck for unfilled placeholder text.
    Deterministic — no API calls. Catches leftover template boilerplate.
    """
    findings = []
    for slide_idx in range(len(prs.slides)):
        slide = prs.slides[slide_idx]
        for shape in slide.shapes:
            tf = _safe_text_frame(shape)
            if not tf:
                continue
            try:
                text_lower = (tf.text or "").lower().strip()
            except Exception:
                continue
            if not text_lower:
                continue
            for pattern in PLACEHOLDER_PATTERNS:
                if pattern.lower() in text_lower:
                    findings.append({
                        "slide_idx": slide_idx,
                        "shape_name": shape.name,
                        "text": (tf.text or "")[:100],
                        "matched_pattern": pattern
                    })
                    break
    return {
        "status": "clean" if not findings else "placeholders_found",
        "findings": findings
    }


def check_brand(prs, slide_idx: int, brand_rules: dict) -> dict:
    """
    Compare shape properties against brand spec.

    brand_rules example:
    {
        "allowed_fonts": ["Calibri", "Arial"],
        "title_min_size_pt": 24,
        "body_min_size_pt": 10,
    }
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide {slide_idx} out of range"}

    violations = []
    slide = prs.slides[slide_idx]
    allowed_fonts = brand_rules.get("allowed_fonts", [])
    title_min_emu = brand_rules.get("title_min_size_pt", 0) * 12700
    body_min_emu = brand_rules.get("body_min_size_pt", 0) * 12700

    for shape in slide.shapes:
        tf = _safe_text_frame(shape)
        if not tf:
            continue
        for para in tf.paragraphs:
            for portion in para.portions:
                pf = portion.portion_format
                # Font check
                if allowed_fonts:
                    try:
                        font_name = str(pf.latin_font) if pf.latin_font else None
                        if font_name and font_name not in allowed_fonts:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Font '{font_name}' not in allowed list",
                                "value": font_name
                            })
                    except Exception:
                        pass
                # Size check
                try:
                    fh = pf.font_height
                    if fh and not math.isnan(fh) and fh > 0:
                        if "title" in shape.name.lower() and fh < title_min_emu:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Title font too small: {fh / 12700:.0f}pt",
                                "value": fh / 12700
                            })
                        elif fh < body_min_emu:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Body font too small: {fh / 12700:.0f}pt",
                                "value": fh / 12700
                            })
                except Exception:
                    pass

    return {
        "status": "clean" if not violations else "violations_found",
        "violations": violations
    }


_NUMERIC_RE = re.compile(r"\d+[.,]\d+|\d{2,}")

_EDIT_ACTIONS = {"edit_run", "edit_paragraph", "edit_table_cell", "edit_table_run"}


def _extract_numbers(text: str) -> set[str]:
    """Pull all numeric tokens from text for comparison."""
    return set(_NUMERIC_RE.findall(text))


def _collect_source_numbers(deck_state: dict) -> set[str]:
    """Gather every number that appears anywhere in the deck state."""
    all_text = []
    for slide in deck_state.get("slides", []):
        for shape in slide.get("shapes", []):
            if shape.get("text"):
                all_text.append(shape["text"])
            for row in shape.get("rows", []):
                for cell in row:
                    cell_text = cell.get("text", "") if isinstance(cell, dict) else str(cell)
                    if cell_text:
                        all_text.append(cell_text)
    return _extract_numbers(" ".join(all_text))


def _check_edit_deterministic(edit_updates: list,
                              source_numbers: set[str]) -> list[str]:
    """
    Deterministic check for edit actions: verify that every number in the
    generated new_text already exists somewhere in the source deck.

    Returns a list of discrepancy descriptions (empty = all clear).
    """
    discrepancies = []
    for update in edit_updates:
        new_text = update.get("new_text", "")
        generated_nums = _extract_numbers(new_text)
        novel = generated_nums - source_numbers
        if novel:
            discrepancies.append(
                f"{update.get('action', 'edit')} on "
                f"{update.get('slide_label', '?')}/{update.get('shape_name', '?')}: "
                f"novel numbers {novel} not found in source data"
            )
    return discrepancies


def validate_data_integrity(content_updates: list, deck_state: dict,
                            provider: str) -> dict:
    """
    Compare generated content against source data.

    Edit actions (edit_run, edit_paragraph, edit_table_cell, edit_table_run)
    are checked deterministically — verify numbers exist in the source deck.
    Fill actions (fill_placeholder, fill_table) use an LLM call because the
    model synthesizes new prose from multiple data points.
    """
    edit_updates = [u for u in content_updates if u.get("action") in _EDIT_ACTIONS]
    fill_updates = [u for u in content_updates if u.get("action") not in _EDIT_ACTIONS]

    all_discrepancies = []

    # Deterministic check for edits
    if edit_updates:
        source_numbers = _collect_source_numbers(deck_state)
        all_discrepancies.extend(
            _check_edit_deterministic(edit_updates, source_numbers)
        )

    # LLM check for fills (synthesized content needs judgment)
    if fill_updates:
        from llm import validate_data

        source_texts = []
        generated_texts = []

        for update in fill_updates:
            slide_label = update.get("slide_label", "")
            for slide in deck_state.get("slides", []):
                if slide.get("label") == slide_label:
                    for shape in slide.get("shapes", []):
                        if shape.get("text"):
                            source_texts.append(shape["text"])
                    break
            if "text" in update:
                generated_texts.append(update["text"])
            elif "rows" in update:
                generated_texts.append(json.dumps(update["rows"]))

        if source_texts and generated_texts:
            source_json = json.dumps(source_texts, indent=2)
            generated_text = "\n".join(generated_texts)
            try:
                result = validate_data(source_json, generated_text, provider)
                if not result.get("accurate", True):
                    all_discrepancies.extend(result.get("discrepancies", []))
            except Exception as e:
                all_discrepancies.append(f"Validation call failed: {str(e)}")

    return {
        "accurate": len(all_discrepancies) == 0,
        "discrepancies": all_discrepancies
    }


def smoke_test(output_path: str) -> dict:
    """
    Open the saved file and verify it loads cleanly.
    If Aspose can open it, PowerPoint can too.
    """
    try:
        test_prs = slides.Presentation(output_path)
        slide_count = len(test_prs.slides)
        return {"status": "ok", "slide_count": slide_count}
    except Exception as e:
        return {"status": "error", "message": str(e)}
