"""
Deterministic Plan Executor.

Receives a plan JSON and walks it mechanically. No LLM calls happen here.
The executor resolves slide labels to indices internally — the LLM never
does index arithmetic.
"""

import traceback

from tools import (
    clone_slide, delete_slides, reorder_slides, duplicate_slide,
    fill_placeholder, fill_table, edit_run, edit_paragraph,
    edit_table_cell, edit_table_run, update_chart,
    create_chart, create_table,
    move_shape, swap_shape_positions, set_shape_fill, swap_table_rows
)

STRUCTURAL_DISPATCH = {
    "clone_slide": clone_slide,
    "delete_slides": delete_slides,
    "reorder_slides": reorder_slides,
    "duplicate_slide": duplicate_slide,
}

CONTENT_DISPATCH = {
    # CREATE: fill new/empty slides from cloned layouts
    "fill_placeholder": fill_placeholder,
    "fill_table": fill_table,
    # CREATE: new charts and tables from scratch
    "create_chart": create_chart,
    "create_table": create_table,
    # EDIT: surgically modify existing slides (per-run targeting)
    "edit_run": edit_run,
    "edit_paragraph": edit_paragraph,
    "edit_table_cell": edit_table_cell,
    "edit_table_run": edit_table_run,
    "update_chart": update_chart,
    # GEOMETRY: move/swap/recolor shapes and overlay-aware row swap
    "move_shape": move_shape,
    "swap_shape_positions": swap_shape_positions,
    "set_shape_fill": set_shape_fill,
    "swap_table_rows": swap_table_rows,
}


def execute_plan(plan: dict, prs, label_list: list) -> dict:
    """
    Execute a plan deterministically. No LLM calls.

    label_list: an ordered list of slide labels, where position = current
    Aspose index. Example: ["slide_0", "slide_1", "slide_2", "slide_3"]
    means label "slide_2" is at index 2 in the Aspose presentation.

    This is the ONLY source of truth for label-to-index resolution.
    Structural operations modify this list directly.

    Returns an execution log with status of each operation.
    """
    log = []

    def resolve(label):
        """Resolve a label to its current Aspose index."""
        try:
            return label_list.index(label)
        except ValueError:
            return None

    # Phase 1: Structural changes
    # Sort: clones/duplicates first, then deletes, then reorders.
    # Ensures donor slides exist when cloning, regardless of LLM plan order.
    _ACTION_PRIORITY = {"clone_slide": 0, "duplicate_slide": 0,
                        "delete_slides": 1, "reorder_slides": 2}
    sorted_structural = sorted(
        plan.get("structural_changes", []),
        key=lambda s: _ACTION_PRIORITY.get(s.get("action", ""), 1)
    )
    for step in sorted_structural:
        action = step["action"]
        args = step.get("args", {})
        try:
            if action == "delete_slides":
                labels_to_delete = args.get("labels", [])
                # Resolve labels to indices, filter out None
                indices = [resolve(l) for l in labels_to_delete]
                indices = [i for i in indices if i is not None]
                # Delete in reverse index order to avoid shifting during deletion
                indices_sorted = sorted(indices, reverse=True)
                for idx in indices_sorted:
                    prs.slides.remove_at(idx)
                # Remove from label_list
                for l in labels_to_delete:
                    if l in label_list:
                        label_list.remove(l)
                log.append({"action": action, "status": "ok",
                           "deleted": labels_to_delete})

            elif action == "clone_slide":
                layout_name = args.get("layout_name")
                insert_at = args.get("insert_at", len(label_list))
                result = clone_slide(prs, layout_name=layout_name,
                                    insert_at=insert_at)
                if result["status"] != "ok":
                    log.append({"action": action, "status": "error",
                               "message": result.get("message", "clone failed")})
                    return {"status": "structural_failure", "log": log,
                            "failed_at": action}
                new_label = step.get("label", f"new_{len(label_list)}")
                label_list.insert(insert_at, new_label)
                donor_idx = result.get("donor_idx")
                log.append({"action": action, "status": "ok",
                           "label": new_label, "donor_idx": donor_idx})

            elif action == "reorder_slides":
                label_order = args.get("label_order", [])
                # Build index mapping
                index_order = [resolve(l) for l in label_order]
                if None in index_order:
                    missing = [l for l, i in zip(label_order, index_order)
                              if i is None]
                    log.append({"action": action, "status": "error",
                               "message": f"Unknown labels: {missing}"})
                    return {"status": "structural_failure", "log": log,
                            "failed_at": action}
                reorder_slides(prs, order=index_order)
                label_list.clear()
                label_list.extend(label_order)
                log.append({"action": action, "status": "ok"})

            elif action == "duplicate_slide":
                source_label = args.get("source_label")
                insert_at = args.get("insert_at", len(label_list))
                source_idx = resolve(source_label)
                if source_idx is None:
                    log.append({"action": action, "status": "error",
                               "message": f"Unknown label: {source_label}"})
                    return {"status": "structural_failure", "log": log,
                            "failed_at": action}
                result = duplicate_slide(prs, source_idx=source_idx,
                                        insert_at=insert_at)
                if result["status"] != "ok":
                    log.append({"action": action, "status": "error",
                               "message": result.get("message", "duplicate failed")})
                    return {"status": "structural_failure", "log": log,
                            "failed_at": action}
                new_label = step.get("label", f"dup_{source_label}")
                label_list.insert(insert_at, new_label)
                log.append({"action": action, "status": "ok",
                           "label": new_label})

            else:
                log.append({"action": action, "status": "error",
                           "message": f"Unknown structural action: {action}"})

        except Exception as e:
            tb = traceback.format_exc()
            log.append({"action": action, "status": "error",
                       "message": str(e), "traceback": tb,
                       "step": step})
            return {"status": "structural_failure", "log": log,
                    "failed_at": action}

    # Phase 2: Content updates (order doesn't matter, all independent)
    for step in plan.get("content_updates", []):
        action = step["action"]
        fn = CONTENT_DISPATCH.get(action)
        if not fn:
            log.append({"action": action, "status": "error",
                       "message": f"Unknown content action: {action}"})
            continue
        try:
            # Resolve slide label to current index
            slide_label = step.get("slide_label")
            slide_idx = resolve(slide_label)
            if slide_idx is None:
                log.append({"action": action, "status": "error",
                           "message": f"Unknown slide label: {slide_label}"})
                continue

            # Build args, replacing label with resolved index
            # Filter out keys that are metadata, not function args
            skip_keys = {"action", "reasoning", "slide_label", "instruction",
                        "char_limit", "columns"}
            args = {k: v for k, v in step.items() if k not in skip_keys}
            args["slide_idx"] = slide_idx

            # Validate required args for table operations
            if action in ("edit_table_cell", "edit_table_run"):
                required = ["row_idx", "col_idx"]
                missing = [k for k in required if k not in args]
                if missing:
                    log.append({"action": action, "status": "warning",
                               "message": f"Missing {missing} for table operation"})
                    continue

            result = fn(prs, **args)
            log.append({"action": action,
                       "shape": step.get("shape_name"),
                       "status": result.get("status", "ok"),
                       "message": result.get("message", "")})
        except Exception as e:
            tb = traceback.format_exc()
            log.append({"action": action,
                       "shape": step.get("shape_name"),
                       "status": "error", "message": str(e),
                       "traceback": tb, "step": step})

    return {"status": "complete", "log": log}
