---
paths:
  - "pipeline.py"
  - "executor.py"
---

# Pipeline & Executor Conventions

## 5 Phases of step3_execute

| Phase | What Happens | Speed |
|-------|-------------|-------|
| **A: Structural** | Clone/delete/reorder slides via executor | Instant (Aspose) |
| **B: Re-harvest + Remap** | Re-harvest deck state, fix manifest shape names for cloned slides | Instant |
| **C: Content Generation** | LLM Pass 2 generates all text | 8-30s |
| **D: Content Execution** | Write generated content into shapes via executor | Instant (Aspose) |
| **E: Validate + Save** | Placeholder detection, data integrity check, smoke test | ~1s |

## Misplaced Action Fix
LLMs sometimes put structural operations (clone_slide, delete_slides) into `content_manifest` instead of `structural_changes`. Pipeline detects and moves them:
```python
_STRUCTURAL_ACTIONS = {"clone_slide", "delete_slides", "reorder_slides", "duplicate_slide"}
for item in plan.get("content_manifest", []):
    if item.get("action") in _STRUCTURAL_ACTIONS:
        structural_changes.append(item)
```

## label_list as Source of Truth
- `label_list` is a Python list where position = Aspose slide index
- `resolve(label)` returns `label_list.index(label)` or None
- Structural operations mutate it directly: `insert()`, `remove()`, `clear()`+`extend()`
- Both pipeline and executor share the same list reference

## Action Priority Sorting (Executor)
Structural actions are sorted before execution:
1. `clone_slide` / `duplicate_slide` (priority 0) — ensure donors exist
2. `delete_slides` (priority 1)
3. `reorder_slides` (priority 2) — must happen last

This ensures donor slides exist when cloning, regardless of LLM plan order.

## Retry Logic
`_call_with_retry()` retries LLM calls up to `MAX_LLM_RETRIES` (default 3) on JSON parse failures. Catches `JSONDecodeError`, `TypeError`, `KeyError`, and general `Exception`.

## Content Dispatch Mapping (Executor)
```python
CONTENT_DISPATCH = {
    "fill_placeholder": fill_placeholder,
    "fill_table": fill_table,
    "edit_run": edit_run,
    "edit_paragraph": edit_paragraph,
    "edit_table_cell": edit_table_cell,
    "edit_table_run": edit_table_run,
    "update_chart": update_chart,
}
```
Content steps are executed in order with `slide_label` resolved to `slide_idx`. Metadata keys (`action`, `reasoning`, `slide_label`, `instruction`, `char_limit`, `columns`) are stripped before passing to tool functions.

## Execution Log Format
Every step appends to `log`:
```python
{"action": "clone_slide", "status": "ok", "label": "new_summary_1", "donor_idx": 2}
{"action": "fill_placeholder", "shape": "Title 1", "status": "ok", "message": ""}
{"action": "edit_run", "shape": "Revenue", "status": "error", "message": "...", "traceback": "..."}
```
Structural failures return `{"status": "structural_failure"}` and halt execution.
