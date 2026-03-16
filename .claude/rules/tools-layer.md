---
paths:
  - "tools.py"
---

# Tools Layer Conventions

## Dict Return Contract
Every tool function returns `{"status": "ok|error", ...}`. Never raise exceptions to the caller. Errors become `{"status": "error", "message": "..."}`.

## Fill vs Edit vs Create Distinction
- **`fill_*`** functions (`fill_placeholder`, `fill_table`): for NEWLY CLONED slides with cleared content
- **`edit_*`** functions (`edit_run`, `edit_paragraph`, `edit_table_cell`, `edit_table_run`): for EXISTING slides with surgical modifications
- **`create_*`** functions (`create_chart`, `create_table`): create NEW shapes from scratch on any slide
- Fill functions reuse/add paragraphs. Edit functions modify existing runs in-place. Create functions add new Aspose shapes.

## Private Helper Naming
All internal helpers use leading underscore: `_find_shape`, `_safe_text_frame`, `_normalize`, `_truncate_to_fit`, `_clear_slide_content`, `_normalize_para_format`, `_clear_portion_junk`, `_inches`, `_get_theme_colors`, `_apply_theme_to_chart`.

## Chart/Table Position Slots
`create_chart` and `create_table` use named position slots (`center`, `left_half`, `right_half`, `bottom_half`) instead of raw coordinates. Constrained set prevents the LLM from generating arbitrary positions that overlap other content.

## Graceful Degradation Pattern
One Aspose property access per try/except block. Never batch multiple property reads in one try:

```python
# CORRECT
try:
    pf.alignment = tpf.alignment
except Exception:
    pass
try:
    pf.margin_left = tpf.margin_left
except Exception:
    pass

# WRONG — one failure skips all remaining
try:
    pf.alignment = tpf.alignment
    pf.margin_left = tpf.margin_left
    pf.indent = tpf.indent
except Exception:
    pass
```

## Template Paragraph Pattern
`fill_placeholder` finds the first paragraph with non-zero indent (a bullet paragraph) and uses it as a formatting template. All new paragraphs copy bullet type, char, height, font, alignment, margin, and indent from this template. Headers (`[H]`/`[HB]` prefixed lines) get `BulletType.NONE` instead.

## char_limit Enforcement
`fill_placeholder` calls `estimate_char_limit()` and then `_truncate_to_fit()` before writing. This catches LLM-generated text that exceeds the shape's physical capacity. Truncation prefers dropping trailing bullets over mid-sentence cuts.

## Index Validation
Every tool that takes `slide_idx` validates it first:
```python
if slide_idx < 0 or slide_idx >= len(prs.slides):
    return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
```
