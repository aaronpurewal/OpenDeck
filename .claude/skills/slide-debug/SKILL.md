---
name: slide-debug
description: Diagnose slide formatting and rendering bugs by tracing data flow through the pipeline
context: fork
allowed-tools:
  - Read
  - Grep
  - Glob
argument-hint: "Describe the rendering issue (e.g., 'text overflows on slide 3', 'bullets have wrong font')"
---

# Slide Debug Investigation

Diagnose a slide rendering or formatting bug by tracing data through the pipeline layers.

## Investigation Procedure

### 1. Identify the Affected Shape
- Find the slide label and shape name from the executor log or user description
- Read the shape's state from `harvest_deck()` output (or re-harvest)
- Note: `char_limit`, paragraph count, font sizes, shape dimensions

### 2. Trace the Pipeline Path
Determine which path the content took:

**If cloned slide** (fill_placeholder/fill_table):
- Was `_find_donor_slide()` used? Check `donor_idx` in executor log
- Did `_remap_manifest_shapes()` fix shape names?
- Was `_clear_slide_content()` called?
- Did `fill_placeholder` find a `template_para`?
- Were `_normalize_para_format()` and `_clear_portion_junk()` applied?

**If existing slide** (edit_run/edit_paragraph/edit_table_cell):
- Was the `run_match` text found? (check `_normalize()` comparison)
- Did the paragraph index exist?

### 3. Variable Checklist
- [ ] `char_limit`: `estimate_char_limit(width, height, font_size_pt)` — is font_size_pt correct or NaN/0?
- [ ] `template_para`: first paragraph with non-zero, non-NaN indent
- [ ] `donor_idx`: index of existing slide using same layout, or None
- [ ] Shape name: manifest name vs actual name after clone
- [ ] `_EVAL_MODE`: is Aspose evaluation version causing text truncation?

### 4. Common Root Causes

| Symptom | Likely Cause | Where to Look |
|---------|-------------|---------------|
| Text overflow | `char_limit` too large, or `_truncate_to_fit` not called | `tools.py:fill_placeholder`, `state.py:estimate_char_limit` |
| Wrong bullets/indent | `template_para` not found or wrong source | `tools.py:_normalize_para_format` |
| Missing content | Shape name mismatch | `pipeline.py:_remap_manifest_shapes` |
| Formatting loss on new paragraphs | Template font not copied | `tools.py:fill_placeholder` (fresh paragraph branch) |
| Ghost blank lines | Extra donor paragraphs not cleared | `tools.py:fill_placeholder` (extra donor branch) |
| "Click to add" blue styling | `_clear_portion_junk` not called | `tools.py:_clear_portion_junk` |
| NaN crashes | `_safe_font_height` not used | `state.py:_safe_font_height` |
| .NET RuntimeError | `Exception` catch instead of `BaseException` | `state.py:_safe_effective_format` |

### 5. Fix Verification
After identifying the root cause:
1. Write a minimal test that reproduces the bug
2. Apply the fix
3. Run `DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH python -m pytest tests/ -v`

$ARGUMENTS
