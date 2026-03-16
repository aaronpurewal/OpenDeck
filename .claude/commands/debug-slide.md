Diagnose a slide rendering or formatting issue using this step-by-step procedure.

## Step 1: Trace the Data Flow
Follow the data through the pipeline to find where it diverges:

1. **Harvest**: Run `harvest_deck(prs)` and inspect the slide's state dict. Check shape names, char_limits, paragraph structure.
2. **Plan (Pass 1)**: Check the structural plan — was the right layout chosen? Were shape names correct?
3. **Remap**: If the slide was cloned, check `_remap_manifest_shapes()` output. Did layout placeholder names get remapped to actual donor shape names?
4. **Content (Pass 2)**: Check the generated content. Did the LLM respect char_limit? Did it use correct shape names from the re-harvested state?
5. **Execution**: Check the executor log for errors/warnings on this slide's content updates.

## Step 2: Check Key Variables
For the affected shape, verify:
- [ ] `char_limit` — is it reasonable for the shape dimensions?
- [ ] `template_para` — was a bullet paragraph found as formatting template? (check `fill_placeholder` logic)
- [ ] `donor_idx` — was a donor slide found for cloning? (check `_find_donor_slide`)
- [ ] Shape name — does the manifest shape name match the actual shape on the slide?
- [ ] Paragraph count — does the donor have enough paragraphs for the content?

## Step 3: Common Root Causes

| Symptom | Root Cause | Fix |
|---------|-----------|-----|
| Text overflows shape | char_limit too generous or LLM ignored it | Check `estimate_char_limit()` params, verify `_truncate_to_fit()` ran |
| Wrong bullet style | `template_para` not found or wrong paragraph used | Check if donor slide has a bullet paragraph with non-zero indent |
| Missing content on cloned slide | Shape name mismatch after donor cloning | Check `_remap_manifest_shapes()` — manifest uses layout names, actual slide uses donor names |
| Formatting loss (wrong font/size) | Fresh paragraphs added without template | Check `fill_placeholder` — new paragraphs should copy font from `template_para` |
| Ghost empty text / blank lines | Extra donor paragraphs not blanked | Verify the `p_idx >= len(new_paragraphs)` branch clears portions |
| "Click to add" styling (blue/underline) | Placeholder default formatting inherited | Check `_clear_portion_junk()` ran on reused portions |
| NaN in output | `font_height` from inherited style | Verify `_safe_font_height()` is used, not raw `portion_format.font_height` |
| .NET proxy error / RuntimeError | `get_effective()` on unsupported property | Verify `BaseException` catch (not just `Exception`) in `_safe_effective_format()` |

## Step 4: Quick Diagnostic Script
```python
from state import harvest_deck
import aspose.slides as slides
import json

prs = slides.Presentation("path/to/deck.pptx")
state = harvest_deck(prs)
# Inspect specific slide
slide_data = state["slides"][SLIDE_INDEX]
print(json.dumps(slide_data, indent=2, default=str))
```

$ARGUMENTS
