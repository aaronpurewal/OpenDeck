---
paths:
  - "tests/**/*"
---

# Testing Conventions

## Structure
- **pytest class-based**: group related tests in classes (`TestCloneSlide`, `TestEditRun`)
- Each test file maps to a source file: `test_tools.py` → `tools.py`, `test_executor.py` → `executor.py`
- 6 test files: `test_tools.py`, `test_edit_tools.py`, `test_executor.py`, `test_harvester.py`, `test_pipeline.py`, `test_validation.py`

## Real Aspose Objects — No Mocking
- Tests create real `slides.Presentation()` objects and real shapes
- Never mock Aspose — the .NET bridge behavior is too nuanced (NaN, proxy errors)
- Fixture files in `tests/fixtures/` for sample decks when needed

## `_EVAL_MODE` Detection Pattern
Aspose evaluation version truncates text and adds watermarks. Tests detect this at module level:

```python
_EVAL_MODE = False
try:
    _prs = slides.Presentation()
    # ... write test text and read it back ...
    _EVAL_MODE = "truncated" in _readback.lower() or len(_readback) < expected
except Exception:
    _EVAL_MODE = True
```

Then skip affected tests:
```python
@pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval truncates text")
```

## sys.path Setup
Every test file starts with:
```python
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
```
This allows importing from the project root without installing as a package.

## Fixtures
- Prefer simple helper functions over complex fixtures
- `prs` fixture: loads `tests/fixtures/sample_deck.pptx` or creates `slides.Presentation()`
- Create minimal shapes inline: `slide.shapes.add_auto_shape(slides.ShapeType.RECTANGLE, ...)`

## Assertion Patterns
- **Dict status checks**: `assert result["status"] == "ok"` / `assert result["status"] == "error"`
- **label_list mutations**: verify `label_list` contents after structural operations
- **Relative slide counts**: in eval mode, check `>=` counts (eval adds a watermark slide); exact counts only when `not _EVAL_MODE`
- **Structure over content**: when eval truncates text, verify dict keys and status codes exist rather than asserting exact text values

## Running Tests
```bash
DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH python -m pytest tests/ -v
```
