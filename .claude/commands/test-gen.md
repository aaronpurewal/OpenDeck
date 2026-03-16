Generate tests for the specified module or file, following APTX_CC test patterns.

Before generating:
1. Read the source file to understand the public interface
2. Read existing tests in `tests/` to avoid duplicating coverage
3. Check `tests/fixtures/` for available test data

## Required Test File Setup

```python
import pytest
import os
import sys
import tempfile
import aspose.slides as slides

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
```

## _EVAL_MODE Detection (include if tests write/read text)

```python
_EVAL_MODE = False
try:
    _prs = slides.Presentation()
    _layout = _prs.masters[0].layout_slides[0]
    _prs.slides.insert_empty_slide(0, _layout)
    _s = _prs.slides[0].shapes.add_auto_shape(slides.ShapeType.RECTANGLE, 100, 100, 500, 300)
    _s.text_frame.paragraphs[0].portions[0].text = "test_long_text_here"
    _readback = _s.text_frame.paragraphs[0].portions[0].text
    _EVAL_MODE = "truncated" in _readback.lower() or len(_readback) < 19
except Exception:
    _EVAL_MODE = True
```

## Test Patterns
- Use `@pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval truncates text")` for text-dependent tests
- Use pytest class-based grouping: `class TestFunctionName:`
- Create real Aspose `Presentation()` objects — never mock
- Prefer helper functions over complex fixtures

## Assertion Patterns
- Dict status: `assert result["status"] == "ok"`
- label_list mutations: verify list contents after structural ops
- Relative slide counts for eval mode: `assert len(prs.slides) >= expected`
- Temp file cleanup: use `tempfile.NamedTemporaryFile` with `delete=True` or explicit `os.unlink`

## Running
```bash
DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH python -m pytest tests/ -v
```

$ARGUMENTS
