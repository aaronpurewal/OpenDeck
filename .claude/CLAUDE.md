# APTX_CC — Surgical Slide Engine

LLM-powered PowerPoint editing via Aspose.Slides for Python. Transforms user instructions into precise slide modifications without regenerating entire decks.

## Architecture: 3-Layer Separation

| Layer | Files | Responsibility | LLM Contact |
|-------|-------|---------------|-------------|
| **Tools** | `tools.py` | All Aspose read/write operations | Never |
| **Executor** | `executor.py` | Deterministic plan execution, label resolution | Never |
| **LLM** | `llm.py`, `prompts.py` | Model calls, JSON extraction, prompt templates | Only layer |
| **State** | `state.py` | Deck harvesting, shape extraction, char_limit estimation | Never |
| **Pipeline** | `pipeline.py` | Orchestrates the 3-step flow (harvest → plan → execute) | Calls LLM layer |
| **Validation** | `validation.py` | Post-execution checks (placeholders, data integrity, smoke test) | LLM call for fill validation only; edit validation is deterministic |
| **Config** | `config.py` | API keys, model names, constants | Never |
| **App** | `app.py` | Streamlit UI | Never |

### Key Principle
The LLM never knows it's editing PowerPoint. It receives/returns structured JSON. Tools never call the LLM. The executor never calls the LLM. This separation is strict.

## Code Style

- **Functional**: standalone functions, no classes in production code (test classes are fine)
- **snake_case** everywhere
- **Dict returns**: every tool returns `{"status": "ok|error", ...}`
- **Type hints** on function signatures (Python 3.10+ `dict`, `list`, `|` union syntax)
- **No global state**: presentation object passed as first arg to tools
- **Private helpers**: leading underscore (`_find_shape`, `_safe_text_frame`)

## Error Handling Patterns

1. **Aspose property access**: individual `try/except Exception: pass` per property. Aspose's .NET bridge throws unpredictable errors — wrap each property access separately.
2. **`.get_effective()` calls**: use `BaseException` (not `Exception`), because .NET proxy errors are RuntimeError subclasses that can bypass Exception.
3. **NaN checks**: always `math.isnan()` before using `font_height` or similar. Aspose returns NaN for inherited values.
4. **Executor**: captures `traceback.format_exc()` into the execution log on failure.
5. **Tools never raise** to the caller — they return `{"status": "error", "message": "..."}`.

## Chart & Table Creation

- **`create_chart`**: Creates a new chart shape on a slide. 6 allowed types: `clustered_bar`, `stacked_bar`, `line`, `pie`, `doughnut`, `clustered_column`. Positioned via named slots (`center`, `left_half`, `right_half`, `bottom_half`). Auto-extracts theme colors from master slide and applies to series fills.
- **`create_table`**: Creates a new table shape on a slide. Headers get bold formatting and theme-colored background. Same position slots as charts.
- **Position slots**: `_POSITION_SLOTS` maps slot names to `(x, y, w, h)` tuples in EMU. `_inches()` helper converts inches to EMU.
- **Theme colors**: `_get_theme_colors()` extracts accent colors 1-6 from the master theme as hex strings. `_apply_theme_to_chart()` applies them to chart series fills.
- **Datapoint dispatch**: Different chart types need different `add_data_point_for_*_series` methods. `_DATAPOINT_METHOD` maps chart type to the correct method name.

## Aspose-Specific Gotchas

- **NaN font sizes**: `portion_format.font_height` returns NaN for inherited fonts. Always use `_safe_font_height()`.
- **`_safe_*` helpers**: `_safe_text_frame()`, `_safe_font_height()`, `_safe_font_name()`, `_safe_font_bold()`, `_safe_font_italic()`, `_safe_effective_format()` — use these instead of direct access.
- **Donor cloning**: `clone_slide` finds an existing slide using the same layout, duplicates it, then clears text. This preserves designer formatting that `insert_empty_slide` loses.
- **Template paragraph**: `fill_placeholder` finds the first bullet paragraph with indent to use as formatting template for all new content paragraphs.
- **Evaluation watermarks**: Aspose eval version injects "Created with Aspose" watermarks and truncates text. Tests detect this with `_EVAL_MODE`.

## Label System

- Every slide gets a label: `slide_0`, `slide_1`, etc.
- New slides get custom labels: `new_summary_1`, `new_financials_2`
- `label_list` (a Python list) is the sole source of truth for label→index mapping
- Position in `label_list` = Aspose slide index
- Structural operations mutate `label_list` directly (insert, remove, clear+extend)
- The LLM references slides by label, never by numeric index

## Pipeline Flow (3 steps)

1. **step1_harvest**: Load PPTX → `harvest_deck()` → state dict
2. **step2_plan** (LLM Pass 1): state + instruction → structural plan + content manifest
3. **step3_execute**: structural changes → re-harvest → remap shapes → content generation (LLM Pass 2) → content execution → validate → save

## Git Conventions

- Imperative verb to start: "Fix ...", "Add ...", "Update ..."
- No conventional commit prefixes (no `feat:`, `fix:`)
- Concise single-line messages describing the change

## Security

- API keys in `.env` file (gitignored), loaded by `config.py`
- Never hardcode keys, tokens, or secrets in source files
- `.env`, `*.lic`, and credential files must stay out of version control

## Claude Code Hooks

Three hooks in `.claude/hooks/` enforce project conventions automatically:

| Hook | Type | Purpose |
|------|------|---------|
| `layer-guard.sh` | PreToolUse (Edit\|Write) | Blocks cross-layer imports (LLM SDK in tools/executor, Aspose in llm/prompts) |
| `parse-check.sh` | PostToolUse (Edit\|Write) | Runs `py_compile` after every Python file edit |
| `prompt-braces.sh` | PostToolUse (Edit\|Write) | Warns about unescaped `{` in prompts.py |

Configured in `.claude/settings.json`.

## Running Tests

```bash
DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH python -m pytest tests/ -v
```

The `DYLD_LIBRARY_PATH` is needed for Aspose's native .NET bridge on macOS.

## Key Files Quick Reference

| File | Purpose |
|------|---------|
| `tools.py` | All Aspose operations (clone, fill, edit, create_chart, create_table, save) |
| `executor.py` | Deterministic plan walker, label resolution |
| `pipeline.py` | 3-step orchestration (harvest → plan → execute) |
| `llm.py` | Model-agnostic LLM wrapper (OpenAI + Anthropic) |
| `prompts.py` | 3 prompt templates (plan, content, validation) |
| `state.py` | Deck harvesting, shape extraction, compact_state |
| `validation.py` | Placeholder detection, data integrity, smoke test |
| `config.py` | API keys, model names, constants |
| `app.py` | Streamlit web UI |
