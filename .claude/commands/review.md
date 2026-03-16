Review the current changes against the APTX_CC project conventions, organized by layer.

## Tool Layer (`tools.py`)
- [ ] Every function returns `{"status": "ok|error", ...}` — no raised exceptions
- [ ] Each Aspose property access has its own try/except (no batching)
- [ ] `_safe_*` helpers used instead of direct Aspose attribute access
- [ ] NaN checks via `math.isnan()` before using font sizes
- [ ] `char_limit` enforced via `_truncate_to_fit()` before writing text

## State Layer (`state.py`)
- [ ] `BaseException` used for `.get_effective()` calls (not just `Exception`)
- [ ] NaN guards on all `font_height` reads
- [ ] `_safe_effective_format()` tried first, raw fallback second

## Executor Layer (`executor.py`)
- [ ] Labels resolved via `label_list.index()` — no hardcoded indices
- [ ] No LLM calls anywhere in executor
- [ ] `traceback.format_exc()` captured on exceptions
- [ ] Structural operations mutate `label_list` correctly (insert/remove/clear+extend)

## Pipeline Layer (`pipeline.py`)
- [ ] Re-harvest happens after structural changes
- [ ] `_remap_manifest_shapes()` called for cloned slides
- [ ] Misplaced structural actions detected and moved from content_manifest
- [ ] `_call_with_retry()` used for all LLM calls

## Prompt Layer (`prompts.py`)
- [ ] All literal JSON braces doubled (`{{`/`}}`) for `.format()` compatibility
- [ ] No references to "PowerPoint", "PPTX", or "Aspose" in prompt text
- [ ] Format placeholders (`{deck_state}`, `{plan}`) match the `.format()` call sites

## General
- [ ] No hardcoded API keys or secrets
- [ ] Type hints on function signatures
- [ ] No classes in production code (test classes are fine)
- [ ] Private helpers use leading underscore
- [ ] Dict returns, not custom exception types

For each finding, report: file and line, issue category, severity (critical/high/medium), and a concrete fix.

Skip: minor style preferences and naming opinions. If no significant issues found, say so explicitly.
