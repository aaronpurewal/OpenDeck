---
paths:
  - "prompts.py"
  - "llm.py"
---

# Prompts & LLM Layer Conventions

## Three Prompt Templates
| Template | Purpose | Token budget |
|----------|---------|-------------|
| `PLAN_PROMPT` | Pass 1: structural plan + content manifest | 4000 |
| `CONTENT_PROMPT` | Pass 2: generate all text content | 16000 |
| `VALIDATION_PROMPT` | Data integrity check for financial content | 2000 |

## Escaped Braces for `.format()`
Prompts use Python's `str.format()` with `{deck_state}`, `{plan}`, etc. All literal JSON braces in the prompt text must be doubled:
```python
"structural_changes": [
    {{"action": "clone_slide", ...}}
]
```
Single `{` means a format placeholder. Double `{{` renders as a literal `{`. Getting this wrong breaks the prompt silently.

## Backslash Line Continuation
Long prompt strings use `\` at end of line for readability within triple-quoted strings:
```python
PLAN_PROMPT = """You are a document editing planner. You receive a document's \
structural state as JSON and a user instruction..."""
```

## "LLM Never Knows It's PowerPoint"
The prompts describe "a document editing system" with "slides" and "shapes." The model never sees `.pptx`, `python-pptx`, or `Aspose`. This prevents the LLM from trying to generate Python code instead of returning structured JSON.

## Model-Agnostic Provider Pattern (`llm.py`)
`_call_llm()` handles three providers:
- `"openai"`: uses `response_format={"type": "json_object"}` for native JSON mode
- `"anthropic"`: strips markdown fences, uses `_extract_json()` for robust parsing
- `"local"`: uses OpenAI SDK with custom `base_url` (LM Studio / Ollama), `_extract_json()` for parsing, `temperature=0.3` for consistency
- Provider selected via `config.LLM_PROVIDER` or function arg

## Lazy SDK Imports
SDK imports happen inside `_call_llm()`, not at module level:
```python
if provider == "openai":
    from openai import OpenAI
```
This avoids ImportError when only one SDK is installed. The `"local"` provider also uses the OpenAI SDK (with custom `base_url`).

## `_extract_json()` — Brace-Depth Parser
Handles Anthropic models that append explanation after JSON. Tracks `{`/`}` depth while respecting string literals and escape sequences. Falls back to this when `json.loads()` fails with "Extra data."

## Two Public LLM Functions + validate_data
- `generate_structure_plan(deck_state_json, user_instruction, provider)` → dict
- `generate_content(plan_json, deck_state_json, provider)` → dict
- `validate_data(source_json, generated_text, provider)` → `{"accurate": bool, "discrepancies": [...]}`
