---
name: architecture-review
description: Verify changes respect the 3-layer separation (tools/executor/LLM) and project conventions
context: fork
allowed-tools:
  - Read
  - Grep
  - Glob
argument-hint: "Describe the change or area to review for architectural compliance"
---

# Architecture Review

Analyze the specified change or area for compliance with APTX_CC's architectural principles.

## Checks to Perform

### 1. Layer Separation
- **Tools** (`tools.py`): Must not import from `llm.py` or `prompts.py`. Must not call any LLM function.
- **Executor** (`executor.py`): Must not import from `llm.py` or `prompts.py`. Receives plan dict, executes mechanically.
- **LLM** (`llm.py`): Must not import from `tools.py` or `executor.py`. Only file that imports SDK clients.
- **Pipeline** (`pipeline.py`): May call both executor and LLM layers. Orchestration only.

### 2. Dict Contract
- Every tool function returns `{"status": "ok|error", ...}`
- No functions raise exceptions to callers (except LLM layer which raises on provider errors)

### 3. Label Indirection
- Slides referenced by label string, never by hardcoded numeric index
- `label_list` is the sole source of truth for label→index mapping
- Structural operations maintain label_list consistency

### 4. State as Context
- LLM receives `compact_state()` output, not raw Aspose objects
- No Aspose objects cross into the LLM layer
- Prompts describe "documents" and "shapes," not "PowerPoint" or "PPTX"

### 5. No Classes in Production
- Production code uses standalone functions, not classes
- Test classes (pytest) are fine

### 6. Deterministic Execution
- Executor produces identical results given identical plan + presentation
- No randomness, no LLM calls in executor
- All variation comes from LLM layer (plan and content generation)

## Output Format
For each finding:
- **Location**: file:line
- **Violation**: which principle is broken
- **Severity**: critical / high / medium
- **Fix**: concrete code change

$ARGUMENTS
