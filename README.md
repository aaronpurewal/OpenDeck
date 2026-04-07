# OpenDeck

LLM-powered PowerPoint editing that **perfectly matches your slide master**. Give it a natural language instruction, and it surgically modifies exactly what you ask. Every font, color, margin, and bullet style stays pixel-perfect.

## The Problem

If you've worked in consulting or finance, you know: **slide formatting is sacred.** A McKinsey deck has specific fonts. A Goldman pitch book has exact margins. A Deloitte deliverable follows a rigid brand template down to the bullet indent level. Getting these wrong isn't a cosmetic issue. It's a credibility issue.

Right now, an analyst spends hours manually updating slides, changing a revenue figure, adding an executive summary, or restructuring a section, all while painstakingly preserving the formatting that took a design team weeks to build.

Every AI slide tool on the market today (Gamma, Beautiful.ai, Tome, SlidesAI) **regenerates your entire deck from scratch.** That means:

- Your firm's branded master slides with custom fonts, colors, and layouts? **Gone.**
- Your 40-slide board deck that just needs one number updated? **Completely rebuilt from zero.**
- The exact paragraph spacing and bullet hierarchy your MD approved? **Destroyed.**

These tools are useless for professional environments where formatting compliance isn't optional.

## The Magic

**OpenDeck reads your existing PPTX, understands every shape on every slide, and only touches exactly what you ask it to change.** Everything else stays untouched. Every font, every color, every indent, every master slide relationship.

Tell it *"update Q3 revenue to $4.2M and add an executive summary slide"* and it will:
1. Find the exact text run containing the old revenue figure and replace just that run
2. Clone an existing slide that uses the right layout from your master, preserving all designer formatting
3. Fill the new slide with content that matches the density and style of your existing slides
4. Validate that no placeholder text was left behind and no numbers were hallucinated

**The result is indistinguishable from a human edit.** Your MD, your client, your design team: nobody can tell AI touched it. That's the point.

## Why This Is Different

**It's fast.** A two-pass architecture means the LLM generates a lightweight plan first (~3 seconds), you review and approve it, then content generation runs. A full deck modification takes 15 to 45 seconds, not the hours an analyst would spend.

**It's secure.** Run it with a local model (Qwen, LLaMA, DeepSeek via LM Studio) and **your data never leaves your machine.** No slides uploaded to third-party servers. No API calls to cloud providers. Your confidential board decks, M&A materials, and client deliverables stay on your hardware. **Zero data exposure.** This is the only AI slide tool that can run fully on-prem.

**It's architecturally bulletproof.** The LLM never sees your PowerPoint file. It receives abstract JSON and returns abstract JSON. A deterministic executor translates those instructions into precise operations. The AI decides *what* to change; the code decides *how*. No hallucinated Python, no corrupted files, no broken XML.

## Aspose + AI = OpenDeck

OpenDeck is built on top of [Aspose.Slides](https://products.aspose.com/slides/), the industry-standard library for programmatic PowerPoint manipulation. Aspose is what gives OpenDeck the ability to read, edit, and write PPTX files without corrupting them. It's battle-tested, used by Fortune 500 enterprises, and handles every edge case of the Office Open XML format.

But Aspose alone isn't enough. It's a low-level library. You can clone slides, fill placeholders, edit text runs, and create charts, but you have to know exactly which shape to touch and what to write. It has no intelligence. A non-developer can't use it.

**OpenDeck is the AI layer that makes Aspose usable by anyone.**

### Why hasn't Aspose built this themselves?

Aspose does have some AI features, but they're limited to exactly two things:

1. **Generation from scratch** (`GeneratePresentationAsync`): you give it a prompt, it creates a brand new deck. Same pattern as Gamma and Tome. Destroys your branded template.
2. **Whole-deck translation** (`TranslateAsync`): walks every text run and swaps it for a translated version. One verb only: translate every string to a target language.

**That's it.** Across their entire AI namespace, there is no feature that reads an existing deck and applies targeted, natural-language-driven edits. You cannot tell Aspose "update the revenue number on slide 4 to $12M" or "add an executive summary slide after slide 2 matching the style of slide 5." Aspose themselves acknowledge in their own docs that the generator output requires "minor post-editing," but they provide no AI tooling to do that editing.

The gap is enormous. Aspose has the world's best PPTX manipulation engine, but their AI layer uses the LLM as a one-shot content source, not as a planner that reasons over an existing deck.

**OpenDeck bridges that gap.** It introduces:

- **Deck harvesting.** Extracting every slide, shape, text run, table, and chart into structured JSON the LLM can reason over.
- **Two-pass planning.** The LLM first generates a lightweight structural plan (clone this, fill that, edit this run), the user approves, then content is generated.
- **Slide-level addressing.** The LLM references slides and shapes by stable labels, never by fragile indices.
- **Deterministic execution.** A plan walker translates the LLM's JSON into precise Aspose operations with no room for hallucinated code.
- **Validation layer.** Placeholder detection, data integrity checks, and smoke tests that re-open the saved file to catch corruption.

This is the architecture Aspose's AI offering is missing: the plan/execute loop that turns natural language into surgical edits on an existing deck. OpenDeck is what you get when you treat Aspose as a precision instrument and use an LLM to wield it.

## What It Does

- Upload any `.pptx` file
- Describe what you want changed in plain English
- Review the plan before anything executes
- Download the modified deck with all original formatting intact

## How It Works

```
  Upload PPTX          "Add an executive summary"        Modified PPTX
       |                        |                              |
       v                        v                              v
 +-----------+    +-------------------------+    +-------------------+
 |  HARVEST  | -> |   PLAN (LLM Pass 1)     | -> |  EXECUTE          |
 |           |    |   ~3 seconds            |    |                   |
 | Extract   |    | Structural plan +       |    | A. Structure      |
 | full deck |    | content manifest        |    | B. Re-harvest     |
 | state as  |    |                         |    | C. Content (LLM)  |
 | JSON      |    |  [User reviews plan]    |    | D. Write shapes   |
 |           |    |                         |    | E. Validate+Save  |
 +-----------+    +-------------------------+    +-------------------+
```

**Step 1, Harvest:** Load the PPTX via Aspose.Slides and extract every slide, shape, text run, table, and chart into a JSON state dict.

**Step 2, Plan (LLM Pass 1):** Send the deck state and user instruction to the LLM. It returns a structural plan (clone, delete, or reorder slides) and a content manifest (which shapes need new content). The user reviews and approves this plan before anything executes. Takes ~3 seconds.

**Step 3, Execute:** Run structural changes, re-harvest the deck, generate all content (LLM Pass 2), write it into shapes, validate, and save. Takes 15 to 45 seconds depending on scope.

The two-pass design means the user gets a checkpoint between planning and execution. Bad plans get rejected before the expensive content generation runs.

## Architecture

The core design principle is a strict 3-layer separation:

| Layer | Files | Role | Touches LLM? |
|-------|-------|------|--------------|
| **Tools** | `tools.py` | All Aspose PPTX read/write operations | Never |
| **Executor** | `executor.py` | Deterministic plan walker, label resolution | Never |
| **LLM** | `llm.py`, `prompts.py` | Model calls, tool schemas, prompt templates | Yes (only layer) |
| **State** | `state.py` | Deck harvesting, shape extraction | Never |
| **Pipeline** | `pipeline.py` | 3-step orchestration | Calls LLM layer |
| **Validation** | `validation.py` | Placeholder detection, data integrity | LLM for fill validation only |

**The LLM never knows it's editing PowerPoint.** It receives structured JSON describing "slides" and "shapes" in a generic document, and returns structured JSON instructions. This prevents the model from trying to generate code. It stays focused on *what* to change, not *how*.

**Structured output via tool use.** All LLM calls use forced tool use for guaranteed schema-compliant JSON. The LLM "calls" a tool whose input schema defines the expected output structure, so no text-based JSON parsing is needed.

## Features

- **Surgical editing.** Modify individual text runs, paragraphs, table cells, or chart data without touching anything else.
- **Slide cloning.** Clone slides using existing slides as formatting donors, preserving all theme styling.
- **Chart creation.** Create bar, stacked bar, line, pie, doughnut, and column charts with automatic theme color extraction.
- **Table creation.** Create formatted tables with theme-colored headers.
- **Multi-provider LLM support.** Switch between OpenAI, Anthropic Claude, or local models (LM Studio, Ollama) at runtime.
- **Human-in-the-loop.** Review and edit the structural plan before execution.
- **Auto-approve mode.** Skip the review step for batch processing.
- **Placeholder detection.** Automatically finds and fixes unfilled template boilerplate.
- **Data integrity validation.** Verifies financial figures against source data.
- **Smoke testing.** Re-opens saved files to verify they're not corrupted.
- **Live stopwatch.** Tracks execution time with pause during plan review.

## Quick Start

### Prerequisites

- Python 3.10+
- An API key for at least one LLM provider (OpenAI, Anthropic, or a local model server)
- macOS: Homebrew-installed .NET runtime (for Aspose's native bridge)

### Install

```bash
git clone https://github.com/aaronpurewal/OpenDeck.git
cd OpenDeck
pip install -r requirements.txt
```

### Configure

Create a `.env` file in the project root:

```bash
# Pick your LLM provider: "anthropic", "openai", or "local"
SSE_LLM_PROVIDER=anthropic

# Add the API key for your chosen provider
ANTHROPIC_API_KEY=sk-ant-...
# or
OPENAI_API_KEY=sk-...
```

For local models (LM Studio, Ollama, or any OpenAI-compatible server):

```bash
SSE_LLM_PROVIDER=local
SSE_LOCAL_API_BASE=http://localhost:1234/v1
SSE_LOCAL_MODEL=qwen3.5-35b-a3b
```

### Run

```bash
# macOS (needed for Aspose's .NET bridge)
DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH streamlit run app.py

# Linux
streamlit run app.py
```

Open `http://localhost:8501` in your browser. Upload a PPTX, type an instruction, review the plan, approve, and download.

### Aspose License (Required)

An Aspose license is required. Without it, the evaluation mode adds watermarks and truncates text, making the output unusable. The good news: **temporary licenses are free and take 30 seconds to get.**

1. Go to [https://purchase.aspose.com/temporary-license](https://purchase.aspose.com/temporary-license)
2. Fill in your email and request a temporary license
3. Download the `.lic` file and place it in the project root as `Aspose Temporary License.lic`

It's auto-detected on startup. No configuration needed.

## Configuration

All configuration is via environment variables (or `.env` file):

| Variable | Default | Description |
|----------|---------|-------------|
| `SSE_LLM_PROVIDER` | `anthropic` | LLM provider: `anthropic`, `openai`, or `local` |
| `SSE_OPENAI_MODEL` | `gpt-4o-mini` | OpenAI model name |
| `SSE_ANTHROPIC_MODEL` | `claude-opus-4-6` | Anthropic model name |
| `SSE_LOCAL_MODEL` | `qwen3.5-35b-a3b` | Local model name |
| `SSE_LOCAL_API_BASE` | `http://localhost:1234/v1` | Local model server URL |
| `OPENAI_API_KEY` | (none) | OpenAI API key |
| `ANTHROPIC_API_KEY` | (none) | Anthropic API key |
| `SSE_OUTPUT_DIR` | `output` | Directory for generated PPTX files |
| `SSE_TEMP_DIR` | `temp` | Directory for temporary files |

## LLM Providers

The engine is model-agnostic. You can switch providers at runtime in the Streamlit sidebar.

**Anthropic Claude.** Best quality. Uses forced tool use (`tool_choice={"type": "tool", "name": "..."}`) for guaranteed JSON schema compliance.

**OpenAI.** Uses forced function calling (`tool_choice={"type": "function", "function": {"name": "..."}}`) for structured output.

**Local Models (LM Studio, Ollama).** Uses the OpenAI SDK with a custom `base_url`. Tool use via `tool_choice="required"` (LM Studio doesn't support forced-name format). Slower on large decks due to prompt size, but free and private. Tested with Qwen 3.5 35B.

## Running Tests

```bash
DYLD_LIBRARY_PATH=/opt/homebrew/lib:$DYLD_LIBRARY_PATH python -m pytest tests/ -v
```

Tests use real Aspose objects (no mocking) because the .NET bridge behavior is too nuanced to mock reliably. The test suite covers:

- **Harvesting.** Char limit estimation, shape extraction, deck state structure.
- **Tools.** Clone, fill, edit, create chart/table, error handling.
- **Executor.** Label resolution, structural operations, content dispatch.
- **Pipeline.** Smoke test, harvest round-trip, shape name remapping for cloned slides.
- **Validation.** Placeholder detection patterns, number extraction, data integrity.
- **Charts and Tables.** Position slots, chart types, theme colors, coordinate conversion.

## Project Structure

```
OpenDeck/
|-- app.py              Streamlit web UI
|-- pipeline.py         3-step orchestration (harvest -> plan -> execute)
|-- llm.py              Model-agnostic LLM wrapper + tool schemas
|-- prompts.py          3 prompt templates (plan, content, validation)
|-- executor.py         Deterministic plan walker, label resolution
|-- state.py            Deck harvesting, shape extraction, compact_state
|-- tools.py            All Aspose PPTX operations
|-- validation.py       Placeholder detection, data integrity, smoke test
|-- config.py           API keys, model names, constants
|-- requirements.txt    Python dependencies
|-- .env                API keys (not committed)
|-- tests/
|   |-- test_tools.py
|   |-- test_edit_tools.py
|   |-- test_executor.py
|   |-- test_harvester.py
|   |-- test_pipeline.py
|   |-- test_validation.py
|   |-- test_chart_table.py
|   +-- fixtures/       Sample PPTX for testing
+-- .claude/            Claude Code hooks and rules
```

## Tech Stack

- **[Aspose.Slides for Python](https://products.aspose.com/slides/python-net/)**. PPTX manipulation via .NET bridge.
- **[Streamlit](https://streamlit.io/)**. Web UI.
- **[OpenAI SDK](https://github.com/openai/openai-python)**. OpenAI + local model provider.
- **[Anthropic SDK](https://github.com/anthropics/anthropic-sdk-python)**. Claude provider.
- **Python 3.10+**. Type hints with `dict`, `list`, `|` union syntax.

## Contributing

Contributions are welcome!

## Commercial Licensing

This project is licensed under the Business Source License 1.1 (BSL). You may freely use it for non-production, personal, educational, and evaluation purposes.

**For production or commercial use or implementation**, please contact:

**Aaron Purewal** (aaron.purewal@gmail.com)

On April 5, 2029, this software will automatically convert to the Apache License 2.0.

## License

Business Source License 1.1. See [LICENSE](LICENSE) for details.

**Permitted:** Non-production use, personal projects, education, evaluation, contributions.

**Requires commercial license:** Production deployment, internal business tools, SaaS offerings, enterprise use.

**Change date:** April 5, 2029 (converts to Apache 2.0).
