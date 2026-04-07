"""
LLM Interface: Model-agnostic wrapper for OpenAI, Anthropic, and local models.

This is the ONLY file that imports an LLM SDK. Swapping models means
editing this file and config.py — nothing else changes.

Local models (e.g., Qwen 3.5 35B via LM Studio) use the OpenAI SDK
with a custom base_url pointing at the local server.

All providers use TOOL USE for structured output. The LLM "calls" a
tool whose input schema matches the expected JSON structure. This
guarantees schema-compliant JSON without parsing hacks.

Exposes two functions matching the two passes:
  - generate_structure_plan (Pass 1, fast, small output)
  - generate_content (Pass 2, slow, large output)
"""

import json
from config import (
    LLM_PROVIDER, OPENAI_MODEL, ANTHROPIC_MODEL, LOCAL_MODEL,
    OPENAI_API_KEY, ANTHROPIC_API_KEY, LOCAL_API_BASE,
    PLAN_MAX_TOKENS, CONTENT_MAX_TOKENS, VALIDATION_MAX_TOKENS
)


# ---------------------------------------------------------------------------
# Tool Schemas — provider-agnostic JSON Schema definitions
# ---------------------------------------------------------------------------

_STRUCTURAL_ACTIONS = [
    "clone_slide", "delete_slides", "reorder_slides", "duplicate_slide"
]

_CONTENT_ACTIONS = [
    "fill_placeholder", "fill_table", "edit_run", "edit_paragraph",
    "edit_table_cell", "edit_table_run", "update_chart",
    "create_chart", "create_table",
    "move_shape", "swap_shape_positions", "set_shape_fill", "swap_table_rows",
    "swap_table_sections"
]

# Pass 1: structural plan + content manifest
PLAN_SCHEMA = {
    "type": "object",
    "properties": {
        "reasoning": {
            "type": "string",
            "description": "Brief explanation of approach (1-3 sentences)"
        },
        "structural_changes": {
            "type": "array",
            "description": "Ordered list of structural operations",
            "items": {
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": _STRUCTURAL_ACTIONS
                    },
                    "args": {
                        "type": "object",
                        "description": "Action-specific arguments"
                    },
                    "label": {
                        "type": "string",
                        "description": "Label for new slides (clone_slide only)"
                    }
                },
                "required": ["action"]
            }
        },
        "content_manifest": {
            "type": "array",
            "description": "What content to generate in Pass 2",
            "items": {
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": _CONTENT_ACTIONS
                    },
                    "slide_label": {"type": "string"},
                    "shape_name": {"type": "string"},
                    "instruction": {
                        "type": "string",
                        "description": "What content to generate"
                    },
                    "char_limit": {"type": "integer"},
                    "para_idx": {"type": "integer"},
                    "run_match": {"type": "string"},
                    "row_idx": {"type": "integer"},
                    "col_idx": {"type": "integer"},
                    "columns": {"type": "integer"},
                    "chart_type": {
                        "type": "string",
                        "enum": ["clustered_bar", "stacked_bar", "line",
                                 "pie", "doughnut", "clustered_column"]
                    },
                    "position": {
                        "type": "string",
                        "enum": ["center", "left_half", "right_half",
                                 "bottom_half"]
                    },
                    "shape_name_a": {"type": "string"},
                    "shape_name_b": {"type": "string"},
                    "row_idx_a": {"type": "integer"},
                    "row_idx_b": {"type": "integer"},
                    "color_hex": {"type": "string"},
                    "x": {"type": "number"},
                    "y": {"type": "number"},
                    "dx": {"type": "number"},
                    "dy": {"type": "number"},
                    "slide_label_a": {"type": "string"},
                    "slide_label_b": {"type": "string"},
                    "section_idx_a": {"type": "integer"},
                    "section_idx_b": {"type": "integer"}
                },
                "required": ["action", "slide_label"]
            }
        }
    },
    "required": ["structural_changes", "content_manifest"]
}

# Pass 2: content updates
CONTENT_SCHEMA = {
    "type": "object",
    "properties": {
        "content_updates": {
            "type": "array",
            "description": "Content for every manifest item",
            "items": {
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": _CONTENT_ACTIONS
                    },
                    "slide_label": {"type": "string"},
                    "shape_name": {"type": "string"},
                    "text": {"type": "string"},
                    "new_text": {"type": "string"},
                    "para_idx": {"type": "integer"},
                    "run_match": {"type": "string"},
                    "row_idx": {"type": "integer"},
                    "col_idx": {"type": "integer"},
                    "headers": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "rows": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": {}
                        }
                    },
                    "series": {},
                    "chart_type": {
                        "type": "string",
                        "enum": ["clustered_bar", "stacked_bar", "line",
                                 "pie", "doughnut", "clustered_column"]
                    },
                    "title": {"type": "string"},
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "position": {
                        "type": "string",
                        "enum": ["center", "left_half", "right_half",
                                 "bottom_half"]
                    },
                    "col_widths": {
                        "type": "array",
                        "items": {"type": "number"}
                    },
                    "x": {"type": "number"},
                    "y": {"type": "number"},
                    "dx": {"type": "number"},
                    "dy": {"type": "number"},
                    "shape_name_a": {"type": "string"},
                    "shape_name_b": {"type": "string"},
                    "row_idx_a": {"type": "integer"},
                    "row_idx_b": {"type": "integer"},
                    "color_hex": {"type": "string"},
                    "slide_label_a": {"type": "string"},
                    "slide_label_b": {"type": "string"},
                    "section_idx_a": {"type": "integer"},
                    "section_idx_b": {"type": "integer"}
                },
                "required": ["action", "slide_label"]
            }
        }
    },
    "required": ["content_updates"]
}

# Validation: data integrity check
VALIDATION_SCHEMA = {
    "type": "object",
    "properties": {
        "accurate": {"type": "boolean"},
        "discrepancies": {
            "type": "array",
            "items": {"type": "string"}
        }
    },
    "required": ["accurate", "discrepancies"]
}


# ---------------------------------------------------------------------------
# Provider-specific tool formatting
# ---------------------------------------------------------------------------

def _anthropic_tool(name: str, description: str, schema: dict) -> dict:
    """Format a tool definition for the Anthropic API."""
    return {
        "name": name,
        "description": description,
        "input_schema": schema
    }


def _openai_tool(name: str, description: str, schema: dict) -> dict:
    """Format a tool definition for the OpenAI API (and LM Studio)."""
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": description,
            "parameters": schema
        }
    }


# ---------------------------------------------------------------------------
# Fallback: text-based JSON extraction (used when tool use fails)
# ---------------------------------------------------------------------------

def _extract_json(text: str) -> dict:
    """
    Extract the first complete JSON object from text.

    Handles cases where the model appends explanation after the JSON,
    which causes json.loads to fail with 'Extra data'.
    """
    # Try direct parse first (fastest path)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # Find the first '{' and match its closing '}'
    start = text.find("{")
    if start == -1:
        raise json.JSONDecodeError("No JSON object found", text, 0)
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(text)):
        c = text[i]
        if escape:
            escape = False
            continue
        if c == "\\":
            escape = True
            continue
        if c == '"' and not escape:
            in_string = not in_string
            continue
        if in_string:
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start:i + 1])
    raise json.JSONDecodeError("Unterminated JSON object", text, start)


# ---------------------------------------------------------------------------
# Core LLM call with tool use
# ---------------------------------------------------------------------------

def _call_llm(system_prompt: str, user_message: str, provider: str,
              max_tokens: int = 16000,
              tool_name: str = None, tool_schema: dict = None) -> dict:
    """
    Call the LLM and return parsed JSON.

    When tool_name and tool_schema are provided, uses tool use (forced)
    for guaranteed schema-compliant output. Falls back to text-based
    JSON extraction if tool use isn't available or fails.
    """
    if provider == "openai":
        return _call_openai(system_prompt, user_message, max_tokens,
                            tool_name, tool_schema)

    elif provider == "anthropic":
        return _call_anthropic(system_prompt, user_message, max_tokens,
                               tool_name, tool_schema)

    elif provider == "local":
        return _call_local(system_prompt, user_message, max_tokens,
                           tool_name, tool_schema)

    else:
        raise ValueError(f"Unknown provider: {provider}")


def _call_openai(system_prompt: str, user_message: str, max_tokens: int,
                 tool_name: str | None, tool_schema: dict | None) -> dict:
    """OpenAI provider — uses response_format for JSON, tool use if schema provided."""
    from openai import OpenAI
    client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else OpenAI()

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_message}
    ]

    if tool_name and tool_schema:
        tools = [_openai_tool(tool_name, "Return structured output", tool_schema)]
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            tools=tools,
            tool_choice={"type": "function", "function": {"name": tool_name}},
            max_tokens=max_tokens
        )
        return _extract_openai_tool_result(response, tool_name)

    # Fallback: JSON mode
    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=messages,
        response_format={"type": "json_object"},
        max_tokens=max_tokens
    )
    return json.loads(response.choices[0].message.content)


def _call_anthropic(system_prompt: str, user_message: str, max_tokens: int,
                    tool_name: str | None, tool_schema: dict | None) -> dict:
    """Anthropic provider — uses forced tool use for structured output."""
    from anthropic import Anthropic
    client = Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else Anthropic()

    if tool_name and tool_schema:
        tools = [_anthropic_tool(tool_name, "Return structured output", tool_schema)]
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=max_tokens,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}],
            tools=tools,
            tool_choice={"type": "tool", "name": tool_name}
        )
        return _extract_anthropic_tool_result(response)

    # Fallback: text-based extraction
    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}]
    )
    text = response.content[0].text
    text = text.replace("```json", "").replace("```", "").strip()
    return _extract_json(text)


def _call_local(system_prompt: str, user_message: str, max_tokens: int,
                tool_name: str | None, tool_schema: dict | None) -> dict:
    """Local model (LM Studio) — uses OpenAI SDK with tool use."""
    from openai import OpenAI
    client = OpenAI(base_url=LOCAL_API_BASE, api_key="lm-studio")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_message}
    ]

    if tool_name and tool_schema:
        tools = [_openai_tool(tool_name, "Return structured output", tool_schema)]
        try:
            # LM Studio only supports string tool_choice values
            # ("none", "auto", "required"), not the forced-name format.
            # With one tool + "required", the model must call our tool.
            response = client.chat.completions.create(
                model=LOCAL_MODEL,
                messages=messages,
                tools=tools,
                tool_choice="required",
                max_tokens=max_tokens,
                temperature=0.3
            )
            return _extract_openai_tool_result(response, tool_name)
        except Exception:
            # Tool use not supported — fall back to text extraction
            pass

    # Fallback: text-based extraction
    response = client.chat.completions.create(
        model=LOCAL_MODEL,
        messages=messages,
        max_tokens=max_tokens,
        temperature=0.3
    )
    text = response.choices[0].message.content
    text = text.replace("```json", "").replace("```", "").strip()
    return _extract_json(text)


# ---------------------------------------------------------------------------
# Result extraction helpers
# ---------------------------------------------------------------------------

def _extract_anthropic_tool_result(response) -> dict:
    """Extract the tool input dict from an Anthropic tool_use response."""
    for block in response.content:
        if block.type == "tool_use":
            return block.input  # Already a parsed dict
    # No tool_use block found — try text fallback
    for block in response.content:
        if hasattr(block, "text") and block.text:
            return _extract_json(block.text)
    raise ValueError("No tool_use or text block in Anthropic response")


def _extract_openai_tool_result(response, tool_name: str) -> dict:
    """Extract the tool arguments dict from an OpenAI tool_calls response."""
    message = response.choices[0].message
    if message.tool_calls:
        for tc in message.tool_calls:
            if tc.function.name == tool_name:
                return json.loads(tc.function.arguments)
    # Fallback: try message content
    if message.content:
        return _extract_json(message.content)
    raise ValueError(f"No tool call '{tool_name}' in OpenAI response")


# ---------------------------------------------------------------------------
# Public API — the two pipeline passes + validation
# ---------------------------------------------------------------------------

def generate_structure_plan(deck_state_json: str, user_instruction: str,
                            provider: str = None) -> dict:
    """
    Pass 1: Generate structural plan + content manifest.
    Small output (~200-400 tokens). Returns in ~3 seconds.
    Uses forced tool use for guaranteed JSON schema compliance.
    """
    if provider is None:
        provider = LLM_PROVIDER
    from prompts import PLAN_PROMPT
    prompt = PLAN_PROMPT.format(deck_state=deck_state_json)
    return _call_llm(prompt, user_instruction, provider,
                     max_tokens=PLAN_MAX_TOKENS,
                     tool_name="submit_structure_plan",
                     tool_schema=PLAN_SCHEMA)


def generate_content(plan_json: str, deck_state_json: str,
                     provider: str = None) -> dict:
    """
    Pass 2: Generate all text content for the approved plan.
    Large output (~1-5K tokens). Returns in ~8-30 seconds.
    Uses forced tool use for guaranteed JSON schema compliance.
    """
    if provider is None:
        provider = LLM_PROVIDER
    from prompts import CONTENT_PROMPT
    prompt = CONTENT_PROMPT.format(
        plan=plan_json,
        deck_state=deck_state_json
    )
    return _call_llm(prompt, "Generate all content now.", provider,
                     max_tokens=CONTENT_MAX_TOKENS,
                     tool_name="submit_content",
                     tool_schema=CONTENT_SCHEMA)


def validate_data(source_json: str, generated_text: str,
                  provider: str = None) -> dict:
    """
    Data integrity validation: compare generated content against source data.
    Returns {"accurate": bool, "discrepancies": [...]}.
    """
    if provider is None:
        provider = LLM_PROVIDER
    from prompts import VALIDATION_PROMPT
    prompt = VALIDATION_PROMPT.format(
        source_json=source_json,
        generated_text=generated_text
    )
    return _call_llm(prompt, "Validate now.", provider,
                     max_tokens=VALIDATION_MAX_TOKENS,
                     tool_name="submit_validation",
                     tool_schema=VALIDATION_SCHEMA)
