"""
LLM Interface: Model-agnostic wrapper for OpenAI and Anthropic.

This is the ONLY file that imports an LLM SDK. Swapping models means
editing this file and config.py — nothing else changes.

Exposes two functions matching the two passes:
  - generate_structure_plan (Pass 1, fast, small output)
  - generate_content (Pass 2, slow, large output)
"""

import json
from config import (
    LLM_PROVIDER, OPENAI_MODEL, ANTHROPIC_MODEL,
    OPENAI_API_KEY, ANTHROPIC_API_KEY,
    PLAN_MAX_TOKENS, CONTENT_MAX_TOKENS, VALIDATION_MAX_TOKENS
)


def _call_llm(system_prompt: str, user_message: str, provider: str,
              max_tokens: int = 16000) -> dict:
    """
    Internal: call the LLM and return parsed JSON.
    Single function handles both providers.
    """
    if provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else OpenAI()
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ],
            response_format={"type": "json_object"},
            max_tokens=max_tokens
        )
        return json.loads(response.choices[0].message.content)

    elif provider == "anthropic":
        from anthropic import Anthropic
        client = Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else Anthropic()
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=max_tokens,
            system=system_prompt,
            messages=[{"role": "user", "content": user_message}]
        )
        text = response.content[0].text
        # Strip markdown code fences if present
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)

    else:
        raise ValueError(f"Unknown provider: {provider}")


def generate_structure_plan(deck_state_json: str, user_instruction: str,
                            provider: str = None) -> dict:
    """
    Pass 1: Generate structural plan + content manifest.
    Small output (~200-400 tokens). Returns in ~3 seconds.
    """
    if provider is None:
        provider = LLM_PROVIDER
    from prompts import PLAN_PROMPT
    prompt = PLAN_PROMPT.format(deck_state=deck_state_json)
    return _call_llm(prompt, user_instruction, provider, max_tokens=PLAN_MAX_TOKENS)


def generate_content(plan_json: str, deck_state_json: str,
                     provider: str = None) -> dict:
    """
    Pass 2: Generate all text content for the approved plan.
    Large output (~1-5K tokens). Returns in ~8-30 seconds.
    """
    if provider is None:
        provider = LLM_PROVIDER
    from prompts import CONTENT_PROMPT
    prompt = CONTENT_PROMPT.format(
        plan=plan_json,
        deck_state=deck_state_json
    )
    return _call_llm(prompt, "Generate all content now.", provider,
                     max_tokens=CONTENT_MAX_TOKENS)


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
                     max_tokens=VALIDATION_MAX_TOKENS)
