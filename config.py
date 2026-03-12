"""
Configuration for the Surgical Slide Engine.

Centralizes API keys, model names, provider selection, and default parameters.
Environment variables take precedence over defaults.
"""

import os
from pathlib import Path

# Load .env file if present
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    for line in _env_path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip())

# --- Aspose License ---
import aspose.slides as slides
_lic_path = Path(__file__).parent / "Aspose Temporary License.lic"
if _lic_path.exists():
    _license = slides.License()
    _license.set_license(str(_lic_path))

# --- LLM Provider ---
# "openai" or "anthropic"
LLM_PROVIDER = os.getenv("SSE_LLM_PROVIDER", "openai")

# --- Model Names ---
OPENAI_MODEL = os.getenv("SSE_OPENAI_MODEL", "gpt-4o-mini")
ANTHROPIC_MODEL = os.getenv("SSE_ANTHROPIC_MODEL", "claude-haiku-4-5-20251001")

# --- API Keys (read from environment) ---
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# --- Pipeline Defaults ---
MAX_LLM_RETRIES = 2
PLAN_MAX_TOKENS = 4000
CONTENT_MAX_TOKENS = 16000
VALIDATION_MAX_TOKENS = 2000

# --- Char Limit ---
CHAR_LIMIT_SAFETY_MARGIN = 0.85
DEFAULT_FONT_SIZE_PT = 12
DEFAULT_LINE_SPACING = 1.2

# --- File Paths ---
DEFAULT_OUTPUT_DIR = os.getenv("SSE_OUTPUT_DIR", "output")
TEMP_DIR = os.getenv("SSE_TEMP_DIR", "temp")

# --- Placeholder Detection ---
PLACEHOLDER_PATTERNS = [
    "xxxx", "XXXX", "lorem", "ipsum", "placeholder", "[placeholder]",
    "TBD", "TODO", "[insert", "sample text", "click to add",
    "type here", "{title}", "{subtitle}", "{body}"
]
