#!/bin/bash
# PreToolUse hook: Block cross-layer imports.
# tools.py/executor.py must never import LLM SDKs.
# llm.py/prompts.py must never import Aspose.
# Exit 0 = allow, exit 2 = block with message on stderr.

INPUT=$(cat)

FILE_PATH=$(echo "$INPUT" | python3 -c "
import sys, json
data = json.load(sys.stdin)
# Edit tool has file_path at top level; Write tool also has file_path
print(data.get('file_path', ''))
" 2>/dev/null)

# Only check relevant files
BASENAME=$(basename "$FILE_PATH" 2>/dev/null)
case "$BASENAME" in
    tools.py|executor.py)
        # Check new content for LLM SDK imports
        CONTENT=$(echo "$INPUT" | python3 -c "
import sys, json
data = json.load(sys.stdin)
# Edit: new_string; Write: content
print(data.get('new_string', data.get('content', '')))
" 2>/dev/null)
        if echo "$CONTENT" | grep -qE '(from (openai|anthropic)|import (openai|anthropic))'; then
            echo "LAYER VIOLATION: $BASENAME must not import LLM SDKs (openai/anthropic). LLM calls belong in llm.py only." >&2
            exit 2
        fi
        ;;
    llm.py|prompts.py)
        CONTENT=$(echo "$INPUT" | python3 -c "
import sys, json
data = json.load(sys.stdin)
print(data.get('new_string', data.get('content', '')))
" 2>/dev/null)
        if echo "$CONTENT" | grep -qE '(import aspose|from aspose)'; then
            echo "LAYER VIOLATION: $BASENAME must not import Aspose. Aspose operations belong in tools.py only." >&2
            exit 2
        fi
        ;;
esac

exit 0
