#!/bin/bash
# PostToolUse hook: Check for unescaped braces in prompts.py.
# .format() requires {{ and }} for literal braces. A single { that isn't
# a known placeholder like {deck_state} is likely a bug.
# Exit 0 = pass, non-zero = feedback with warning on stderr.

INPUT=$(cat)

FILE_PATH=$(echo "$INPUT" | python3 -c "
import sys, json
data = json.load(sys.stdin)
print(data.get('file_path', ''))
" 2>/dev/null)

BASENAME=$(basename "$FILE_PATH" 2>/dev/null)

# Only check prompts.py
if [ "$BASENAME" != "prompts.py" ]; then
    exit 0
fi

if [ ! -f "$FILE_PATH" ]; then
    exit 0
fi

# Check for single braces that aren't doubled or known placeholders
ISSUES=$(python3 -c "
import re, sys

KNOWN_PLACEHOLDERS = {'deck_state', 'plan', 'source_json', 'generated_text'}

with open('$FILE_PATH') as f:
    content = f.read()

# Find all { that are not {{ and not in known placeholders
# Strategy: find single { by removing all {{ first, then checking what remains
collapsed = content.replace('{{', '').replace('}}', '')
issues = []
for i, ch in enumerate(collapsed):
    if ch == '{':
        # Extract placeholder name
        end = collapsed.find('}', i)
        if end == -1:
            issues.append(f'Line ~{content[:i].count(chr(10))+1}: unmatched opening brace')
            continue
        name = collapsed[i+1:end].strip()
        if name not in KNOWN_PLACEHOLDERS:
            issues.append(f'Suspicious single brace: {{{name}}} — not a known placeholder ({name})')

for issue in issues:
    print(issue)
" 2>/dev/null)

if [ -n "$ISSUES" ]; then
    echo "BRACE WARNING in prompts.py — possible unescaped braces for .format():" >&2
    echo "$ISSUES" >&2
    exit 1
fi

exit 0
