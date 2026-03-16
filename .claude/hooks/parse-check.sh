#!/bin/bash
# PostToolUse hook: Run py_compile after any Python file edit.
# Exit 0 = pass, non-zero = feedback with error on stderr.

INPUT=$(cat)

FILE_PATH=$(echo "$INPUT" | python3 -c "
import sys, json
data = json.load(sys.stdin)
print(data.get('file_path', ''))
" 2>/dev/null)

# Only check Python files
case "$FILE_PATH" in
    *.py)
        if [ -f "$FILE_PATH" ]; then
            ERROR=$(python3 -m py_compile "$FILE_PATH" 2>&1)
            if [ $? -ne 0 ]; then
                echo "SYNTAX ERROR in $FILE_PATH:" >&2
                echo "$ERROR" >&2
                exit 1
            fi
        fi
        ;;
esac

exit 0
