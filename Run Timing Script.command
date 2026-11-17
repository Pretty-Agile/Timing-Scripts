#!/bin/bash

# Move to the folder where this .command file lives (so Python can find the script)
cd "$(dirname "$0")"

# Use the first Python that has the libraries the script needs (python-pptx and openpyxl).
# "python3" alone is not reliable: a newly installed Python can jump the queue without them.
PYTHON=""
for candidate in \
    python3 \
    /usr/local/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
    /opt/homebrew/bin/python3 \
    /usr/bin/python3
do
    if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c "import pptx, openpyxl" >/dev/null 2>&1; then
        PYTHON="$candidate"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "No Python on this Mac has the libraries this script needs."
    echo "Install them by pasting this line into Terminal, then run this again:"
    echo ""
    echo "    python3 -m pip install python-pptx openpyxl"
else
    "$PYTHON" "SAFe Timing Sheet Script.py"
fi

echo ""
echo "Done! Press any key to close this window."
read -n 1
