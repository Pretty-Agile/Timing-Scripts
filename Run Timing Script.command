#!/bin/bash

# Move to the folder where this .command file lives (so Python can find the script)
cd "$(dirname "$0")"

python3 "SAFe Timing Sheet Script.py"

echo ""
echo "Done! Press any key to close this window."
read -n 1
