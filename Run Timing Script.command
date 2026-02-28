#!/bin/bash

# Move to the folder where this .command file lives (so Python can find the script)
cd "$(dirname "$0")"

# Ask the user to pick their .pptx folder using a native Mac dialog
INPUT_FOLDER=$(osascript -e 'tell application "Finder" to set theFolder to choose folder with prompt "Select the folder containing your .pptx files"
return POSIX path of theFolder')

# If the user cancelled, exit quietly
if [ -z "$INPUT_FOLDER" ]; then
    echo "No folder selected. Exiting."
    exit 0
fi

echo "Running script on: $INPUT_FOLDER"
echo ""

python3 "SAFe Timing Sheet Script.py" --input-folder "$INPUT_FOLDER"

echo ""
echo "Done! Press any key to close this window."
read -n 1
