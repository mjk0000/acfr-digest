#!/usr/bin/env bash
# Run parse_cafr.py on every PDF in FY25_City_ACFRs_Set2/
# Output: FY25_City_ACFRs_Set2/set2_output.xlsx
# Usage: bash run_set2.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
INPUT_DIR="$SCRIPT_DIR/FY25_City_ACFRs_Set2"
OUTPUT="$INPUT_DIR/set2_output.xlsx"

cd "$INPUT_DIR"
/usr/bin/python3 "$SCRIPT_DIR/parse_cafr.py" *.pdf --output "$OUTPUT"
