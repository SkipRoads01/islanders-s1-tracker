#!/usr/bin/env bash
# Rebuild index.html + version.txt from the workbook and publish.
# Usage: site/deploy.sh ["commit message"]
set -euo pipefail
cd "$(dirname "$0")/.."
python3 site/build_site.py
build="$(cat version.txt)"
grep -q "name=\"build\" content=\"$build\"" index.html || { echo "build id mismatch between index.html and version.txt"; exit 1; }
# the page must be pure ASCII outside base64 image data
LC_ALL=C grep -nP '[^\x00-\x7F]' index.html && { echo "non-ASCII bytes in index.html"; exit 1; }
git add index.html version.txt "S1 NY Islanders.xlsx" site CLAUDE.md
git commit -m "${1:-Publish site update} ($build)" || echo "nothing to commit"
git push -u origin "$(git rev-parse --abbrev-ref HEAD)"
