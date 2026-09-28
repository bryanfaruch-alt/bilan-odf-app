#!/bin/bash
# Publie une nouvelle version : ./publish.sh <version> "<notes>"
set -e
APPDIR="/Users/bryanfaruch/Documents/Claude/Projects/PROJET BILAN ODF AUTO/BilanODF_App"
PUB="$(cd "$(dirname "$0")" && pwd)"
VER="$1"; shift; NOTES="$*"
[ -z "$VER" ] && { echo "usage: publish.sh <version> <notes>"; exit 1; }
for f in app.py run_native.py pipeline.py word_import.py backup.py mass_import.py appareils_ext.py bilan_pdf_rl.py requirements.txt; do cp "$APPDIR/$f" "$PUB/$f"; done
mkdir -p "$PUB/assets"; cp "$APPDIR/assets/chu_nice.png" "$PUB/assets/chu_nice.png"
# La version est definie dans app.py (source unique, lue aussi par run_native)
sed -i '' "s/^APP_VERSION = \"[^\"]*\"/APP_VERSION = \"$VER\"/" "$PUB/app.py"
python3 - "$VER" "$NOTES" > "$PUB/version.json" <<'PY'
import json,sys
print(json.dumps({"version":sys.argv[1],"notes":(sys.argv[2] if len(sys.argv)>2 else ""),
 "files":["app.py","run_native.py","pipeline.py","word_import.py","backup.py","mass_import.py","appareils_ext.py","bilan_pdf_rl.py","requirements.txt","assets/chu_nice.png"]},ensure_ascii=False,indent=1))
PY
cd "$PUB"; git add -A
git commit -m "Release v$VER: $NOTES" \
  -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>" \
  -m "Claude-Session: https://claude.ai/code/session_018tCmFHJGZsKfVy8pvHVDcU" || echo "(rien a committer)"
git push
echo "PUBLISHED v$VER"
