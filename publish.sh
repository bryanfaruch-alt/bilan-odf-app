#!/bin/bash
# Publie une nouvelle version : ./publish.sh <version> "<notes>"
# Les fichiers de code sont references par le COMMIT precis (ref) dans version.json
# -> URL immuable, jamais de contenu perime servi par le cache CDN de GitHub.
set -e
APPDIR="/Users/bryanfaruch/Documents/Claude/Projects/PROJET BILAN ODF AUTO/BilanODF_App"
PUB="$(cd "$(dirname "$0")" && pwd)"
VER="$1"; shift; NOTES="$*"
[ -z "$VER" ] && { echo "usage: publish.sh <version> <notes>"; exit 1; }
CO1="Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
CO2="Claude-Session: https://claude.ai/code/session_018tCmFHJGZsKfVy8pvHVDcU"
FILES="app.py run_native.py pipeline.py word_import.py backup.py mass_import.py appareils_ext.py bilan_pdf_rl.py requirements.txt photo_train_seed.jsonl"
for f in $FILES; do cp "$APPDIR/$f" "$PUB/$f"; done
mkdir -p "$PUB/assets"; cp "$APPDIR/assets/chu_nice.png" "$PUB/assets/chu_nice.png"
# La version est definie dans app.py (source unique, lue aussi par run_native)
sed -i '' "s/^APP_VERSION = \"[^\"]*\"/APP_VERSION = \"$VER\"/" "$PUB/app.py"
cd "$PUB"
# 1) commit du code (version.json provisoire, sans ref) -> SHA fige
python3 - "$VER" "$NOTES" "" > version.json <<'PY'
import json,sys
print(json.dumps({"version":sys.argv[1],"notes":(sys.argv[2] if len(sys.argv)>2 else ""),"ref":(sys.argv[3] if len(sys.argv)>3 else ""),
 "files":["app.py","run_native.py","pipeline.py","word_import.py","backup.py","mass_import.py","appareils_ext.py","bilan_pdf_rl.py","requirements.txt","photo_train_seed.jsonl","assets/chu_nice.png"]},ensure_ascii=False,indent=1))
PY
git add -A
git commit -m "Code v$VER: $NOTES" -m "$CO1" -m "$CO2" || echo "(rien a committer)"
SHA=$(git rev-parse HEAD)
# 2) version.json avec le ref = ce commit, puis push
python3 - "$VER" "$NOTES" "$SHA" > version.json <<'PY'
import json,sys
print(json.dumps({"version":sys.argv[1],"notes":(sys.argv[2] if len(sys.argv)>2 else ""),"ref":sys.argv[3],
 "files":["app.py","run_native.py","pipeline.py","word_import.py","backup.py","mass_import.py","appareils_ext.py","bilan_pdf_rl.py","requirements.txt","photo_train_seed.jsonl","assets/chu_nice.png"]},ensure_ascii=False,indent=1))
PY
git add version.json
git commit -m "Version v$VER -> $SHA" -m "$CO1" -m "$CO2" || echo "(version.json inchange)"
git push
echo "PUBLISHED v$VER (ref=$SHA)"
