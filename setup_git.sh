#!/bin/bash
# Configuration initiale du depot (a lancer UNE SEULE FOIS).
# Prerequis : avoir cree le depot PUBLIC "bilan-odf-app" sur github.com (compte bryanfaruch-alt).
set -e
PUB="$(cd "$(dirname "$0")" && pwd)"
cd "$PUB"
git config --global credential.helper osxkeychain || true
[ -z "$(git config --global user.name)" ]  && git config --global user.name "Bryan Faruch" || true
[ -z "$(git config --global user.email)" ] && git config --global user.email "bryan.faruch@gmail.com" || true
if [ ! -d .git ]; then git init -q; git branch -M main; fi
git add -A
git commit -q -m "Bilan ODF v1 - code + mise a jour auto" \
  -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>" \
  -m "Claude-Session: https://claude.ai/code/session_018tCmFHJGZsKfVy8pvHVDcU" || echo "(deja commite)"
git remote remove origin 2>/dev/null || true
git remote add origin https://github.com/bryanfaruch-alt/bilan-odf-app.git
echo ">>> Envoi vers GitHub (identifiant + token demandes une seule fois)..."
git push -u origin main
echo ">>> OK - depot en ligne : https://github.com/bryanfaruch-alt/bilan-odf-app"
