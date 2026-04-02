#!/bin/bash
# Deploy docx_to_qti.html to iwant2study.org and push to GitHub

set -e

REMOTE_URL="ftp://iwant2study.org/public_html/lookangejss/QTIlowJunHua/docx_to_qti.html"
LOCAL_FILE="$(dirname "$0")/docx_to_qti.html"

echo "▶ Uploading to server..."
curl -T "$LOCAL_FILE" "$REMOTE_URL" --netrc --ftp-create-dirs --silent --show-error
echo "✓ Server updated: https://iwant2study.org/lookangejss/QTIlowJunHua/docx_to_qti.html"

echo "▶ Pushing to GitHub..."
cd "$(dirname "$0")"
git push
echo "✓ GitHub updated: https://github.com/lookang/QTIconvertertool"
