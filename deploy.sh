#!/usr/bin/env bash
# Push code to Apps Script and redeploy the existing web app deployment (URL stays the same).
set -euo pipefail
cd "$(dirname "$0")"
DEPLOYMENT_ID="AKfycbzcxx2rrzSYrH7UrrjWE1pIMsqfacHB5uD-YtEhIqHV9azgVNK8-KiG2r5bvyVbJ2DO3A"
MSG="${1:-update $(date '+%Y-%m-%d %H:%M')}"
clasp push --force
clasp deploy --deploymentId "$DEPLOYMENT_ID" --description "$MSG"
