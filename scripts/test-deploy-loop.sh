#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$REPO_DIR"

DEPLOY_ID="${TEST_DEPLOYMENT_ID:-}"
if [[ -z "$DEPLOY_ID" && -f ".clasp-test-deploy-id" ]]; then
  DEPLOY_ID="$(tr -d '[:space:]' < .clasp-test-deploy-id)"
fi

if [[ -z "$DEPLOY_ID" ]]; then
  echo "Missing test deployment ID."
  echo "Set TEST_DEPLOYMENT_ID or create .clasp-test-deploy-id in repo root."
  exit 1
fi

DESC="${1:-test loop update $(date '+%Y-%m-%d %H:%M:%S')}"

echo "Using deployment ID: $DEPLOY_ID"
echo "Description: $DESC"

clasp push -f
clasp deploy -i "$DEPLOY_ID" -d "$DESC"
