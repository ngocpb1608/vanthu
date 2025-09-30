#!/usr/bin/env bash
set -euo pipefail
THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$THIS_DIR"
PY=python3; command -v python3 >/dev/null 2>&1 || PY=python
$PY -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
export SECRET_KEY="${SECRET_KEY:-something-strong}"
export DATABASE_URL="${DATABASE_URL:-sqlite:///data.db}"
python -m flask --app server run --debug --port 5050
