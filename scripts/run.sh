#!/usr/bin/env bash
cd "$(dirname "$0")/.."
exec python -m uvicorn webapp.app:app --host 127.0.0.1 --port 8000
