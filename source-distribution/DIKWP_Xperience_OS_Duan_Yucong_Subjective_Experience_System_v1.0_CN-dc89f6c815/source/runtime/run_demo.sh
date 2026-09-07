#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
python "$ROOT/runtime/xperience_runtime.py" demo \
  --scenarios "$ROOT/runtime/demo_scenarios.json" \
  --out "$ROOT/examples/demo_results.json"
