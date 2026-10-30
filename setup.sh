#!/usr/bin/env bash
# セットアップ（macOS / Linux / Git Bash）: .venv を作って design-diff を入れる
set -euo pipefail
cd "$(dirname "$0")"
# Git Bash（Windows）には python3 が無いことがあり、venv の中は .venv/Scripts になる
PY=python3; command -v python3 >/dev/null 2>&1 || PY=python
"$PY" -c 'import sys; sys.exit(sys.version_info < (3, 10))' || { echo "Python 3.10 以上が要ります（$("$PY" -V 2>&1)）" >&2; exit 1; }
"$PY" -m venv .venv
BIN=.venv/bin; [ -d .venv/Scripts ] && BIN=.venv/Scripts
"$BIN/python" -m pip install -q --upgrade pip || true
# 依存の版は constraints.txt でそろえる（CI と同じ. deps-lock が作る）
C=""; [ -f constraints.txt ] && C="-c constraints.txt"
# shellcheck disable=SC2086
"$BIN/python" -m pip install -q $C -e ".[dev]"
echo "完了: source $BIN/activate のあと design-diff --help"
