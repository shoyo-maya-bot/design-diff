@echo off
rem セットアップ（Windows）: .venv を作って design-diff を入れる
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% -m venv .venv || goto :err
.venv\Scripts\python -m pip install -q --upgrade pip
rem 依存の版は constraints.txt でそろえる（CI と同じ）
set C=
if exist constraints.txt set C=-c constraints.txt
.venv\Scripts\python -m pip install -q %C% -e ".[dev]" || goto :err
echo 完了: .venv\Scripts\activate のあと design-diff --help
pause
exit /b 0
:err
echo セットアップに失敗しました。Python 3.10 以上と Git が入っているか確認してください。
pause
exit /b 1
