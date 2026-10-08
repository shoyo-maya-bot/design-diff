"""リポジトリの設定の確認: セットアップスクリプトの constraints、CI のカバレッジ下限、AGENTS.md の記載."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_setup_scripts_use_constraints():
    sh = (ROOT / "setup.sh").read_text(encoding="utf-8")
    assert '[ -f constraints.txt ] && C="-c constraints.txt"' in sh
    assert all("$C" in ln for ln in sh.splitlines() if "pip install" in ln and "-e" in ln)
    raw = (ROOT / "setup.cmd").read_bytes()
    assert b"\r\n" in raw
    cmd = raw.decode("cp932")
    assert "if exist constraints.txt set C=-c constraints.txt" in cmd
    assert all("%C%" in ln for ln in cmd.splitlines() if "pip install" in ln and "-e" in ln)


def test_ci_measures_coverage_with_floor():
    import sys

    if sys.version_info >= (3, 11):
        import tomllib
    else:  # pragma: no cover
        import tomli as tomllib
    conf = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert any(d.startswith("pytest-cov") for d in conf["project"]["optional-dependencies"]["dev"])
    assert conf["tool"]["coverage"]["report"]["fail_under"] >= 90
    ci = (ROOT / ".github/workflows/ci.yml").read_text(encoding="utf-8")
    assert "pytest -q --cov" in ci


def test_agents_md_has_push_check_and_no_weakening_promise():
    text = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    assert "push" in text and "pytest -q --cov" in text
    assert "チェックを甘くしない" in text
    assert "AGENTS.md" in (ROOT / "README.md").read_text(encoding="utf-8")
