"""入力 xlsx（zip）の展開の上限: 圧縮率・部品の数・展開後の大きさを超えたら入力の誤り（終了コード 2）."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import pytest
from devtools_common.cli import run_main
from test_robustness import TABLE_CONF, book

from design_diff.cli import main as cli


def run(*args: str) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(cli.main, list(args))
    return rc, out.getvalue() + err.getvalue()


def _bomb(path: Path) -> Path:
    import zipfile

    book(path, {"A1": "x"})
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/media/pad.bin", b"\0" * (8 * 1024 * 1024))  # 8MB の 0（圧縮率が非常に高い）
    return path


def test_zip_bomb_xlsx_is_input_error(tmp_path: Path):
    bomb = _bomb(tmp_path / "bomb.xlsx")
    ok = book(tmp_path / "ok.xlsx", {"A1": "x"})
    conf_dir = tmp_path / "config"
    conf_dir.mkdir()
    import yaml

    (conf_dir / "t.yaml").write_text(yaml.safe_dump(TABLE_CONF), encoding="utf-8")
    rc, out = run("--config-dir", str(conf_dir), "excel", "--type", "t", "--before", str(ok), "--after", str(bomb))
    assert rc == 2, out
    assert "圧縮率" in out and "Traceback" not in out


def test_zip_limits_entries_and_total(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    from design_diff.infrastructure import excel_reader

    path = book(tmp_path / "a.xlsx", {"A1": "x"})
    excel_reader.check_zip_limits(path)  # 普通の xlsx は通る
    monkeypatch.setattr(excel_reader, "MAX_ENTRIES", 2)
    with pytest.raises(excel_reader.UnreadableBookError, match="部品の数"):
        excel_reader.check_zip_limits(path)
    monkeypatch.setattr(excel_reader, "MAX_ENTRIES", 10_000)
    monkeypatch.setattr(excel_reader, "MAX_UNZIPPED", 100)
    with pytest.raises(excel_reader.UnreadableBookError, match="展開後の大きさ"):
        excel_reader.check_zip_limits(path)
