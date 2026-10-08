"""新旧比較の設定の表の key は列名の配列. 文字列（1 文字ずつの列名になる）などは設定の誤り（終了コード 2）."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import pytest
from devtools_common.cli import run_main
from openpyxl import Workbook

from design_diff.application import excel_diff
from design_diff.cli import main as cli


def _conf(key) -> dict:
    return {"sheets": [{"name": "*", "items": [{"kind": "table", "name": "入力項目", "header": "項目名", "key": key}]}]}


@pytest.mark.parametrize("key", ["項目名", 1, {"a": 1}, ["項目名", ""], ["項目名", 2], [None]])
def test_key_must_be_list_of_column_names(key):
    with pytest.raises(ValueError, match="key"):
        excel_diff.validate_config(_conf(key), "api.yaml")


def test_list_key_is_accepted():
    excel_diff.validate_config(_conf(["項目名", "位置"]), "api.yaml")


def test_cli_string_key_is_usage_error(tmp_path: Path):
    (tmp_path / "x.yaml").write_text(
        "sheets:\n  - name: '*'\n    items:\n      - {kind: table, name: 入力項目, header: 項目名, key: 項目名}\n", encoding="utf-8"
    )
    for name in ("b.xlsx", "a.xlsx"):
        Workbook().save(tmp_path / name)
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(
            cli.main,
            ["excel", "--config-dir", str(tmp_path), "--type", "x"]
            + ["--before", str(tmp_path / "b.xlsx"), "--after", str(tmp_path / "a.xlsx")],
        )
    assert rc == 2
    assert "key は列名の配列" in err.getvalue()
