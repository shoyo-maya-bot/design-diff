"""第 3 回の指摘の回帰テスト: キー列が空の行、OpenAPI の壊れた入力、同じ API ID のシート."""

from __future__ import annotations

import contextlib
import io
from pathlib import Path

import pytest
from devtools_common.cli import run_main
from openpyxl import Workbook
from test_robustness import TABLE_CONF, book, table_cells

from design_diff.application import excel_diff, spec_diff
from design_diff.cli import main as cli
from design_diff.domain.excel_diff import blank_key_rows

ROOT = Path(__file__).resolve().parent.parent
SPEC_CONF = {
    "sheets": "API*",
    "api_id": {"label": "API ID"},
    "method": {"label": "メソッド"},
    "path": {"label": "パス"},
}


def run(*args: str) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = run_main(cli.main, list(args))
    return rc, out.getvalue() + err.getvalue()


# --- 1. キー列が空の行 --------------------------------------------------------------------------


def test_blank_key_rows():
    rows = [{"項番": "1"}, {"項番": ""}, {"項番": " "}, {"項番": "2"}]
    assert blank_key_rows(["項番"], rows) == [1, 2]
    assert blank_key_rows(["a", "b"], [{"a": "", "b": "x"}, {"a": "", "b": ""}]) == [1]


def test_blank_key_rows_are_warned_with_count_and_row_numbers(tmp_path: Path):
    header = ["項番", "名前", "型"]
    before = book(tmp_path / "b.xlsx", table_cells(header, [[1, "a", "数値"], [None, "b", "文字列"], [2, "c", "数値"]]))
    after = book(
        tmp_path / "a.xlsx", table_cells(header, [[1, "a", "数値"], [None, "b2", "文字列"], [None, "x", "日付"], [2, "c", "数値"]])
    )
    result = excel_diff.run(TABLE_CONF, before, after)
    alerts = [w for rule, w in result.alerts if rule == "blank_key"]
    assert len(alerts) == 2, result.alerts
    assert "変更前" in alerts[0] and "1 行" in alerts[0] and "行 3）" in alerts[0]
    assert "変更後" in alerts[1] and "2 行" in alerts[1] and "行 3、4）" in alerts[1]
    # 空キーでない行に差分が無ければ、差分は 0 件（空キーの行は比べていないことを警告で知らせる）
    assert [c for c in result.changes if c.category not in ("outside_changed",)] == []


def test_no_blank_key_no_alert(tmp_path: Path):
    header = ["項番", "名前"]
    b = book(tmp_path / "b.xlsx", table_cells(header, [[1, "a"]]))
    assert not [r for r, _ in excel_diff.run(TABLE_CONF, b, b).alerts if r == "blank_key"]


# --- 2. OpenAPI の壊れた入力は終了コード 2 --------------------------------------------------------


@pytest.mark.parametrize(
    "content",
    [
        "openapi: 3.0.0\npaths: {\n  /x: [\n",  # YAML の構文の誤り
        "openapi: 3.0.0\npaths:\n  /x: [1, 2]\n",  # パスの中身が辞書でない
        "openapi: 3.0.0\npaths:\n  /x:\n    get: 1\n",  # 操作が辞書でない
        "openapi: 3.0.0\npaths:\n  /x:\n    get:\n      parameters: [1]\n",  # パラメーターが辞書でない
        "openapi: 3.0.0\npaths: [1]\n",  # paths が辞書でない
        "openapi: 3.0.0\npaths:\n  /x:\n    get:\n      responses: {'200': {content: {application/json: {schema: {$ref: '#/nope'}}}}}\n",
    ],
)
def test_broken_openapi_is_input_error(tmp_path: Path, content: str):
    oas = tmp_path / "openapi.yaml"
    oas.write_text(content, encoding="utf-8")
    rc, out = run("spec", "--excel", str(ROOT / "samples/spec/API定義書.xlsx"), "--oas", str(oas))
    assert rc == 2, out
    assert "Traceback" not in out


def test_broken_openapi_raises_value_error(tmp_path: Path):
    from design_diff.infrastructure import oas_reader

    oas = tmp_path / "o.yaml"
    oas.write_text("paths: {\n", encoding="utf-8")
    with pytest.raises(ValueError, match="OpenAPI"):
        oas_reader.read_with_notes(oas)


# --- 3. 同じ API ID のシート ------------------------------------------------------------------------


def test_duplicate_api_id_uses_first_sheet_and_warns(tmp_path: Path):
    wb = Workbook()
    wb.remove(wb.active)
    for title, method in (("API1", "GET"), ("API2", "POST")):
        ws = wb.create_sheet(title)
        ws["A1"], ws["B1"] = "API ID", "API0001"
        ws["A2"], ws["B2"] = "メソッド", method
        ws["A3"], ws["B3"] = "パス", "/x"
    path = tmp_path / "spec.xlsx"
    wb.save(path)
    oas = tmp_path / "o.yaml"
    oas.write_text("openapi: 3.0.0\npaths:\n  /x:\n    get:\n      operationId: API0001\n      responses: {'200': {}}\n", encoding="utf-8")
    result = spec_diff.run(SPEC_CONF, path, oas)
    alerts = [w for rule, w in result.alerts if rule == "duplicate_api_id"]
    assert len(alerts) == 1 and "「API1」" in alerts[0] and "「API2」" in alerts[0]
    assert not [c for c in result.changes if c.category == "method_mismatch"]  # 最初のシート（GET）を使う


# --- 5. セットアップの constraints と、入力 xlsx の展開の上限 ---------------------------------------------


def test_setup_scripts_use_constraints():
    sh = (ROOT / "setup.sh").read_text(encoding="utf-8")
    assert '[ -f constraints.txt ] && C="-c constraints.txt"' in sh
    assert all("$C" in ln for ln in sh.splitlines() if "pip install" in ln and "-e" in ln)
    raw = (ROOT / "setup.cmd").read_bytes()
    assert b"\r\n" in raw
    cmd = raw.decode("cp932")
    assert "if exist constraints.txt set C=-c constraints.txt" in cmd
    assert all("%C%" in ln for ln in cmd.splitlines() if "pip install" in ln and "-e" in ln)


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


# --- 4. CI でカバレッジを測り、下限を設ける ------------------------------------------------------------


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


# --- 6. AGENTS.md ---------------------------------------------------------------------------------


def test_agents_md_has_push_check_and_no_weakening_promise():
    text = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    assert "push" in text and "pytest -q --cov" in text
    assert "チェックを甘くしない" in text
    assert "AGENTS.md" in (ROOT / "README.md").read_text(encoding="utf-8")
