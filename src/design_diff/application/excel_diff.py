"""Excel 設計書の新旧比較（設定に従って抽出し、ドメインの比較を呼ぶ）."""

from __future__ import annotations

import re
from fnmatch import fnmatch
from pathlib import Path
from typing import Any

from ..domain import excel_diff
from ..domain.model import Change, DiffResult
from ..infrastructure import excel_reader

KINDS = {"label", "block", "table"}


def validate_config(conf: dict[str, Any], where: str) -> None:
    if not isinstance(conf.get("sheets"), list) or not conf["sheets"]:
        raise ValueError(f"{where}: sheets（シートごとの比較設定）が必要です")
    for i, s in enumerate(conf["sheets"]):
        for j, item in enumerate(s.get("items") or []):
            if item.get("kind") not in KINDS:
                raise ValueError(f"{where}: sheets[{i}].items[{j}].kind は label / block / table です")
            if not item.get("name"):
                raise ValueError(f"{where}: sheets[{i}].items[{j}] に name がありません")
            if item["kind"] in ("label", "block") and not item.get("label"):
                raise ValueError(f"{where}: sheets[{i}].items[{j}] に label がありません")
            if item["kind"] == "table" and not (item.get("header") and item.get("key")):
                raise ValueError(f"{where}: sheets[{i}].items[{j}] に header と key が必要です")
        if s.get("pair_by"):
            try:
                if re.compile(s["pair_by"]).groups < 1:
                    raise ValueError(f"{where}: sheets[{i}].pair_by にはグループ ( ) が 1 つ必要です")
            except re.error as e:
                raise ValueError(f"{where}: sheets[{i}].pair_by の正規表現が不正です: {e}") from None


def run(conf: dict[str, Any], before: Path, after: Path) -> DiffResult:
    b_book, a_book = excel_reader.open_book(before), excel_reader.open_book(after)
    result = DiffResult(f"Excel 差分: {after.name}", f"変更前 {before.name}", f"変更後 {after.name}")
    for sheet_conf in conf["sheets"]:
        pattern = sheet_conf.get("name", "*")
        b_keys = _keyed([n for n in b_book.sheetnames if fnmatch(n, pattern)], sheet_conf.get("pair_by"))
        a_keys = _keyed([n for n in a_book.sheetnames if fnmatch(n, pattern)], sheet_conf.get("pair_by"))
        result.changes += [Change("sheet_removed", n) for k, n in b_keys.items() if k not in a_keys]
        result.changes += [Change("sheet_added", n) for k, n in a_keys.items() if k not in b_keys]
        for key in [k for k in b_keys if k in a_keys]:
            b_name, a_name = b_keys[key], a_keys[key]
            if b_name != a_name:
                result.changes.append(Change("sheet_renamed", key, "シート名", b_name, a_name))
            result.changes += _compare_sheet(a_name, sheet_conf.get("items") or [], b_book[b_name], a_book[a_name])
    return result


def _keyed(names: list[str], pair_by: str | None) -> dict[str, str]:
    """シートの対応付けのキー → シート名. pair_by（正規表現の 1 つ目のグループ）があればそれで、無ければシート名で."""
    if not pair_by:
        return {n: n for n in names}
    out: dict[str, str] = {}
    for n in names:
        m = re.search(pair_by, n)
        out[m.group(1) if m else n] = n
    return out


def _compare_sheet(sheet: str, items: list[dict], b_ws, a_ws) -> list[Change]:
    out: list[Change] = []
    for item in items:
        loc = f"{sheet} / {item['name']}"
        kind = item["kind"]
        if kind == "label":
            off = tuple(item.get("offset", [0, 1]))
            out += excel_diff.compare_value(
                sheet, item["name"], excel_reader.read_label(b_ws, item["label"], off), excel_reader.read_label(a_ws, item["label"], off)
            )
        elif kind == "block":
            off = tuple(item.get("offset", [1, 0]))
            out += excel_diff.compare_value(
                sheet, item["name"], excel_reader.read_block(b_ws, item["label"], off), excel_reader.read_block(a_ws, item["label"], off)
            )
        else:
            bt = excel_reader.read_table(b_ws, item["header"], item.get("title"))
            at = excel_reader.read_table(a_ws, item["header"], item.get("title"))
            if bt is None or at is None:
                out.append(Change("table_not_found", loc, "", "あり" if bt else "なし", "あり" if at else "なし"))
                continue
            out += excel_diff.compare_table(loc, list(item["key"]), bt, at)
    return out
