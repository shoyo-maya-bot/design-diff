"""Excel 設計書の新旧比較（設定に従って抽出し、ドメインの比較を呼ぶ）."""

from __future__ import annotations

import re
from fnmatch import fnmatchcase
from pathlib import Path
from typing import Any

from ..domain import excel_diff
from ..domain.model import Change, DiffResult
from ..infrastructure import excel_reader

KINDS = {"label", "block", "table"}
OUTSIDE_MODES = ("summary", "all", "off")  # 設定に無いセルの変更: 件数と場所 / 1 セルずつ / 比べない
OUTSIDE_SHOWN = 10  # summary で場所を並べる数


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
            if item["kind"] == "table" and not _is_column_list(item["key"]):
                raise ValueError(
                    f"{where}: sheets[{i}].items[{j}].key は列名の配列です（例: key: [項目名]）。"
                    f"文字列 1 つ・数値・空の列名は書けません: {item['key']!r}"
                )
            for opt, hint in (("header_rows", "（省略で自動判定）"), ("end_blank_rows", f"（省略で {excel_reader.END_BLANK_ROWS}）")):
                v = item.get(opt)
                if v is not None and (not isinstance(v, int) or isinstance(v, bool) or v < 1):
                    raise ValueError(f"{where}: sheets[{i}].items[{j}].{opt} は 1 以上の整数です{hint}")
        if s.get("pair_by"):
            try:
                if re.compile(s["pair_by"]).groups < 1:
                    raise ValueError(f"{where}: sheets[{i}].pair_by にはグループ ( ) が 1 つ必要です")
            except re.error as e:
                raise ValueError(f"{where}: sheets[{i}].pair_by の正規表現が不正です: {e}") from None


def _is_column_list(v: Any) -> bool:
    """key は空でない列名（文字列）の配列. 文字列のままだと 1 文字ずつの列名になるので誤りとする."""
    return isinstance(v, list) and bool(v) and all(isinstance(c, str) and c.strip() for c in v)


def run(conf: dict[str, Any], before: Path, after: Path, *, outside: str = "summary") -> DiffResult:
    """設定に従って新旧を比べる.

    outside: 設定に無いセル（承認者欄など）の変更の出し方. summary = シートごとに件数と場所（既定）、
    all = 1 セルずつ、off = 比べない.
    """
    if outside not in OUTSIDE_MODES:
        raise ValueError(f"outside は {' / '.join(OUTSIDE_MODES)} のどれかです: {outside}")
    (b_book, b_warn), (a_book, a_warn) = excel_reader.load_book(before), excel_reader.load_book(after)
    result = DiffResult(f"Excel 差分: {after.name}", f"変更前 {before.name}", f"変更後 {after.name}")
    result.warnings = b_warn + a_warn
    configured_b: set[str] = set()
    configured_a: set[str] = set()
    for sheet_conf in conf["sheets"]:
        pattern = sheet_conf.get("name", "*")
        b_keys = _keyed([n for n in b_book.sheetnames if fnmatchcase(n, pattern)], sheet_conf.get("pair_by"), before.name, result)
        a_keys = _keyed([n for n in a_book.sheetnames if fnmatchcase(n, pattern)], sheet_conf.get("pair_by"), after.name, result)
        configured_b |= set(b_keys.values())
        configured_a |= set(a_keys.values())
        result.changes += [Change("sheet_removed", n) for k, n in b_keys.items() if k not in a_keys]
        result.changes += [Change("sheet_added", n) for k, n in a_keys.items() if k not in b_keys]
        for key in [k for k in b_keys if k in a_keys]:
            b_name, a_name = b_keys[key], a_keys[key]
            if b_name != a_name:
                result.changes.append(Change("sheet_renamed", key, "シート名", b_name, a_name))
            items = sheet_conf.get("items") or []
            b_ctx, a_ctx = _context(items), _context(items)
            result.changes += _compare_sheet(a_name, items, b_book[b_name], a_book[a_name], b_ctx, a_ctx)
            _collect(result, b_ctx, a_ctx)
            if outside != "off":
                result.changes += _compare_outside(
                    a_name,
                    excel_reader.outside_rows(b_book[b_name], b_ctx.covered),
                    excel_reader.outside_rows(a_book[a_name], a_ctx.covered),
                    outside,
                )
    if outside != "off":
        # 設定のどのシートにも当たらないシート（表紙など）は、シート名で対応付けて全部のセルを比べる
        rest_b = [n for n in b_book.sheetnames if n not in configured_b]
        rest_a = [n for n in a_book.sheetnames if n not in configured_a]
        for name in rest_b + [n for n in rest_a if n not in rest_b]:
            b_rows = excel_reader.outside_rows(b_book[name], set()) if name in rest_b else []
            a_rows = excel_reader.outside_rows(a_book[name], set()) if name in rest_a else []
            result.changes += _compare_outside(name, b_rows, a_rows, outside)
    result.alerts += excel_reader.book_alerts(b_book, before.name) + excel_reader.book_alerts(a_book, after.name)
    return result


def _context(items: list[dict]) -> excel_reader.ReadContext:
    """1 シート分の読み取りの文脈. 設定にある全項目のラベル・表題・見出しを「次の節の始まり」とみなす."""
    labels = {excel_reader.norm_label(item[k]) for item in items for k in ("label", "title", "header") if item.get(k)}
    return excel_reader.ReadContext(stop_labels=labels)


def _collect(result: DiffResult, *ctxs: excel_reader.ReadContext) -> None:
    for ctx in ctxs:
        result.alerts += [a for a in ctx.alerts if a not in result.alerts]
        result.notes += [n for n in ctx.notes if n not in result.notes]


def _compare_outside(sheet: str, before: list, after: list, mode: str) -> list[Change]:
    diffs = excel_diff.compare_outside(before, after)
    if not diffs:
        return []
    if mode == "all":
        return [Change("outside_cell_changed", sheet, ref, b, a) for ref, b, a in diffs]
    refs = ", ".join(ref for ref, _, _ in diffs[:OUTSIDE_SHOWN]) + (" ほか" if len(diffs) > OUTSIDE_SHOWN else "")
    return [Change("outside_changed", sheet, "比較対象外のセル", "", f"{len(diffs)} 件: {refs}（--outside all で 1 セルずつ）")]


def _keyed(names: list[str], pair_by: str | None, book: str = "", result: DiffResult | None = None) -> dict[str, str]:
    """シートの対応付けのキー → シート名. pair_by（正規表現の 1 つ目のグループ）があればそれで、無ければシート名で.

    pair_by のキーが同じシートが 2 枚以上あると、対応付けられるのは 1 枚だけ（最初のシート）. 残りは比べずに
    警告 duplicate_pair_key を出す（以前は黙って後のシートで上書きしていた）.
    """
    if not pair_by:
        return {n: n for n in names}
    out: dict[str, str] = {}
    dups: dict[str, list[str]] = {}
    for n in names:
        m = re.search(pair_by, n)
        key = m.group(1) if m else n
        if key in out:
            dups.setdefault(key, [out[key]]).append(n)
            continue
        out[key] = n
    if result is not None:
        for key, sheets in dups.items():
            listed = "・".join(f"「{n}」" for n in sheets)
            result.alerts.append(
                (
                    "duplicate_pair_key",
                    f"{book}: pair_by のキー「{key}」のシートが {len(sheets)} 枚あります（{listed}）。"
                    f"最初の「{sheets[0]}」だけを対応付けて比べ、残りは比べていません。pair_by の正規表現を見直してください",
                )
            )
    return out


BLANK_KEY_SHOWN = 10  # 警告で並べる行番号の数


def _warn_blank_key(side: str, sheet: str, item: dict, table, row_numbers: list[int], ctx: excel_reader.ReadContext) -> None:
    """キー列が空の行（対応付けられないので比べていない行）があれば、件数と行番号で警告する（blank_key）."""
    key = list(item["key"])
    if any(k not in table[0] for k in key):
        return  # キー列が無い（key_column_missing で知らせる）
    idx = excel_diff.blank_key_rows(key, table[1])
    if not idx:
        return
    nums = [row_numbers[i] for i in idx if i < len(row_numbers)]
    listed = "、".join(map(str, nums[:BLANK_KEY_SHOWN])) + (" ほか" if len(nums) > BLANK_KEY_SHOWN else "")
    ctx.alerts.append(
        (
            "blank_key",
            f"{side} {sheet}: 表「{item['name']}」にキー列（{'・'.join(key)}）が空の行が {len(idx)} 行あります（行 {listed}）。"
            "対応付けられないため比べていません。キーを埋めるか、設定の key を見直してください",
        )
    )


def _compare_sheet(
    sheet: str, items: list[dict], b_ws, a_ws, b_ctx: excel_reader.ReadContext, a_ctx: excel_reader.ReadContext
) -> list[Change]:
    out: list[Change] = []
    for item in items:
        loc = f"{sheet} / {item['name']}"
        kind = item["kind"]
        for ctx in (b_ctx, a_ctx):
            ctx.end_blank_rows = item.get("end_blank_rows") or excel_reader.END_BLANK_ROWS
        if kind == "label":
            off = tuple(item.get("offset", [0, 1]))
            out += excel_diff.compare_value(
                sheet,
                item["name"],
                excel_reader.read_label(b_ws, item["label"], off, b_ctx),
                excel_reader.read_label(a_ws, item["label"], off, a_ctx),
            )
        elif kind == "block":
            off = tuple(item.get("offset", [1, 0]))
            out += excel_diff.compare_value(
                sheet,
                item["name"],
                excel_reader.read_block(b_ws, item["label"], off, b_ctx),
                excel_reader.read_block(a_ws, item["label"], off, a_ctx),
            )
        else:
            hr = item.get("header_rows")
            bt = excel_reader.read_table(b_ws, item["header"], item.get("title"), hr, b_ctx)
            at = excel_reader.read_table(a_ws, item["header"], item.get("title"), hr, a_ctx)
            if bt is None or at is None:
                out.append(Change("table_not_found", loc, "", "あり" if bt else "なし", "あり" if at else "なし"))
                continue
            b_rows, a_rows = list(b_ctx.table_rows), list(a_ctx.table_rows)
            out += excel_diff.compare_table(loc, list(item["key"]), bt, at)
            for side, ws, table, nums, ctx in (("変更前", b_ws, bt, b_rows, b_ctx), ("変更後", a_ws, at, a_rows, a_ctx)):
                _warn_blank_key(side, ws.title, item, table, nums, ctx)
    return out
