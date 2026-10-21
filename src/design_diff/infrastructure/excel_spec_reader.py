"""Excel の API 定義書を、設定（config/spec-*.yaml）に従って正規化した ApiSpec にする.

設定の例:

    sheets: "API*"                       # API 1 本 = 1 シート（シート名のパターン）
    api_id: {label: API ID}              # ラベルの右隣の値（offset で変更可）
    method: {label: メソッド}
    path: {label: パス}
    request:  {title: 入力項目, columns: {name: 項目名, type: 型, required: 必須, max_length: 桁, location: 位置}}
    response: {title: 出力項目, columns: {name: 項目名, type: 型, required: 必須, max_length: 桁}}
    statuses: {title: 応答, column: ステータス}
    type_map: {文字列: string, 数値: integer, 小数: number, 真偽値: boolean, 日付: string, 配列: array, オブジェクト: object}
    required_marks: ["○", "◯", "必須", "Y"]

表の見出しが 2 段のときの列名は「親/子」（例: 型情報/型）. 段数は request などの header_rows で指定できる（省略で自動判定）.
"""

from __future__ import annotations

import math
import unicodedata
from fnmatch import fnmatchcase
from pathlib import Path
from typing import Any

from ..domain import spec_diff
from . import excel_reader


class SpecConfigError(ValueError):
    """整合チェックの設定（config/spec.yaml）の誤り. CLI は終了コード 2."""


def validate_config(conf: Any, name: str = "spec.yaml") -> None:
    """設定の形を、読み始める前に確かめる. 誤りは SpecConfigError（どの項目かを添える）.

    読み取りの途中の KeyError などを「設定の誤り」とみなして終了コード 2 にすると、ツールの不具合（本来は 3）まで
    隠れてしまうため、設定の誤りはここで先に見つける.
    """

    def bad(msg: str) -> SpecConfigError:
        return SpecConfigError(f"{name}: {msg}")

    if not isinstance(conf, dict):
        raise bad("設定は「項目: 値」の辞書で書きます")
    if not isinstance(conf.get("sheets", "*"), str):
        raise bad('sheets はシート名のパターン（文字列）で書きます（例: "API*"）')
    for key in ("api_id", "method", "path"):
        item = conf.get(key)
        if not isinstance(item, dict) or not item.get("label"):
            raise bad(f"{key} は {{label: ラベル}} で書きます（例: {key}: {{label: API ID}}）")
        o = item.get("offset", [0, 1])
        if not (isinstance(o, list) and len(o) == 2 and all(isinstance(x, int) and not isinstance(x, bool) for x in o)):
            raise bad(f"{key} の offset は [行, 列] の整数 2 つで書きます（例: [0, 1]）")
    for key in ("request", "response"):
        part = conf.get(key)
        if part is None:
            continue
        if not isinstance(part, dict) or not isinstance(part.get("columns"), dict) or not part["columns"].get("name"):
            raise bad(f"{key} は {{title: 表題, columns: {{name: 項目名の列, ...}}}} で書きます（columns.name が必要）")
    st = conf.get("statuses")
    if st is not None and (not isinstance(st, dict) or not st.get("column")):
        raise bad("statuses は {title: 表題, column: ステータスの列} で書きます（column が必要）")
    if not isinstance(conf.get("type_map", {}), dict):
        raise bad("type_map は {設計書の型: OpenAPI の型} の辞書で書きます")
    if not isinstance(conf.get("required_marks", []), list):
        raise bad('required_marks は必須を表す印の並び（例: ["○", "必須"]）で書きます')


def _offset(conf: dict) -> tuple[int, int]:
    o = conf.get("offset", [0, 1])
    return int(o[0]), int(o[1])


def _int(v: Any) -> int | None:
    """桁などの数値. セルの値が数値ならそのまま（表示形式に関係なく）、文字列なら「1,000」「５０」も数値として読む.

    読めなければ None（空も None）.
    """
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if math.isfinite(v) else None
    s = unicodedata.normalize("NFKC", str(v or "")).strip().replace(",", "")
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if math.isfinite(f) else None


def _number_cell(ws: Any, ctx: excel_reader.ReadContext, i: int, column: str, shown: str) -> tuple[int | None, bool]:
    """表の i 行目の列 column の数値 → (値, 読めたか). 値はセルの値（表示形式を当てる前）で読む. 空は (None, True)."""
    if not shown:
        return None, True
    col = ctx.table_cols.get(column)
    raw = excel_reader.raw_value(ws, ctx.table_rows[i], col) if col is not None and i < len(ctx.table_rows) else shown
    n = _int(raw)
    if n is None and not isinstance(raw, (int, float)):
        n = _int(shown)
    return n, n is not None


def read(path: Path, conf: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[str], list[str]]:
    """ApiSpec・メモ・警告（数式の計算結果が無いなど）. 読み方の警告（rule_id 付き）は read_with_alerts."""
    spec, notes, warnings, _ = read_with_alerts(path, conf)
    return spec, notes, warnings


def read_with_alerts(path: Path, conf: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[str], list[str], list[tuple[str, str]]]:
    """ApiSpec・メモ・警告・読み方の警告（rule_id, 文）.

    同じ API ID のシートが複数あれば、最初のシートを使い、警告 duplicate_api_id を出す
    （新旧比較の pair_by の duplicate_pair_key と同じ扱い）.
    """
    wb, warnings = excel_reader.load_book(path)
    notes: list[str] = []
    alerts: list[tuple[str, str]] = []
    sheets_of: dict[str, list[str]] = {}
    spec: dict[str, dict[str, Any]] = {}
    type_map = conf.get("type_map", {})
    marks = set(conf.get("required_marks", ["○", "◯", "必須", "Y"]))
    # 表の終わりの目印（設定にある別のラベル・表題）. 表は空行を挟んで続き、次の節の始まりで終わる
    stops = {conf[k]["label"] for k in ("api_id", "method", "path") if isinstance(conf.get(k), dict) and conf[k].get("label")}
    stops |= {conf[k]["title"] for k in ("request", "response", "statuses") if isinstance(conf.get(k), dict) and conf[k].get("title")}
    for ws in wb.worksheets:
        if not fnmatchcase(ws.title, conf.get("sheets", "*")):
            continue
        api_id = excel_reader.read_label(ws, conf["api_id"]["label"], _offset(conf["api_id"]))
        if not api_id:
            notes.append(f"シート「{ws.title}」: API ID のラベル「{conf['api_id']['label']}」が見つからないため読み飛ばしました")
            continue
        sheets_of.setdefault(api_id, []).append(ws.title)
        if len(sheets_of[api_id]) > 1:
            continue  # 同じ API ID の 2 枚目からは読まない（最初のシートを使う）
        entry: dict[str, Any] = {
            "method": (excel_reader.read_label(ws, conf["method"]["label"], _offset(conf["method"])) or "").upper(),
            "path": excel_reader.read_label(ws, conf["path"]["label"], _offset(conf["path"])) or "",
            "request": {},
            "response": {},
            "statuses": [],
        }
        for part in ("request", "response"):
            pc = conf.get(part)
            if not pc:
                continue
            cols = pc["columns"]
            ctx = excel_reader.ReadContext(stop_labels={excel_reader.norm_label(x) for x in stops})
            table = excel_reader.read_table(ws, cols["name"], pc.get("title"), pc.get("header_rows"), ctx)
            notes += [n for n in ctx.notes if n not in notes]
            warnings += [w for _, w in ctx.alerts if w not in warnings]
            if table is None:
                notes.append(f"{api_id}: {part} の表（{pc.get('title') or cols['name']}）が見つかりません")
                continue
            for i, row in enumerate(table[1]):
                name = row.get(cols["name"], "")
                if not name:
                    continue
                raw_type = row.get(cols.get("type", ""), "")
                shown = row.get(cols.get("max_length", ""), "")
                max_length, ok = _number_cell(ws, ctx, i, cols.get("max_length", ""), shown)
                if not ok:
                    alerts.append(
                        (
                            "not_a_number",
                            f"{ws.title}: {api_id} の {part} の項目「{name}」の{cols['max_length']}「{shown}」を数値として読めないため、"
                            "桁は比べていません",
                        )
                    )
                spec_diff.add_field(  # 名前と位置で識別（path の id と body の id は別の項目）
                    entry[part],
                    name,
                    {
                        "type": type_map.get(raw_type, raw_type or None),
                        "required": (row.get(cols.get("required", ""), "") in marks) if cols.get("required") else None,
                        "max_length": max_length,
                        "location": row.get(cols.get("location", ""), "") or ("body" if part == "request" else None),
                    },
                )
        st = conf.get("statuses")
        if st:
            ctx = excel_reader.ReadContext(stop_labels={excel_reader.norm_label(x) for x in stops})
            table = excel_reader.read_table(ws, st["column"], st.get("title"), st.get("header_rows"), ctx)
            notes += [n for n in ctx.notes if n not in notes]
            warnings += [w for _, w in ctx.alerts if w not in warnings]
            if table:
                entry["statuses"] = sorted({r[st["column"]] for r in table[1] if r.get(st["column"])})
        spec[api_id] = entry
    for api_id, titles in sheets_of.items():
        if len(titles) > 1:
            listed = "・".join(f"「{t}」" for t in titles)
            alerts.append(
                (
                    "duplicate_api_id",
                    f"{path.name}: API ID「{api_id}」のシートが {len(titles)} 枚あります（{listed}）。"
                    f"最初の「{titles[0]}」だけを比べ、残りは比べていません。API ID を見直してください",
                )
            )
    warnings += [w for _, w in excel_reader.book_alerts(wb, path.name) if w not in warnings]
    return spec, notes, warnings, alerts
