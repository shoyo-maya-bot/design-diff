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

from fnmatch import fnmatch
from pathlib import Path
from typing import Any

from . import excel_reader


def _offset(conf: dict) -> tuple[int, int]:
    o = conf.get("offset", [0, 1])
    return int(o[0]), int(o[1])


def _int(v: str) -> int | None:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def read(path: Path, conf: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[str], list[str]]:
    """ApiSpec・メモ・警告（数式の計算結果が無いなど）."""
    wb, warnings = excel_reader.load_book(path)
    notes: list[str] = []
    spec: dict[str, dict[str, Any]] = {}
    type_map = conf.get("type_map", {})
    marks = set(conf.get("required_marks", ["○", "◯", "必須", "Y"]))
    # 表の終わりの目印（設定にある別のラベル・表題）. 表は空行を挟んで続き、次の節の始まりで終わる
    stops = {conf[k]["label"] for k in ("api_id", "method", "path") if isinstance(conf.get(k), dict) and conf[k].get("label")}
    stops |= {conf[k]["title"] for k in ("request", "response", "statuses") if isinstance(conf.get(k), dict) and conf[k].get("title")}
    for ws in wb.worksheets:
        if not fnmatch(ws.title, conf.get("sheets", "*")):
            continue
        api_id = excel_reader.read_label(ws, conf["api_id"]["label"], _offset(conf["api_id"]))
        if not api_id:
            notes.append(f"シート「{ws.title}」: API ID のラベル「{conf['api_id']['label']}」が見つからないため読み飛ばしました")
            continue
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
            for row in table[1]:
                name = row.get(cols["name"], "")
                if not name:
                    continue
                raw_type = row.get(cols.get("type", ""), "")
                entry[part][name] = {
                    "type": type_map.get(raw_type, raw_type or None),
                    "required": (row.get(cols.get("required", ""), "") in marks) if cols.get("required") else None,
                    "max_length": _int(row.get(cols.get("max_length", ""), "")),
                    "location": row.get(cols.get("location", ""), "") or ("body" if part == "request" else None),
                }
        st = conf.get("statuses")
        if st:
            ctx = excel_reader.ReadContext(stop_labels={excel_reader.norm_label(x) for x in stops})
            table = excel_reader.read_table(ws, st["column"], st.get("title"), st.get("header_rows"), ctx)
            notes += [n for n in ctx.notes if n not in notes]
            warnings += [w for _, w in ctx.alerts if w not in warnings]
            if table:
                entry["statuses"] = sorted({r[st["column"]] for r in table[1] if r.get(st["column"])})
        if api_id in spec:
            notes.append(f"{api_id}: 複数のシートにあります（後のシート「{ws.title}」を使いました）")
        spec[api_id] = entry
    return spec, notes, warnings
