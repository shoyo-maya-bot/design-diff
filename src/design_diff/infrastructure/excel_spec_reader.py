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


def read(path: Path, conf: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[str]]:
    wb = excel_reader.open_book(path)
    notes: list[str] = []
    spec: dict[str, dict[str, Any]] = {}
    type_map = conf.get("type_map", {})
    marks = set(conf.get("required_marks", ["○", "◯", "必須", "Y"]))
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
            table = excel_reader.read_table(ws, cols["name"], pc.get("title"))
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
            table = excel_reader.read_table(ws, st["column"], st.get("title"))
            if table:
                entry["statuses"] = sorted({r[st["column"]] for r in table[1] if r.get(st["column"])})
        if api_id in spec:
            notes.append(f"{api_id}: 複数のシートにあります（後のシート「{ws.title}」を使いました）")
        spec[api_id] = entry
    return spec, notes
