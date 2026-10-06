"""API 定義書（Excel）と OpenAPI の比較."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..domain import spec_diff
from ..domain.model import DiffResult
from ..infrastructure import excel_spec_reader, oas_reader


def run(conf: dict[str, Any], excel: Path, oas: Path) -> DiffResult:
    doc, notes, warnings = excel_spec_reader.read(excel, conf)
    api, oas_notes = oas_reader.read_with_notes(oas)
    result = DiffResult(f"設計書と OpenAPI の差分: {excel.name} ↔ {oas.name}", f"設計書 {excel.name}", f"OpenAPI {oas.name}")
    result.notes = notes + oas_notes + [f"設計書の API {len(doc)} 件 / OpenAPI の API {len(api)} 件を比較しました"]
    result.warnings = warnings
    result.changes = spec_diff.compare(doc, api)
    return result
