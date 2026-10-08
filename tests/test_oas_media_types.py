"""OpenAPI の要求・応答の本体: application/json 以外の JSON のメディアタイプ（パラメーター付き・*+json・*/*）も読む."""

from __future__ import annotations

from pathlib import Path

import pytest

from design_diff.infrastructure import oas_reader

OAS = """openapi: 3.0.3
info: {{title: t, version: '1'}}
paths:
  /items:
    post:
      x-api-id: API0001
      requestBody:
        content:
          {req}:
            schema:
              type: object
              properties:
                title: {{type: string, maxLength: 20}}
      responses:
        '200':
          content:
            {res}:
              schema:
                type: object
                properties:
                  id: {{type: integer}}
"""


def _read(tmp_path: Path, req: str, res: str) -> dict:
    p = tmp_path / "openapi.yaml"
    p.write_text(OAS.format(req=req, res=res), encoding="utf-8")
    return oas_reader.read(p)["API0001"]


@pytest.mark.parametrize(
    "media",
    ["application/json", "'application/json; charset=utf-8'", "application/problem+json", "application/vnd.example+json", "'*/*'"],
)
def test_json_like_media_types_are_read(tmp_path: Path, media: str):
    api = _read(tmp_path, media, media)
    assert api["request"]["title"]["max_length"] == 20
    assert api["response"]["id"]["type"] == "integer"


def test_xml_only_is_not_read(tmp_path: Path):
    api = _read(tmp_path, "application/xml", "text/plain")
    assert api["request"] == {} and api["response"] == {}


def test_json_is_preferred_over_wildcard():
    content = {
        "*/*": {"schema": {"type": "object", "properties": {"wild": {"type": "string"}}}},
        "application/json;charset=UTF-8": {"schema": {"type": "object", "properties": {"json": {"type": "string"}}}},
    }
    assert set(oas_reader.json_schema(content)["properties"]) == {"json"}
    assert oas_reader.json_schema({"*/*": content["*/*"]}) is content["*/*"]["schema"]
    assert oas_reader.json_schema(None) is None


@pytest.mark.parametrize(
    ("key", "expected"),
    [
        ("application/json", True),
        ("Application/JSON; charset=utf-8", True),
        ("application/merge-patch+json", True),
        ("application/jsonl", False),
        ("text/json-ish", False),
        ("*/*", False),
    ],
)
def test_is_json_media_type(key: str, expected: bool):
    assert oas_reader.is_json_media_type(key) is expected
