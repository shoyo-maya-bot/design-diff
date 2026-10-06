"""OpenAPI の allOf / oneOf / anyOf / discriminator の読み方のテスト（小さな OpenAPI を一時ファイルに書いて読む）."""

from __future__ import annotations

import textwrap
from pathlib import Path

from design_diff.domain.spec_diff import compare
from design_diff.infrastructure import oas_reader

HEAD = """\
openapi: 3.0.3
info: {title: t, version: "1"}
paths:
  /pets:
    post:
      operationId: API0001
      requestBody:
        content:
          application/json:
            schema: {$ref: "#/components/schemas/Req"}
      responses:
        "200":
          description: ok
          content:
            application/json:
              schema: {$ref: "#/components/schemas/Res"}
components:
  schemas:
"""


def _write(tmp_path: Path, schemas: str) -> Path:
    p = tmp_path / "openapi.yaml"
    p.write_text(HEAD + textwrap.indent(textwrap.dedent(schemas), "    "), encoding="utf-8")
    return p


def test_allof_merges_properties_and_required(tmp_path: Path):
    p = _write(
        tmp_path,
        """
        Base:
          type: object
          required: [id]
          properties:
            id: {type: string, maxLength: 10}
        Named:
          allOf:
            - $ref: "#/components/schemas/Base"
            - type: object
              required: [name]
              properties:
                name: {type: string, maxLength: 50}
        Req:
          allOf:
            - $ref: "#/components/schemas/Named"
            - type: object
              properties:
                memo: {type: string}
                owner:
                  allOf:
                    - $ref: "#/components/schemas/Base"
          required: [memo]
        Res: {$ref: "#/components/schemas/Named"}
        """,
    )
    spec, notes = oas_reader.read_with_notes(p)
    req = spec["API0001"]["request"]
    assert set(req) == {"id", "name", "memo", "owner", "owner.id"}
    assert req["id"] == {"type": "string", "required": True, "max_length": 10, "location": "body"}
    assert req["name"]["required"] is True and req["name"]["max_length"] == 50
    assert req["memo"]["required"] is True  # 親の required が allOf の中の項目にも効く
    assert req["owner"]["type"] == "object"
    assert set(spec["API0001"]["response"]) == {"id", "name"}
    assert notes == []
    # 以前は allOf を読めず、設計書との比較で項目が全部「OpenAPI に無い」になっていた
    doc = {"API0001": {**spec["API0001"], "request": {"id": {"type": "string", "required": True, "max_length": 10}}}}
    assert [c.category for c in compare(doc, spec) if c.location == "API0001 request" and c.field == "id"] == []


def test_oneof_collects_union_and_notes_variants(tmp_path: Path):
    p = _write(
        tmp_path,
        """
        Cat:
          type: object
          required: [petType, meow]
          properties:
            petType: {type: string}
            meow: {type: string, maxLength: 20}
        Dog:
          type: object
          required: [petType]
          properties:
            petType: {type: string}
            bark: {type: boolean}
        Req:
          type: object
          required: [ownerId]
          properties:
            ownerId: {type: string}
          oneOf:
            - $ref: "#/components/schemas/Cat"
            - $ref: "#/components/schemas/Dog"
          discriminator:
            propertyName: petType
        Res:
          anyOf:
            - type: object
              properties:
                a: {type: string}
            - type: object
              properties:
                b: {type: integer}
        """,
    )
    spec, notes = oas_reader.read_with_notes(p)
    req = spec["API0001"]["request"]
    assert set(req) == {"ownerId", "petType", "meow", "bark"}
    assert req["ownerId"]["required"] is True and "variant" not in req["ownerId"]
    assert req["petType"]["required"] is True  # 全候補で必須
    assert req["petType"]["variant"] == "oneOf 候補1・2"
    assert req["meow"]["required"] is None  # 候補によって違うので比べない
    assert req["meow"]["variant"] == "oneOf 候補1"
    assert req["bark"]["required"] is False and req["bark"]["variant"] == "oneOf 候補2"  # 唯一の候補で任意
    res = spec["API0001"]["response"]
    assert res["a"]["variant"] == "anyOf 候補1" and res["b"]["type"] == "integer"
    assert len(notes) == 2
    assert "API0001 request" in notes[0] and "oneOf" in notes[0] and "petType" in notes[0]
    assert "anyOf" in notes[1]
    # 必須を比べない項目は、設計書が「必須」でも required_mismatch にならない
    doc = {"API0001": {**spec["API0001"], "request": {k: {**v, "required": True} for k, v in req.items()}}}
    assert [c.field for c in compare(doc, spec) if c.category == "required_mismatch"] == ["bark"]  # meow は比べない
    # 設計書に無い候補の項目は、どの候補の項目かを添えて出す
    doc = {"API0001": {**spec["API0001"], "request": {"ownerId": req["ownerId"]}}}
    missing = {c.field: c.after for c in compare(doc, spec) if c.category == "field_missing_in_doc"}
    assert missing["bark"] == "oneOf 候補2"


def test_nullable_oneof_property_keeps_type(tmp_path: Path):
    p = _write(
        tmp_path,
        """
        Req:
          type: object
          properties:
            code:
              oneOf:
                - {type: string, maxLength: 5}
                - {type: "null"}
            mixed:
              anyOf:
                - {type: string}
                - {type: integer}
            labels: {type: [string, "null"], maxLength: 3}
        Res:
          type: array
          items:
            oneOf:
              - type: object
                properties:
                  x: {type: string}
        """,
    )
    spec, notes = oas_reader.read_with_notes(p)
    req = spec["API0001"]["request"]
    assert set(req) == {"code", "mixed", "labels"}
    assert req["code"]["type"] == "string" and req["code"]["max_length"] == 5
    assert req["mixed"]["type"] is None  # 型が候補で違う → 型は比べない
    assert req["labels"]["type"] == "string"
    assert spec["API0001"]["response"]["x"]["variant"] == "oneOf 候補1"
    assert len(notes) == 2  # null との oneOf（null 許可）は確認不要なので出さない
    assert "mixed" in notes[0] and "string" in notes[0] and "integer" in notes[0]
    assert "response" in notes[1]


def test_read_still_returns_spec_only(tmp_path: Path):
    p = _write(tmp_path, "Req: {type: object, properties: {a: {type: string}}}\nRes: {type: object}\n")
    assert set(oas_reader.read(p)["API0001"]["request"]) == {"a"}


def test_self_reference_does_not_loop(tmp_path: Path):
    p = _write(
        tmp_path,
        """
        Node:
          allOf:
            - type: object
              properties:
                name: {type: string}
                child: {$ref: "#/components/schemas/Node"}
        Req: {$ref: "#/components/schemas/Node"}
        Res: {type: object}
        """,
    )
    spec, notes = oas_reader.read_with_notes(p)
    req = spec["API0001"]["request"]
    assert set(req) == {"name", "child"}  # 循環する child の中は展開しない
    assert any("child" in n and "循環" in n for n in notes)
