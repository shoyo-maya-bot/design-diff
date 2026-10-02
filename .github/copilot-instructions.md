# Copilot への指示（design-diff）

設計書レビューのための差分ツールです。説明は日本語で行ってください。

- 層の向きは `cli → application → domain / infrastructure`。domain は Excel・YAML・出力形式を知らない
- Excel の読み方（どのラベル・どの表を比べるか）はコードに書かず `config/*.yaml` に書く
- 差分の種類（category）を増やしたら `application/render.py` の表示名と `docs/SPECIFICATION.md` も直す
- 差分が出たことはエラーではない（終了コード 0）。設定の見直しが必要なとき（ラベル・表・キー列が見つからない）だけ 1
- サンプルの設計書は架空のもの。実在の会社名・システム名・個人名を入れない
- コードを変更したら `ruff check . && ruff format --check . && pytest` を実行して結果を報告する
