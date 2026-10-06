# design-diff

設計書のレビューで「どこが変わったか」「設計書と API 仕様が合っているか」を、レビューにそのまま貼れる Markdown で出すツールです。

| コマンド | 比べるもの | 使いどころ |
|---|---|---|
| `design-diff excel` | Excel 設計書の **新旧**（API・バッチ・ファイル連携など） | 改訂レビューで変更箇所を漏れなく確認する |
| `design-diff spec` | **API 定義書（Excel）と OpenAPI** | 設計書と実装用の API 仕様のずれを見つける |

Excel のどこを読むか（ラベル・表・キー列）はコードではなく `config/*.yaml` に書くので、
設計書の様式が違っても設定だけで合わせられます。

## すぐ使う

1. clone する
2. **Windows**: `setup.cmd` をダブルクリック ／ **macOS・Linux**: `./setup.sh`
3. `.venv` を有効にしてサンプルで試す

```bash
design-diff excel --type api   --before samples/before/API0001_会員照会.xlsx --after samples/after/API0001_会員照会.xlsx
design-diff excel --type batch --before samples/before/JOB0001_注文取込.xlsx --after samples/after/JOB0001_注文取込.xlsx
design-diff excel --all                      # config/global.yaml の pairs をまとめて output/ へ
design-diff spec --excel samples/spec/API定義書.xlsx --oas samples/spec/openapi.yaml
```

サンプルは架空の会員・注文システムの設計書です（`python samples/make_samples.py` で作り直せます）。

出力の例（新旧差分の一部）:

```markdown
## セルの変更

| 場所 | 項目 | 左 | 右 |
|---|---|---|---|
| API0001_会員詳細照会 / 出力項目 | name の 桁 | 50 | 60 |
```

## オプション

| オプション | 意味 |
|---|---|
| `-f markdown｜json｜…` | 出力形式（既定 Markdown。JSON は devtools 共通のレポート形式） |
| `-o FILE` | ファイルに書き出す |
| `--fail-on-diff` | 差分があれば終了コード 1（CI で「変更なし」を確かめるとき） |
| `--config-dir DIR` | 設定の置き場（既定 `config/`） |
| `spec --config FILE` | 整合チェックの設定（既定 `config/spec.yaml`） |

終了コード: `0` 正常（差分の有無は問わない）／ `1` `--fail-on-diff` で差分あり、または設定の見直しが必要 ／ `2` 引数・設定の誤り ／ `3` 実行時エラー

## 自分の設計書に合わせる

`config/api.yaml` をコピーして、比べたい項目を並べます。

```yaml
sheets:
  - name: "API*"              # 対象シート（ワイルドカード）
    pair_by: "^(API\\d+)"      # シート名が変わっても API ID で新旧を対応付ける（省略時はシート名で）
    items:
      - {kind: label, name: API 名, label: API 名}                         # ラベルの右隣の値
      - {kind: block, name: 処理概要, label: 処理概要}                       # ラベルの下の段落（空行まで）
      - {kind: table, name: 入力項目, title: 入力項目, header: 項目名, key: [項目名]}  # 表（キー列で行を対応付け）
```

- `label` の位置がずれている様式は `offset: [行, 列]` で指定します（右隣は `[0, 1]`、真下は `[1, 0]`）
- 表は「見出し行から、全列が空の行まで」です。同じ列名の表が複数あるときは `title`（表の上のラベル）で区別します
- 見出しが 2 段の表（1 段目で結合したグループ見出し、2 段目に子見出し）は、列名を「親/子」にします（例: `型情報/桁`）。段数は結合セルから判断します。うまく判断できないときは `header_rows: 2` のように指定します
- 結合セルは左上の値として読みます。ラベルの末尾の「：」や空白の違いは無視します
- 数式セルは計算結果で比べます。ツールで作った Excel など計算結果が保存されていないときは数式の文字列で比べ、レポートに**注意**を出します（Excel で開いて保存し直すと計算結果で比べられます）
- 「設定の見直しが必要」と出たら、ラベル名か `key` の列名が設計書と合っていません

整合チェック（`spec`）の設定は `config/spec.yaml` です。OpenAPI 側は operation の `x-api-id`（無ければ `operationId`）を
API ID として設計書のシートと対応付けます。型の名前は `type_map` で OpenAPI の型にそろえます。
OpenAPI の `allOf` はつないで 1 つのスキーマとして、`oneOf` / `anyOf` は全候補の項目を合わせて比べます。
候補があることはレポートの冒頭に注記するので、どの候補の項目かを人が確認してください（`discriminator` は候補の選び分けには使いません）。

詳しい仕様（差分の種類の一覧、比べる・比べない項目）は [docs/SPECIFICATION.md](docs/SPECIFICATION.md) を参照してください。

## Copilot から使う

VS Code の Copilot Chat で `/design-diff-review` を実行すると、差分を取ってレビューの観点（MUST / WANT / NITS / Q）で要約します。
レビューハーネス（office-review）の指摘の書式と同じです。

## 開発

```bash
./setup.sh                    # Windows: setup.cmd
ruff check . && ruff format --check . && pytest
```

構成: `cli`（引数）→ `application`（設定に従って読む・比べる・出力する）→ `domain`（比較のルール）/ `infrastructure`（Excel・OpenAPI の読み取り）

## セキュリティ・ライセンス

- 脆弱性を見つけたら、公開の Issue ではなく [SECURITY.md](SECURITY.md) の方法で知らせてください
- ライセンスは [MIT](LICENSE) です
