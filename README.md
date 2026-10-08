# design-diff

設計書のレビューで「どこが変わったか」「設計書と API 仕様が合っているか」を、レビューにそのまま貼れる Markdown で出すツールです。

| コマンド | 比べるもの | 使いどころ |
|---|---|---|
| `design-diff excel` | Excel 設計書の **新旧**（API・バッチ・ファイル連携など） | 改訂レビューで変更箇所を確認する（設定した項目は項目ごとに、設定に無いセルは件数と場所を出す） |
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
design-diff excel --all                      # config/global.yaml の pairs をまとめて output/ へ（一覧は標準出力）
design-diff spec --excel samples/spec/API定義書.xlsx --oas samples/spec/openapi.yaml
python -m design_diff excel --all            # design-diff コマンドが PATH に無いときは python -m でも同じ
```

`config/global.yaml` の `pairs`（`before` / `after`）と `output_dir` の相対パスは、**global.yaml のあるフォルダから**解決します
（どこで実行しても同じファイルを読む）。以前の「実行場所からの相対」で書いたパスも、global.yaml 基準で見つからなければ
そのまま使い、書き換えを促す警告を標準エラーに出します。

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
| `-f markdown｜json｜…` | 出力形式（既定 Markdown。JSON は devtools 共通のレポート形式）。`--all` ではペアごとのファイルもこの形式（拡張子 `.md` / `.json` …） |
| `-o FILE` | ファイルに書き出す。`--all` では一覧（ペアごとの差分の件数と出力先）を書く |
| `--fail-on-diff` | 差分があれば終了コード 1（CI で「変更なし」を確かめるとき） |
| `--config-dir DIR` | 設定の置き場（既定 `config/`）。サブコマンドの前でも後でも書ける |
| `excel --outside summary｜all｜off` | 設定に無いセル（承認者欄など）の変更の出し方。既定 `summary` はシートごとに件数と場所、`all` は 1 セルずつ、`off` は比べない |
| `spec --config FILE` | 整合チェックの設定（既定 `config/spec.yaml`） |

### 終了コード

| コード | 意味 | 主な場合 |
|---|---|---|
| `0` | 正常 | 比べ終えた（`--fail-on-diff` なしなら差分の有無は問わない） |
| `1` | **差分あり・要確認** | `--fail-on-diff` で差分があった／「設定の見直しが必要」な差分（`label_not_found` など）があった／`--all` で比べられなかったペア（`pair_missing` / `pair_failed`）があった |
| `2` | **入力・使い方の誤り** | 引数の誤り、設定ファイルが無い・誤り（`kind`・`pair_by` など）、`--before` / `--after` / `--excel` / `--oas` のファイルが無い、Excel を開けない（壊れている・xlsx の形式でない）、`spec` の OpenAPI の構文（YAML / JSON）・構造の誤り |
| `3` | **内部エラー** | 上のどれでもない想定外の例外（`--log-level DEBUG` で詳細） |

CI では「`0` 以外は止める」、差分だけを見るなら「`1` は差分、`2` と `3` は実行の失敗」と区別できます。

`--all` は 1 つのペアのファイルが無い・読めないときも、警告（標準エラーと一覧の `pair_missing` / `pair_failed`）を出して残りのペアを比べ、最後に終了コード 1 で知らせます。
出力名（`diff_<変更後のファイル名>`）が他のペアと重なるときは `diff_spec-2.md` のように番号を付けます（上書きしません）。

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
- 表は「見出し行から、表の終わりまで」です。途中の空行は飛ばして読み、**空行が 2 行続く**（`end_blank_rows` で変更）か、
  次の節の始まり（設定にある別のラベル・表題、空行の後の 1 セルだけの行、同じ見出し行の繰り返し）で終わりとします。
  空行を挟んで読んだときは注記を、空行で終わりとした下に表の続きらしい行があるときは警告を出します。文章（`block`）も同じです
- 同じ列名の表が複数あるときは `title`（表の上のラベル）で区別します
- 同じ名前の列が 2 列以上ある表（「備考」が 2 列など）は、2 列目から `備考#2` のように番号を付けて両方比べ、警告を出します
- キーが重複した行は、同じキーの中で出現順に対応付けて比べます（2 件目は「キー（2 件目）」と表示）
- 設定に無いセル（表紙・承認者欄など、どの項目でも読まなかったセル）も比べ、変更があれば「比較対象外のセルの変更」に件数と場所を出します。
  行の挿入でずれただけのセルは差分にしません。1 セルずつ見るときは `--outside all` です
- 見出しが 2 段の表（1 段目で結合したグループ見出し、2 段目に子見出し）は、列名を「親/子」にします（例: `型情報/桁`）。段数は結合セルから判断します。うまく判断できないときは `header_rows: 2` のように指定します
- 結合セルは左上の値として読みます。ラベルの末尾の「：」や空白の違いは無視します
- セルの値は**表示形式どおりの文字列**で比べます（例: 表示形式 `000` の数値 1 は `001`、`#,##0` の 1500 は `1,500`、`yyyy/mm/dd` の日付は `2026/04/05`）。文字列で入力した `001` と、表示形式でゼロ埋めした数値 1 は同じ値とみなします。値が同じでも表示形式を変えて見た目が変わったセルは差分になります
- 数式セルは計算結果で比べます。ツールで作った Excel など計算結果が保存されていないときは数式の文字列で比べ、レポートに**注意**を出します（Excel で開いて保存し直すと計算結果で比べられます）
- 「設定の見直しが必要」と出たら、ラベル名か `key` の列名が設計書と合っていません
- `pair_by` のキーが同じシートが 2 枚以上あると、最初のシートだけを対応付けて比べ、警告 `duplicate_pair_key` を出します
- 表のキー列が空の行は対応付けられないので比べず、警告 `blank_key`（件数と行番号）を出します
- `spec` で同じ API ID のシートが 2 枚以上あると、最初のシートだけを比べ、警告 `duplicate_api_id` を出します
- 走査するのは値のあるセルだけです（書式だけが末尾まで広がったシートでも重くならない）。1 シートの値のあるセルが 100 万個を超えると、
  そこまでの行だけを読んで警告 `sheet_too_large` を出します

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
pip install -c constraints.txt -e ".[dev]"   # CI と同じ版で入れる場合（constraints.txt は deps-lock ワークフローが更新）
ruff check . && ruff format --check . && pytest
```

構成: `cli`（引数）→ `application`（設定に従って読む・比べる・出力する）→ `domain`（比較のルール）/ `infrastructure`（Excel・OpenAPI の読み取り）

## セキュリティ・ライセンス

- 脆弱性を見つけたら、公開の Issue ではなく [SECURITY.md](SECURITY.md) の方法で知らせてください
- ライセンスは [MIT](LICENSE) です
