# design-diff 仕様

## 1. Excel 設計書の新旧差分（`design-diff excel`）

### 読み方

| kind | 読むもの | 設定 |
|---|---|---|
| `label` | ラベルのセルから `offset`（既定 `[0, 1]`）離れたセルの値 | `label`, `offset` |
| `block` | ラベルから `offset`（既定 `[1, 0]`）の位置から下へ、空のセルまでの各行を改行でつないだ文章 | `label`, `offset` |
| `table` | 見出し行（`header` の列名を含む行、または `title` のラベルの次の行）から、全列が空の行までの表 | `header`, `title`, `key`, `header_rows` |

- 値は計算結果（数式ではなく表示される値）を、**セルの表示形式どおりの文字列**で読みます（devtools-common の `devtools_common.excel.render_value`）。例: 表示形式 `000` の 1 は `001`、`#,##0` の 1500 は `1,500`、`0.0%` の 0.125 は `12.5%`、`yyyy/mm/dd` の日付は `2026/04/05`、`ggge"年"` は `令和8年`。表示形式が「標準」なら `10.0` のような整数の小数は `10`。前後の空白は無視します
  - 文字列で入力した `001` と、表示形式 `000` の数値 1 は同じ値です。値が同じでも表示が変わる表示形式の変更（`000` → `0`）は差分になります
  - 再現できない表示形式（指数・分数・条件付き 等）は書式なしの値で比べます
- 計算結果が保存されていない数式セル（ツールで作った Excel によくあります）は、数式の文字列（例: `=SUM(B4:B5)`）を値として読みます。
  このときはレポートの冒頭に**注意**（JSON では `rule_id` = `formula_cache_missing`、`warning`）を出します。
  Excel で開いて保存し直すと計算結果が保存され、計算結果で比べられます。
  片方のファイルだけ計算結果が無いと、同じ内容でも「計算結果」と「数式」の違いとして差分に出ます。`spec` の Excel も同じ読み方です
- 結合セルは範囲の左上の値として読みます
- 見出しが複数段の表は、各段の見出しを「/」でつないだ列名にします（例: 1 段目「型情報」を 2 列に結合し、2 段目が「型」「桁」なら `型情報/型`・`型情報/桁`）。縦に結合した見出し（例: 2 段ぶんの「項目名」）はそのまま `項目名` です
  - 段数は結合セルから判断します（見出し行の結合セルが次の行まで縦に続き、その行が文字だけ、または横に結合した見出しの下の各列がすべて文字）。最大 3 段です。データ行まで同じように横結合しているだけの表や、結合の下が数値の表は 1 段です（doc-to-markdown と同じ規則。devtools-common の `header_depth`）
  - `header_rows`（1 以上の整数）で段数を指定できます。`header` には 1 段目の列名を書きます。`key` や `spec` の `columns` には「親/子」の列名を書きます
- ラベルの照合では末尾の「:」「：」、空白、セル内改行を無視します
- シートは `name`（ワイルドカード）で選び、新旧をシート名で対応付けます。`pair_by`（正規表現の 1 つ目のグループ）を書くと、その部分で対応付けます

### 差分の種類

| category | 表示名 | 補足 |
|---|---|---|
| `sheet_added` / `sheet_removed` | 追加・削除されたシート | |
| `sheet_renamed` | シート名の変更 | `pair_by` で対応付けたとき |
| `value_changed` | 値の変更 | label / block |
| `label_added` / `label_removed` | 追加・削除された項目 | 片方にだけラベルがある |
| `label_not_found` | どちらにも無いラベル | **設定の見直しが必要** |
| `table_not_found` | 表が見つからない | **設定の見直しが必要** |
| `column_added` / `column_removed` | 追加・削除された列 | |
| `key_column_missing` | キー列が無い | **設定の見直しが必要**（行の比較はしない） |
| `duplicate_key` | キーの重複 | **設定の見直しが必要**（後の行で比べる） |
| `row_added` / `row_removed` | 追加・削除された行 | 行の内容を「列=値」で要約 |
| `cell_changed` | セルの変更 | 「キー の 列名」で場所を示す |

「設定の見直しが必要」な差分が 1 件でもあると終了コード 1 です（差分を取り違えたまま「変更なし」と判断しないため）。

### 比べないもの

書式（色・罫線・フォント）、図形・画像、コメント、非表示の状態、印刷設定、数式そのもの（計算結果が保存されていないセルを除く）。
文章やコメントまで含めて見たいときは doc-to-markdown で Markdown にしてから `git diff` してください。

## 2. API 定義書と OpenAPI の整合（`design-diff spec`）

両方を次の形にそろえてから比べます。

```text
API ID → method, path, request{項目名: type, required, max_length, location}, response{…}, statuses[]
```

- 設計書: `sheets` に合うシートを 1 API とし、`api_id` のラベルの値を API ID にします
- OpenAPI: `x-api-id` → `operationId` → `"METHOD path"` の順で API ID にします。
  path / query / header のパラメータと JSON の requestBody を request、最初の 2xx 応答の JSON を response とします。
  入れ子のオブジェクトは `親.子` の項目名に、配列は要素の項目にします。`$ref` はファイル内のものだけ解決します
- OpenAPI のスキーマの組み合わせは次のように読みます
  - `allOf`: 全部の `properties` と `required` をつないで 1 つのスキーマにします（`$ref` や入れ子の `allOf` も展開します）。
    親の `required` は `allOf` の中で定義された項目にも効きます
  - `oneOf` / `anyOf`: 全候補の項目を合わせて（和集合で）比べます。必須は「全候補で必須なら必須、どの候補でも任意なら任意」、
    候補によって違う項目は必須を比べません。どの候補がどの項目を持つかと、候補があることをレポートの冒頭に注記（`info`）するので、人が確認してください。
    設計書に無い候補の項目（`field_missing_in_doc`）は、右の列に「oneOf 候補2」のようにどの候補の項目かを出します
  - `oneOf: [X, {type: "null"}]` のような null を許すだけの書き方と、OpenAPI 3.1 の `type: [string, "null"]` は、null でない方の型として読みます（注記は出しません）。
    候補で型が違う項目（例: `anyOf: [string, integer]`）は型を比べません
  - `discriminator`: 候補の選び分けには使いません（`mapping` も読みません）。判別用の項目名を注記に添えるだけで、その項目自体は各候補の `properties` から読みます
  - 自分自身を参照するスキーマ（木構造など）は、2 回目に出てきたところで展開をやめ、注記します
- パスは `{…}` の名前の違いと末尾の `/` を無視して比べます

| category | 表示名 |
|---|---|
| `api_missing_in_oas` / `api_missing_in_doc` | 片方にしか無い API |
| `method_mismatch` / `path_mismatch` | メソッド・パスの不一致 |
| `field_missing_in_oas` / `field_missing_in_doc` | 片方にしか無い項目 |
| `type_mismatch` | 型の不一致（設計書の型は `type_map` で変換） |
| `required_mismatch` | 必須の不一致（`required_marks` のどれかなら必須） |
| `length_mismatch` | 桁の不一致（設計書に桁があるときだけ） |
| `location_mismatch` | 受け渡し位置の不一致（設計書に位置があるときだけ） |
| `status_missing_in_oas` / `status_missing_in_doc` | 片方にしか無い応答ステータス（設計書に応答の表があるときだけ） |

読み飛ばしたシート（API ID が無い）、同じ API ID が複数シートにあること、OpenAPI の `oneOf` / `anyOf` や循環参照は、レポートの冒頭に注記します。

## 3. 出力

- Markdown: 種類ごとの件数の表と、種類ごとの「場所 / 項目 / 左 / 右」の表。左は変更前（または設計書）、右は変更後（または OpenAPI）
- JSON など: devtools-common の共通レポート形式。`rule_id` = category、`expected` = 左、`actual` = 右。
  設定の見直しが必要な差分は `error`、それ以外の差分と注意（`formula_cache_missing`）は `warning`、注記は `info`
