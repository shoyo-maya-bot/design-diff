# design-diff 仕様

## 1. Excel 設計書の新旧差分（`design-diff excel`）

### 読み方

| kind | 読むもの | 設定 |
|---|---|---|
| `label` | ラベルのセルから `offset`（既定 `[0, 1]`）離れたセルの値 | `label`, `offset` |
| `block` | ラベルから `offset`（既定 `[1, 0]`）の位置から下へ、空のセルまでの各行を改行でつないだ文章 | `label`, `offset` |
| `table` | 見出し行（`header` の列名を含む行、または `title` のラベルの次の行）から、全列が空の行までの表 | `header`, `title`, `key` |

- 値は計算結果（数式ではなく表示される値）を読みます。前後の空白は無視し、`10.0` のような整数の小数は `10` として比べます
- 結合セルは範囲の左上の値として読みます
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

書式（色・罫線・フォント）、図形・画像、コメント、非表示の状態、印刷設定、数式そのもの。
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

読み飛ばしたシート（API ID が無い）や、同じ API ID が複数シートにあることは、レポートの冒頭に注記します。

## 3. 出力

- Markdown: 種類ごとの件数の表と、種類ごとの「場所 / 項目 / 左 / 右」の表。左は変更前（または設計書）、右は変更後（または OpenAPI）
- JSON など: devtools-common の共通レポート形式。`rule_id` = category、`expected` = 左、`actual` = 右。
  設定の見直しが必要な差分は `error`、それ以外の差分は `warning`、注記は `info`
