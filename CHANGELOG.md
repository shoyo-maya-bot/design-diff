# Changelog

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)、版は [SemVer](https://semver.org/lang/ja/) に従います。

## [Unreleased]

### Changed

- Excel の読み取り（表示形式・結合セル・複数段の見出し・計算結果の無い数式セルの補完）を devtools-common の
  `devtools_common.excel` に移し、doc-to-markdown と同じ規則で読むようにした（devtools-common は `[excel]` extra 付きで固定）
- **セルの値を表示形式どおりの文字列で比べる**。文字列の `001` と表示形式 `000` の数値 1 は同じ値になり、差分に出なくなる。
  レポートの値も Excel の見た目どおりになる（例: `#,##0` の 1500 は `1,500`、`yyyy/mm/dd` の日付は `2026-04-05 00:00:00` ではなく `2026/04/05`、
  真偽値は `TRUE` / `FALSE`）。値が同じでも表示が変わる表示形式の変更（`000` → `0`）は差分になる
- 見出しの段数の判定: 子見出しとみなすのは文字のセルだけにした（横結合した見出しの下が数値のデータ行を 2 段目の見出しと誤認しない）。
  横結合の下の子見出しが同じ名前でも 2 段と判断する。縦結合の横の行が数値だけなら 1 段
