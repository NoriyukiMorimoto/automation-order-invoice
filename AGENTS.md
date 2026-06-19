# AGENTS.md

このリポジトリは Excel 帳票（`自動化_注文書_請求書.xlsm`）の VBA ソースを管理します。
エージェント（Cursor / Codex / Claude 等）は以下の規約に従って作業してください。

## ソース形式（厳守）
- 文字コード: **Shift-JIS (CP932)**。UTF-8 で保存しない（VBE インポート時に文字化けする）。
- 改行コード: **CRLF**。
- 各 `.bas` は **`Option Explicit` 始まり**。その上にヘッダー行を置かない。
- 日本語リテラルは互換性のため `ChrW$` で構築し、`Static` 変数でキャッシュする。

## ファイルの場所
- 運用中の VBA: `注文書_請求書VBA/運用中VBA/`（標準 `.bas` / フォーム `.frm`,`.frx` / クラス `.cls`）
- 改修履歴: `CHANGELOG.md`（各モジュールのヘッダーは「改修履歴: CHANGELOG.md 参照」）

## README.md の用途と更新ルール
- `README.md` は VBA モジュールの索引（マニフェスト）。ファイルを素早く特定するために使う。
- モジュール（`.bas` / `.frm` / `.cls`）を **追加・削除・改名** したら、**同じコミット内で** `README.md` の該当表を必ず更新する。
- 各行は「ファイル名・役割・ログタグ」を記載。役割は 1 行で簡潔に。
- README は自動生成されない。更新は人またはエージェントの責任で行う。

## 計算（再計算）
- ブック全体の再計算は `Application.Calculate` を使う。
- `Workbook` に Calculate メソッドは無い。`ThisWorkbook.Calculate` / `<ブック変数>.Calculate` はコンパイルエラーになるため使わない。
- 特定シートのみは `Worksheet.Calculate`。

## 性能
- 大量セルへの書き込みは、配列に組み立ててから範囲へ一括代入する（セル単位の往復を避ける）。
- 重い処理は `ScreenUpdating=False` / `Calculation=xlCalculationManual` で囲み、再計算は最後に 1 回だけ行う。
