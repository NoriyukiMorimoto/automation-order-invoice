# automation-order-invoice

注文書・請求書をデータ取り込み等で自動作成する Excel 帳票（`自動化_注文書_請求書.xlsm`）の VBA 一式。

改修履歴は [`CHANGELOG.md`](CHANGELOG.md) を参照。各モジュールのヘッダーには `改修履歴: CHANGELOG.md 参照` と記載しています。

## フォルダ構成

| パス | 内容 |
|---|---|
| `注文書_請求書VBA/運用中VBA/` | 運用中の VBA ソース（標準モジュール・フォーム・クラスモジュール） |
| `注文書_請求書VBA/VBA同期/` | 開いたままの帳票へ VBA を同期するスクリプト（`.vbs` / `.cmd`） |
| `単価マスタ/` | 単価マスタ関連ファイル |
| `CHANGELOG.md` | 改修履歴 |

## 標準モジュール（`注文書_請求書VBA/運用中VBA/`）

| ファイル | 役割 | ログタグ |
|---|---|---|
| `mod_common.bas` | 共通ユーティリティ（文字列正規化、基本情報シート取得、日本語名生成、ADO 接続生成、年抽出、配列ソート 等）。各モジュールから参照する基盤。 | |
| `mod_DebugLog.bas` | デバッグログ出力（イミディエイト＋ログシート）。`[タグ]` 形式で記録。 | |
| `mod_UiMessages.bas` | UI メッセージ文言（`UiMsg*` テキスト）の集約。 | |
| `mod_BasicInfoCalendar.bas` | 基本情報シートの日付入力（カレンダー）関連。 | |
| `mod_BasicInfoUpdate.bas` | 基本情報シートの期間／請求回数の更新・クリア。 | |
| `mod_FillManagerName.bas` | 出張所長名の自動入力／支店・出張所バリデーションの再構築。 | `[FillMgr]` |
| `mod_VendorMaster.bas` | 業者マスタ参照・業者選択。業者別単価ブロックの工事単価シートへの展開（`RefreshAllVendorUnitPricesForBasicInfo` 等）。 | `[VendorMaster]` |
| `mod_WeldingUnitPrice.bas` | 「○○保線区_レール溶接単価」シートへ施工会社別単価を展開（`ApplyWeldingVendorUnitPricesForBasicInfo`）。 | `[WeldingUP]` |
| `mod_MaterialPriceImport.bas` | 工事単価インポート。単価表・購入充当単価・レール溶接単価シートの作成（`ImportUnitPriceData`）。 | `[UnitPrice]` |
| `mod_Construction_Order_Import.bas` | 施工指示書・施工通知書・購入充当通知の取込と単価照合・出力整形。 | `[ConstructionImport]` |
| `mod_subcontractorselector.bas` | 外注業者選択（`frmSubconSelector`）関連の制御。 | |
| `ModuleExport.bas` | VBA モジュールのエクスポート用。 | |

## フォーム（`注文書_請求書VBA/運用中VBA/`、`.frm` / `.frx`）

| ファイル | 役割 |
|---|---|
| `Project_Number_Selection.frm` | 工事番号選択フォーム。 |
| `SelectLineName.frm` | 適用線区（単価シート）選択フォーム。 |
| `AllVenderSelection.frm` | 全業者選択フォーム。 |
| `frmSubconSelector.frm` | 外注業者選択フォーム。 |

## クラスモジュール（`注文書_請求書VBA/運用中VBA/`、`.cls`）

| ファイル | 役割 |
|---|---|
| `ThisWorkbook.cls` | ブックレベルのイベント。 |
| `Sheet1.cls` | 「基本情報」シートのイベント（業者・期間・単価関連の入力処理）。 |
| `Sheet*.cls` | 各ワークシートのシートモジュール群（線区別シート 等）。 |

## VBA 同期スクリプト（`注文書_請求書VBA/VBA同期/`）

| ファイル | 役割 |
|---|---|
| `sync_open_workbook_vba.vbs` | 開いたままの帳票へ VBA を同期。 |
| `開いたままVBA更新.cmd` | 上記スクリプトの起動用バッチ。 |

## ソース形式の取り決め

- 文字コード: **Shift-JIS（CP932）**
- 改行コード: **CRLF**
- 各 `.bas` は **`Option Explicit` 始まり**（その上にヘッダー行を置かない）
- 日本語リテラルは互換性のため `ChrW$` で構築し、`Static` 変数でキャッシュする方針

## 計算（再計算）に関する注意

ブック全体の再計算は **`Application.Calculate`** を使用する（`Workbook` オブジェクトに `Calculate` メソッドは無いため、`ThisWorkbook.Calculate` / `<ブック変数>.Calculate` はコンパイルエラーになる）。特定シートのみは `Worksheet.Calculate` を使用する。

## モジュール一覧の更新について

この一覧は自動生成ではありません。モジュール（`.bas` / `.frm` / `.cls`）を追加・削除・改名した場合は、同じ変更内でこの README の該当表も更新してください（運用指針は `AGENTS.md` 参照）。
