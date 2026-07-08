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
| `mod_BasicInfoGuide.bas` | 基本情報シートの入力ガイド（ハイライト・コメント）。 |  |
| `mod_BasicInfoGuideTexts.bas` | 入力ガイド文言の集約（mod_BasicInfoGuide から分割）。 |  |
| `mod_BasicInfoCellDropdown.bas` | 基本情報 C22/C23 のダブルクリックドロップダウン。 | `[CellDropdown]` |
| `mod_PrefectureSelector.bas` | 基本情報 C13 の都道府県複数選択（`frmPrefectureSelector`）。 |  |
| `mod_WorkbookOptimize.bas` | 保存前の余剰セル書式クリーンアップ。 | `[WorkbookOptimize]` |
| `mod_FillManagerName.bas` | 出張所長名の自動入力／支店・出張所バリデーションの再構築。 | `[FillMgr]` |
| `mod_VendorMaster.bas` | 業者マスタ参照・業者選択。単価展開・ブロック増減は下記2モジュールへ委譲（既存呼び出し互換のスタブを保持）。 | `[VendorMaster]` |
| `mod_VendorBlockLayout.bas` | 基本情報の業者ブロック増減・レイアウト（`SyncVendorBlocksFromCount` 等、mod_VendorMaster から分割）。 | `[VendorMaster]` |
| `mod_VendorUnitPrice.bas` | 工事単価シートへの業者別単価展開・装飾（`RefreshAllVendorUnitPricesForBasicInfo` 等、mod_VendorMaster から分割）。 | `[VendorMaster]` |
| `mod_VendorInfoColors.bas` | 業者情報色（基本情報10行目・施工指示書等の施工会社列・将来のシートタブ色） | |
| `mod_WeldingUnitPrice.bas` | 「○○保線区_レール溶接単価」シートへ施工会社別単価を展開（`ApplyWeldingVendorUnitPricesForBasicInfo`）。 | `[WeldingUP]` |
| `mod_WeldingTemotoMaster.bas` | レール溶接の手元マスタ（外注費率一覧）のパス解決・シート名照合（mod_WeldingUnitPrice から分割）。 |  |
| `mod_MaterialPriceImport.bas` | 工事単価インポート。単価表・購入充当単価・レール溶接単価シートの作成（`ImportUnitPriceData`）。 | `[UnitPrice]` |
| `mod_Construction_Order_Import.bas` | 施工指示書取込の公開 API ファサード（外部モジュールからの呼び出し口）。 | `[ConstructionImport]` |
| `mod_Construction_Import_Shared.bas` | 施工指示書取込の共通定数・モジュール変数・ログ (`LogCI`)。 | `[ConstructionImport]` |
| `mod_Construction_Import_Load.bas` | 施工指示書・施工通知書・購入充当通知のファイル取込と出力シート作成。 | `[ConstructionImport]` |
| `mod_Construction_SubconPrice.bas` | 施工会社別単価・金額列の展開と更新 (`RefreshSubcontractorPriceColumns`)。 | `[ConstructionImport]` |
| `mod_Construction_BasicTotals.bas` | 基本情報の工事合計・業者別合計・税込計算、業者別名マップ。 | `[ConstructionImport]` |
| `mod_Construction_OutputFormat.bas` | 産廃行制限、合計行描画、参照単価列 (M/N) の再計算。 | `[ConstructionImport]` |
| `mod_Construction_LineMapping.bas` | 工事件名別マスタ F/G 線区名エイリアスと単価シート名解決。 | `[ConstructionImport]` |
| `mod_Construction_OutputLayout.bas` | 出力シート列位置・溶接判定・購入充当単価転記などレイアウト補助。 | `[ConstructionImport]` |
| `mod_subcontractorselector.bas` | 外注業者選択（`frmSubconSelector`）関連の制御。 | |
| `mod_OrderTpl_Shared.bas` | 注文書テンプレート取込の共通定数・業者マスタ照合(A/P/O列)・部店コード解決・取込元シート探索。 | `[OrderTpl]` |
| `mod_OrderTpl_Generate.bas` | 施工会社確定時のテンプレート5シート挿入・削除・再作成(`GenerateVendorOrderSheets`/`RefreshAllVendorOrderDetails`)。 | `[OrderTpl]` |
| `mod_OrderTpl_Header.bas` | 内訳明細ヘッダー部(部店コード・注文番号・工期・外注会社等)の転記。 | `[OrderTpl]` |
| `mod_OrderTpl_Detail.bas` | 内訳明細明細部の転記エンジン(セクション構築・ソート・行挿入・書式)。 | `[OrderTpl]` |
| `ModuleExport.bas` | VBA モジュールのエクスポート用。 | |

## フォーム（`注文書_請求書VBA/運用中VBA/`、`.frm` / `.frx`）

| ファイル | 役割 |
|---|---|
| `Project_Number_Selection.frm` | 工事番号選択フォーム。 |
| `SelectLineName.frm` | 適用線区（単価シート）選択フォーム。 |
| `AllVenderSelection.frm` | 全業者選択フォーム。 |
| `frmSubconSelector.frm` | 外注業者選択フォーム。 |
| `frmPrefectureSelector.frm` | 都道府県複数選択フォーム。 |

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
