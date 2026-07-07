Option Explicit

Public Const DOC_ORDER As Long = 1
Public Const DOC_NOTICE As Long = 2

Public Const DATA_START_ROW As Long = 22

Public Const COL_VENDOR         As Long = 1
Public Const COL_SEIRI          As Long = COL_VENDOR + 1
Public Const COL_TYPE           As Long = COL_SEIRI + 1
Public Const COL_DAYNIGHT       As Long = COL_TYPE + 1
Public Const COL_UNIT           As Long = COL_DAYNIGHT + 1
Public Const COL_QTY            As Long = COL_UNIT + 1
Public Const COL_LINE           As Long = COL_QTY + 1
Public Const COL_MGR            As Long = COL_LINE + 1
Public Const COL_JR_PRICE       As Long = COL_MGR + 1
Public Const COL_JR_AMOUNT      As Long = COL_JR_PRICE + 1
Public Const COL_OUT_PRICE      As Long = COL_JR_AMOUNT + 1
Public Const COL_OUT_AMOUNT     As Long = COL_OUT_PRICE + 1
Public Const COL_KIND           As Long = COL_OUT_AMOUNT + 1
Public Const COL_GAP_AFTER_DATA As Long = COL_KIND + 1
Public Const COL_AUTO_PRICE     As Long = COL_GAP_AFTER_DATA + 1
Public Const COL_AUTO_AMOUNT    As Long = COL_AUTO_PRICE + 1
Public Const COL_PRICE_COMPARE  As Long = COL_AUTO_AMOUNT + 1
Public Const COL_PRICE_GUIDANCE As Long = COL_PRICE_COMPARE + 1
Public Const PRICE_GUIDANCE_COLUMN_WIDTH As Double = 59#
Public Const COL_FLAG_SIDE      As Long = COL_PRICE_GUIDANCE + 9
Public Const COL_FLAG_WELD      As Long = COL_FLAG_SIDE + 1
Public Const OUTPUT_COL_COUNT   As Long = COL_KIND
Public Const WELDING_OUTPUT_COL_OFFSET As Long = 1
Public Const WELDING_OUTPUT_COL_COUNT As Long = OUTPUT_COL_COUNT + WELDING_OUTPUT_COL_OFFSET
Public Const WELD_COL_WELDING_VENDOR As Long = 1
Public Const WELD_COL_TRACK_VENDOR As Long = 2
Public Const WELDING_VENDOR_HEADER As String = "溶接会社"
Public Const TRACK_VENDOR_HEADER As String = "軌道手元会社"
Public Const WELDING_WORK_TYPE_KEYWORD As String = "溶接工事"
Public Const TRACK_WORK_TYPE_KEYWORD As String = "軌道工事"
Public Const BASIC_INFO_VENDOR_WORK_TYPE_ROW As Long = 10

Public Const MGR_MASTER_SHEET As String = "JR管理室対応出張所"
Public Const MGR_MASTER_BRANCH_COL As Long = 2
Public Const MGR_MASTER_OFFICE_COL As Long = 3
Public Const MGR_MASTER_ROOM_COL As Long = 6
Public Const MGR_MASTER_START_ROW As Long = 2

Public Const PURCHASE_KEYWORD As String = "購入充当"
Public Const SIDELINE_KEYWORD As String = "側線"
Public Const WELDING_KEYWORD As String = "レール溶接"
Public Const CONSTRUCTION_SHEET_SUFFIX_WELDING As String = "(溶接)"
Public Const CONSTRUCTION_SHEET_SUFFIX_WORKS As String = "(工事)"
Public Const WELDING_SEIRI_OFFSET As Long = 20000

Public Const SANPAI_KEYWORD As String = "産廃処理"
Public Const SANPAI_FALLBACK_FILL_COLOR As Long = 14277081   ' RGB(217,217,217)

Public Const BASIC_INFO_BRANCH_CELL As String = "B6"
Public Const BASIC_INFO_OFFICE_CELL As String = "C6"
Public Const BASIC_INFO_PUBLIC_CELL As String = "B4"
Public Const BASIC_INFO_AMOUNT_CELL As String = "C22"
Public Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Public Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"
Public Const BASIC_INFO_WORKS_TOTAL_CELL As String = "C31"
Public Const BASIC_INFO_PURCHASE_TOTAL_CELL As String = "C32"
Public Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Public Const BASIC_INFO_VENDOR_TOTAL_ROW As Long = 33
Public Const BASIC_INFO_VENDOR_FIRST_COL As Long = 6
Public Const BASIC_INFO_VENDOR_STEP_COLS As Long = 3
Public Const BASIC_INFO_VENDOR_MAX_BLOCKS As Long = 10
Public Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Public Const BASIC_INFO_TOTAL_NUMBER_FORMAT As String = "#,##0;[赤]-#,##0"
Public Const BASIC_INFO_SUBTOTAL_CELL As String = "C33"
Public Const BASIC_INFO_TAX_CELL As String = "C34"
Public Const BASIC_INFO_GRAND_TOTAL_CELL As String = "C35"
Public Const BASIC_INFO_TAX_LABEL_CELL As String = "B34"
Public Const BASIC_INFO_TAX_RATE_DEFAULT As Double = 0.1
Public Const BASIC_INFO_YEN_TOTAL_RANGE As String = "C31:C35"
Public Const PRICE_GUIDANCE_AMOUNT_TYPE_MESSAGE As String = _
    "基本情報シートのC22セル：単価適用区分(年初単価or設計変更単価)を確認して下さい。"

Public Const REF_VALUE_SOURCE_CELL_NOTICE As String = "H9"
Public Const REF_VALUE_SOURCE_CELL_ORDER As String = "F9"
Public Const BASIC_INFO_REF_VALUE_CELL As String = "C12"
Public Const BASIC_INFO_REF_FONT_NAME As String = "BIZ UDゴシック"

Public Const PROJECT_MASTER_START_ROW As Long = 2
Public Const PROJECT_MASTER_LINE_CODE_COL As Long = 4
Public Const PROJECT_MASTER_UNIT_PRICE_LINE_COL As Long = 6
Public Const PROJECT_MASTER_SOURCE_LINE_COL As Long = 7
Public Const PROJECT_MASTER_LINE_ORDER_UNKNOWN_RANK As Long = 999999
Public Const MASTER_DATA_FOLDER As String = "マスタデータ"
Public Const UNIT_PRICE_LINE_MASTER_FILE As String = "出張所別_単価適用線区.xlsx"
Public Const VENDOR_MASTER_FILE As String = "業者マスタ(全社版).xlsx"
Public Const VENDOR_MASTER_ABBREV_COL As Long = 1     ' A列 業者名(略称) … 帳票・シートヘッダー表記
Public Const VENDOR_MASTER_OFFICIAL_COL As Long = 2   ' B列 請求者氏名(正規名) … 基本情報F11表記
Public Const UNIT_PRICE_DATA_START_ROW As Long = 7

Public Const PRICE_LINE_SHEET As String = "単価適用線区"
Public Const PRICE_LINE_BRANCH_COL As Long = 2
Public Const PRICE_LINE_OFFICE_COL As Long = 3
Public Const PRICE_LINE_NAME_COL As Long = 5
Public Const PRICE_LINE_START_ROW As Long = 2
Public Const PURCHASE_PRICE_SHEET_SUFFIX As String = "_購入充当単価"

' 対象線区(単価)シートの列構成。COL_SEIRI(=2,B)を整理番号キーとする。
Public Const UNIT_PRICE_WORK_KIND_COL As Long = 1   ' A 工種
Public Const UNIT_PRICE_TYPE_COL As Long = 3        ' C 種別
Public Const UNIT_PRICE_UNIT_COL As Long = 4        ' D 単位
Public Const UNIT_PRICE_DAY_PRICE_COL As Long = 5   ' E 昼単価
Public Const UNIT_PRICE_NIGHT_PRICE_COL As Long = 6 ' F 夜単価
Public Const MISSING_SEIRI_FILL_COLOR As Long = 65535 ' RGB(255,255,0) 黄色

Public mVendorAliasMapCache As Object
Public Const WELDING_PRICE_SHEET_SUFFIX As String = "_レール溶接単価"

Public Const PURCHASE_PRICE_KEY_COL As Long = 1
Public Const PURCHASE_PRICE_VALUE_COL As Long = 6
Public Const PURCHASE_PRICE_DATA_START_ROW As Long = 2
Public Const PURCHASE_NOTICE_SEIRI_COL As Long = 1
Public Const PURCHASE_PRICE_LOOKUP_COL As Long = PURCHASE_NOTICE_SEIRI_COL
Public Const PURCHASE_NOTICE_TYPE_COL As Long = COL_TYPE - 1
Public Const PURCHASE_NOTICE_DAYNIGHT_COL As Long = COL_DAYNIGHT - 1
Public Const PURCHASE_NOTICE_UNIT_COL As Long = COL_UNIT - 1
Public Const PURCHASE_NOTICE_QTY_COL As Long = COL_QTY - 1
Public Const PURCHASE_NOTICE_LINE_COL As Long = COL_LINE - 1
Public Const PURCHASE_NOTICE_MGR_COL As Long = COL_MGR - 1
Public Const PURCHASE_NOTICE_JR_PRICE_COL As Long = COL_JR_PRICE - 1
Public Const PURCHASE_NOTICE_JR_AMOUNT_COL As Long = COL_JR_AMOUNT - 1
Public Const PURCHASE_NOTICE_OUT_PRICE_COL As Long = COL_OUT_PRICE - 1
Public Const PURCHASE_NOTICE_OUT_AMOUNT_COL As Long = COL_OUT_AMOUNT - 1
Public Const PURCHASE_NOTICE_KIND_COL As Long = COL_KIND - 1
Public Const PURCHASE_NOTICE_GAP_AFTER_DATA_COL As Long = COL_GAP_AFTER_DATA - 1
Public Const PURCHASE_NOTICE_AUTO_PRICE_COL As Long = COL_AUTO_PRICE - 1
Public Const PURCHASE_NOTICE_AUTO_AMOUNT_COL As Long = COL_AUTO_AMOUNT - 1
Public Const PURCHASE_NOTICE_PRICE_COMPARE_COL As Long = COL_PRICE_COMPARE - 1
Public Const PURCHASE_NOTICE_PRICE_GUIDANCE_COL As Long = COL_PRICE_GUIDANCE - 1

Public Const SUBCON_PRICE_FIRST_COL As Long = 11
Public Const WELDING_SUBCON_PRICE_FIRST_COL As Long = SUBCON_PRICE_FIRST_COL + WELDING_OUTPUT_COL_OFFSET
Public Const WELDING_PRICE_SEIRI_COL As Long = 2
Public Const WELDING_PRICE_DATA_START_ROW As Long = 7
Public Const WELDING_PRICE_WELDING_DAY_COL As Long = 7
Public Const WELDING_PRICE_FIRST_RAIL_DAY_COL As Long = 9
Public Const WELDING_PRICE_VENDOR_NAME_ROW As Long = 5
Public Const UNIT_PRICE_VENDOR_NAME_ROW As Long = 5
Public Const UNIT_PRICE_VENDOR_FIRST_DAY_COL As Long = 7
Public Const UNIT_PRICE_WORK_TYPE_COL As Long = 3
Public Const UNIT_PRICE_REF_DAY_COL As Long = 5
Public Const UNIT_PRICE_REF_NIGHT_COL As Long = 6

Public Const SOURCE_SHEET_NAME_CELL As String = "A3"

Public mSuppressOverwritePrompt As Boolean
Public mLastCreatedImportSheet As Worksheet
Public mSanpaiFillColorCached As Boolean
Public mSanpaiFillColorCache As Long
Public mProjectLineNameAliasMapWelding As Object
Public mProjectLineNameAliasMapConstruction As Object
Public mProjectLineNameReverseAliasWelding As Object
Public mProjectLineNameReverseAliasConstruction As Object
Public mProjectMasterLineOrderRankMap As Object

Public Sub LogCI(ByVal msg As String)
    Debug.Print "[ConstructionImport] " & Format(Now, "hh:mm:ss") & "  " & msg
End Sub
