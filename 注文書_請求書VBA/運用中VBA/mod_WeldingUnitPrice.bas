Option Explicit

' =====================================================================
' mod_WeldingUnitPrice
' 「○○保線区_レール溶接単価」シートへ施工会社別単価を展開するモジュール
'
' ・溶接会社 : G列(昼)/H列(夜)
'              単価 = JR単価 ×(1－手元比率)× 溶接工事外注比率(基本情報 31行目)
' ・軌道会社 : I列(昼)/J列(夜)から1社2列ずつ右へ追加
'              照合キー = 整理番号(溶接単価シートB列 = マスタ溶接手元割合シートA列)
'              単価 = JR単価 ×(100/100.7) × 手元割合(マスタ昼E/夜F) × 軌道外注比率(基本情報31行)
'                     軌道外注比率: 1社目F31 / 2社目I31 …(3列ずつ右)
'                     基本情報C23=溶接工事ありの時、軌道工事ブロックの31行目が空欄なら60.1%を既定入力(手入力は尊重)
'                     1行目: 昼列=「外注比率＝」/ 夜列=基本情報の当社列31行目(F31等)参照
'                     丸め: 整数部4桁以上は上位3桁＋以降0埋め(切り捨て3桁)、3桁以下はROUNDDOWN
'              4行目ヘッダー=(回数)(年度)溶接手元単価
' ・手元比率 : マスタデータ\レール溶接_軌道会社外注費率一覧*.xlsx の
'              「溶接手元割合」シートから整理番号で参照し、数式へ数値リテラルで埋め込む
' ・外注比率 : 基本情報シートのセル参照として数式に残す(既存ロジックと同様)
' ・丸め     : 有効3桁 ROUND(x,-INT(LOG10(x))+2) ※試算シートの丸めと同等
' ・パック工種(整理番号5000番台)は計算対象外(グレー塗り)
' ・書式/罫線/会社名結合は mod_VendorMaster の単価シート作成ロジックと同一仕様
'
' 想定エントリポイント:
'   ApplyWeldingVendorUnitPricesForBasicInfo
'     - mod_MaterialPriceImport.ImportUnitPriceData の溶接単価シート作成直後に呼び出す
'     - Alt+F8 から手動実行も可能
' =====================================================================

' --- 基本情報シート 業者ブロック ---
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10    ' 工事種別(軌道工事/溶接工事)
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11         ' 会社名
Private Const BASIC_INFO_RAIL_RATIO_ROW As Long = 29          ' 軌道工事外注比率
Private Const BASIC_INFO_WELDING_RATIO_ROW As Long = 31       ' 溶接工事外注比率
Private Const BASIC_INFO_WELDING_FLAG_ROW As Long = 23        ' 溶接工事有無フラグ(C23)
Private Const BASIC_INFO_WELDING_FLAG_COL As Long = 3         ' C列
Private Const WUP_RAIL_FIXED_RATIO As Double = 0.601          ' 軌道工事ブロック31行目へ入力する固定比率(60.1%)
Private Const BASIC_INFO_VENDOR_BLOCK_VALUE_COL As Long = 6   ' F列(1社目)
Private Const BASIC_INFO_VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const MAX_VENDOR_BLOCK_COUNT As Long = 20

' --- レール溶接単価シート レイアウト ---
Private Const WUP_RATIO_ROW As Long = 1           ' 外注比率表示行
Private Const WUP_HEADER_ROW As Long = 4          ' 「(回数)(年度)外注単価」結合ヘッダー行
Private Const WUP_NAME_ROW As Long = 5            ' 会社名結合行
Private Const WUP_LABEL_ROW As Long = 6           ' 昼間/夜間ラベル行
Private Const WUP_DATA_START_ROW As Long = 7
Private Const WUP_SEIRI_COL As Long = 2           ' B列 整理番号
Private Const WUP_WORK_NAME_COL As Long = 3       ' C列 工種名
Private Const WUP_JR_DAY_COL As Long = 5          ' E列 JR単価(昼)
Private Const WUP_JR_NIGHT_COL As Long = 6        ' F列 JR単価(夜)
Private Const WUP_WELDING_DAY_COL As Long = 7     ' G列 溶接会社(昼)
Private Const WUP_FIRST_RAIL_DAY_COL As Long = 9  ' I列 軌道会社1社目(昼)
Private Const WUP_RAIL_JR_FACTOR As String = "(100/100.7)"  ' 軌道会社: JR単価に掛ける係数(AG5)
Private Const WUP_PATTERN_ROW As Long = 3              ' 外注費算出パターン行
Private Const WUP_PATTERN_LABEL_COL_FIRST As Long = 4  ' D列(ラベル結合 開始)
Private Const WUP_PATTERN_LABEL_COL_LAST As Long = 5   ' E列(ラベル結合 終了)
Private Const WUP_PATTERN_SELECT_COL As Long = 6       ' F列(パターン選択ドロップダウン)
Private Const WUP_PACK_SEIRI_MIN As Long = 5000   ' パック工種の整理番号下限
' パック構成(マスタ): 0始まりフィールド = Excel列-1。J=9,K=10,L=11,M=12,N=13,O=14
Private Const WUP_MASTER_PACK_COMP1_FIELD As Long = 9   ' J列 構成1の整理番号
Private Const WUP_MASTER_PACK_QTY1_FIELD As Long = 10   ' K列 構成1の数量
Private Const WUP_MASTER_PACK_COMP2_FIELD As Long = 11  ' L列 構成2の整理番号
Private Const WUP_MASTER_PACK_QTY2_FIELD As Long = 12   ' M列 構成2の数量
Private Const WUP_MASTER_PACK_COMP3_FIELD As Long = 13  ' N列 構成3の整理番号
Private Const WUP_MASTER_PACK_QTY3_FIELD As Long = 14   ' O列 構成3の数量
Private Const WUP_NUMBER_FORMAT As String = "#,##0"
Private Const WUP_RATIO_NUMBER_FORMAT As String = "0.0%"
Private Const WUP_RATIO_FONT_SIZE As Long = 11
Private Const WUP_FILL_COLOR_R As Long = 128
Private Const WUP_FILL_COLOR_G As Long = 128
Private Const WUP_FILL_COLOR_B As Long = 128

Private Type WeldingVendorBlock
    valueColumn As Long       ' 基本情報シート上の値列(F/I/L...)
    vendorName As String
    ratioAddress As String    ' '基本情報'!$L$31 等(数式参照用)
    ratioPercent As Variant   ' 0～1 正規化済み(表示用)
    hasRatio As Boolean
End Type

' パック工種計算用の参照マップ(実行中のみ使用)
'   mPackMap     : 整理番号(5000+) -> 構成Collection(各要素 Array(構成整理番号As String, 数量As Double))。実行毎に設定。
'   mSeiriRowMap : 整理番号 -> 当該シートの行番号。シート毎に設定。
Private mPackMap As Object
Private mSeiriRowMap As Object

' =====================================================================
' Public エントリポイント
' =====================================================================

' ブック内の全「_レール溶接単価」シートへ施工会社別単価を展開する
'   preferredRatioColumn: 基本情報シートで31行目(溶接外注比率)が変更された列。
'                         指定された場合、その列のブロックをG/H列の比率参照先として優先採用する。
Public Sub ApplyWeldingVendorUnitPricesForBasicInfo(Optional ByVal wsInfo As Worksheet, _
                                                    Optional ByVal showWarnings As Boolean = False, _
                                                    Optional ByVal preferredRatioColumn As Long = 0)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim weldingSheets As Collection
    Set weldingSheets = CollectWeldingUnitPriceSheets(targetBook)
    If weldingSheets.Count = 0 Then
        LogWUP "レール溶接単価シートなし -> スキップ"
        Exit Sub
    End If

    Dim warningTexts As Collection
    Set warningTexts = New Collection

    ' --- 手元比率マスタの読み込み ---
    Dim temotoMap As Object
    Dim loadErrorText As String
    Set temotoMap = LoadTemotoRatioMap(loadErrorText)
    If temotoMap Is Nothing Then
        LogWUP "手元比率マスタ読込失敗: " & loadErrorText
        If showWarnings Then MsgBox "手元比率マスタを読み込めませんでした。" & vbCrLf & loadErrorText, vbExclamation
        Exit Sub
    End If
    LogWUP "手元比率マスタ読込完了 件数=" & CStr(temotoMap.Count)

    ' --- 基本情報の業者ブロック分類とG/H列(溶接施工会社)の決定 ---
    Dim weldingBlock As WeldingVendorBlock
    Dim weldingNameSources As Collection
    Dim railBlocks() As WeldingVendorBlock
    Dim railBlockCount As Long
    Set weldingNameSources = New Collection
    ScanVendorBlocks wsInfo, weldingBlock, weldingNameSources, railBlocks, railBlockCount, preferredRatioColumn

    If Not weldingBlock.hasRatio Then
        warningTexts.Add WarnWeldingBlockText(weldingBlock)
    End If
    If railBlockCount = 0 Then
        warningTexts.Add "基本情報に「軌道工事」の業者ブロックが見つかりません。"
    End If

    ' --- 基本情報C23=溶接工事ありの場合、軌道工事ブロックの31行目へ固定比率(60.1%)を入力 ---
    ' (ScanVendorBlocks後に実行。走査時の溶接ブロック判定には影響させない)
    Dim railRatioWritten As Boolean
    railRatioWritten = WriteRailOutsourceRatioToBasicInfo(wsInfo)
    LogWUP "軌道31行目比率(60.1%)入力=" & CStr(railRatioWritten)

    ' 軌道会社ブロックの比率参照を29行目→31行目へ切替(60.1%入力後に再構築)。
    ' これにより数式・1行目とも基本情報!$<列>$31 を参照する。
    Dim rbIndex As Long
    For rbIndex = 1 To railBlockCount
        railBlocks(rbIndex) = BuildVendorBlock(wsInfo, railBlocks(rbIndex).valueColumn, _
                                               BASIC_INFO_WELDING_RATIO_ROW)
    Next rbIndex

    Dim missingSeiriMap As Object
    Set missingSeiriMap = CreateObject("Scripting.Dictionary")

    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Cleanup

    Dim wsWelding As Variant
    For Each wsWelding In weldingSheets
        ApplyWeldingVendorUnitPricesToSheet wsWelding, wsInfo, weldingBlock, weldingNameSources, _
                                            railBlocks, railBlockCount, temotoMap, missingSeiriMap
    Next wsWelding

    If missingSeiriMap.Count > 0 Then
        warningTexts.Add BuildMissingSeiriWarningText(missingSeiriMap)
    End If

Cleanup:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    On Error Resume Next
    targetBook.Calculate
    On Error GoTo 0

    If showWarnings And warningTexts.Count > 0 Then
        MsgBox JoinCollectionText(warningTexts, vbCrLf & vbCrLf), vbExclamation, "レール溶接単価"
    End If
End Sub

' 工事種別(10行目)が変わらない会社名変更時は、5行目の会社名表示だけを更新する。
' 全行展開(ApplyWeldingVendorUnitPricesForBasicInfo)は数十秒かかるため、名称変更では呼ばない。
Public Sub UpdateWeldingVendorDisplayNamesForBasicInfo(Optional ByVal wsInfo As Worksheet, _
                                                       Optional ByVal preferredRatioColumn As Long = 0)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim weldingSheets As Collection
    Set weldingSheets = CollectWeldingUnitPriceSheets(targetBook)
    If weldingSheets.Count = 0 Then Exit Sub

    Dim weldingBlock As WeldingVendorBlock
    Dim weldingNameSources As Collection
    Dim railBlocks() As WeldingVendorBlock
    Dim railBlockCount As Long
    Set weldingNameSources = New Collection
    ScanVendorBlocks wsInfo, weldingBlock, weldingNameSources, railBlocks, railBlockCount, preferredRatioColumn

    Dim vendorUnitPriceNameMap As Object
    Set vendorUnitPriceNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)

    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Dim wsWelding As Variant
    For Each wsWelding In weldingSheets
        On Error Resume Next
        UpdateWeldingVendorDisplayNamesOnSheet wsWelding, weldingBlock, weldingNameSources, _
            railBlocks, railBlockCount, vendorUnitPriceNameMap
        If Err.Number <> 0 Then
            LogWUP "表示名更新スキップ sheet=[" & wsWelding.Name & "] Err=" & CStr(Err.Number)
            Err.Clear
        End If
        On Error GoTo Cleanup
    Next wsWelding

Cleanup:
    Application.ScreenUpdating = previousScreenUpdating
End Sub

Private Sub UpdateWeldingVendorDisplayNamesOnSheet(ByVal wsWelding As Worksheet, _
                                                  ByRef weldingBlock As WeldingVendorBlock, _
                                                  ByVal weldingNameSources As Collection, _
                                                  ByRef railBlocks() As WeldingVendorBlock, _
                                                  ByVal railBlockCount As Long, _
                                                  ByVal vendorUnitPriceNameMap As Object)
    If weldingBlock.valueColumn > 0 And weldingBlock.hasRatio Then
        UpdateWeldingVendorDisplayNameOnly wsWelding, WUP_WELDING_DAY_COL, _
            BuildWeldingDisplayName(weldingNameSources, vendorUnitPriceNameMap)
    End If

    Dim railIndex As Long
    For railIndex = 1 To railBlockCount
        Dim railDayCol As Long
        railDayCol = WUP_FIRST_RAIL_DAY_COL + ((railIndex - 1) * 2)
        If railBlocks(railIndex).hasRatio Then
            UpdateWeldingVendorDisplayNameOnly wsWelding, railDayCol, _
                ResolveVendorUnitPriceNameWUP(vendorUnitPriceNameMap, railBlocks(railIndex).vendorName)
        End If
    Next railIndex
End Sub

Private Sub UpdateWeldingVendorDisplayNameOnly(ByVal wsWelding As Worksheet, _
                                               ByVal dayCol As Long, _
                                               ByVal displayName As String)
    ApplyMergedCell wsWelding, WUP_NAME_ROW, dayCol, dayCol + 1, displayName, True
End Sub

' 指定した基本情報列( F/I/L… )に対応する溶接単価列だけを展開する。
' 2社目の初回入力など、全シート全列の再展開を避けるために使用する。
Public Sub ApplyWeldingVendorUnitPricesForBasicInfoColumns(ByVal wsInfo As Worksheet, _
                                                           ByVal targetValueColumns As Collection, _
                                                           Optional ByVal preferredRatioColumn As Long = 0)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    If targetValueColumns Is Nothing Then Exit Sub
    If targetValueColumns.Count = 0 Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim weldingSheets As Collection
    Set weldingSheets = CollectWeldingUnitPriceSheets(targetBook)
    If weldingSheets.Count = 0 Then Exit Sub

    Dim temotoMap As Object
    Dim loadErrorText As String
    Set temotoMap = LoadTemotoRatioMap(loadErrorText)
    If temotoMap Is Nothing Then
        LogWUP "手元比率マスタ読込失敗: " & loadErrorText
        Exit Sub
    End If

    Dim weldingBlock As WeldingVendorBlock
    Dim weldingNameSources As Collection
    Dim railBlocks() As WeldingVendorBlock
    Dim railBlockCount As Long
    Set weldingNameSources = New Collection
    ScanVendorBlocks wsInfo, weldingBlock, weldingNameSources, railBlocks, railBlockCount, preferredRatioColumn

    WriteRailOutsourceRatioToBasicInfo wsInfo

    Dim rbIndex As Long
    For rbIndex = 1 To railBlockCount
        railBlocks(rbIndex) = BuildVendorBlock(wsInfo, railBlocks(rbIndex).valueColumn, _
                                               BASIC_INFO_WELDING_RATIO_ROW)
    Next rbIndex

    Dim missingSeiriMap As Object
    Set missingSeiriMap = CreateObject("Scripting.Dictionary")

    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Cleanup

    Dim wsWelding As Variant
    For Each wsWelding In weldingSheets
        ApplyWeldingVendorUnitPricesToSheetColumns wsWelding, wsInfo, weldingBlock, weldingNameSources, _
            railBlocks, railBlockCount, temotoMap, missingSeiriMap, targetValueColumns
    Next wsWelding

Cleanup:
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    On Error Resume Next
    targetBook.Calculate
    On Error GoTo 0
End Sub

Private Sub ApplyWeldingVendorUnitPricesToSheetColumns(ByVal wsWelding As Worksheet, _
                                                       ByVal wsInfo As Worksheet, _
                                                       ByRef weldingBlock As WeldingVendorBlock, _
                                                       ByVal weldingNameSources As Collection, _
                                                       ByRef railBlocks() As WeldingVendorBlock, _
                                                       ByVal railBlockCount As Long, _
                                                       ByVal temotoMap As Object, _
                                                       ByVal missingSeiriMap As Object, _
                                                       ByVal targetValueColumns As Collection)
    Dim lastRow As Long
    lastRow = wsWelding.Cells(wsWelding.Rows.Count, WUP_SEIRI_COL).End(xlUp).Row
    If lastRow < WUP_DATA_START_ROW Then Exit Sub

    SetupOutsourcePatternSelector wsWelding
    Set mSeiriRowMap = BuildSeiriRowMapWUP(wsWelding, lastRow)

    Dim outsourceHeaderText As String
    Dim railHeaderText As String
    outsourceHeaderText = BuildWeldingUnitPriceHeaderText(wsInfo, True)
    railHeaderText = BuildWeldingUnitPriceHeaderText(wsInfo, False)

    Dim vendorUnitPriceNameMap As Object
    Set vendorUnitPriceNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)

    If weldingBlock.valueColumn > 0 And weldingBlock.hasRatio Then
        If CollectionContainsLongWUP(targetValueColumns, weldingBlock.valueColumn) Then
            ApplyWeldingVendorBlock wsWelding, lastRow, WUP_WELDING_DAY_COL, weldingBlock, True, _
                                    outsourceHeaderText, vendorUnitPriceNameMap, temotoMap, missingSeiriMap, _
                                    BuildWeldingDisplayName(weldingNameSources, vendorUnitPriceNameMap)
        Else
            UpdateWeldingVendorDisplayNameOnly wsWelding, WUP_WELDING_DAY_COL, _
                BuildWeldingDisplayName(weldingNameSources, vendorUnitPriceNameMap)
        End If
    End If

    Dim railIndex As Long
    For railIndex = 1 To railBlockCount
        Dim railDayCol As Long
        railDayCol = WUP_FIRST_RAIL_DAY_COL + ((railIndex - 1) * 2)
        If CollectionContainsLongWUP(targetValueColumns, railBlocks(railIndex).valueColumn) Then
            If railBlocks(railIndex).hasRatio Then
                ApplyWeldingVendorBlock wsWelding, lastRow, railDayCol, railBlocks(railIndex), False, _
                                        railHeaderText, vendorUnitPriceNameMap, temotoMap, missingSeiriMap
            Else
                ClearWeldingVendorBlock wsWelding, lastRow, railDayCol
            End If
        End If
    Next railIndex

    LogWUP "部分展開完了 sheet=[" & wsWelding.Name & "] 対象列数=" & CStr(targetValueColumns.Count)
End Sub

Private Function CollectionContainsLongWUP(ByVal values As Collection, ByVal targetValue As Long) As Boolean
    Dim item As Variant
    For Each item In values
        If CLng(item) = targetValue Then
            CollectionContainsLongWUP = True
            Exit Function
        End If
    Next item
End Function

' シート名が「_レール溶接単価」を含むかどうか
Public Function IsWeldingUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function
    IsWeldingUnitPriceSheet = _
        (InStr(1, NormalizeMatchTextWUP(CStr(targetSheet.Name)), _
               NormalizeMatchTextWUP(WeldingSheetSuffixText()), vbTextCompare) > 0)
End Function

' =====================================================================
' シート単位の展開処理
' =====================================================================

Private Sub ApplyWeldingVendorUnitPricesToSheet(ByVal wsWelding As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByRef weldingBlock As WeldingVendorBlock, _
                                                ByVal weldingNameSources As Collection, _
                                                ByRef railBlocks() As WeldingVendorBlock, _
                                                ByVal railBlockCount As Long, _
                                                ByVal temotoMap As Object, _
                                                ByVal missingSeiriMap As Object)
    Dim lastRow As Long
    lastRow = wsWelding.Cells(wsWelding.Rows.Count, WUP_SEIRI_COL).End(xlUp).Row
    If lastRow < WUP_DATA_START_ROW Then
        LogWUP "データ行なし sheet=[" & wsWelding.Name & "]"
        Exit Sub
    End If

    ' D3:E3=「外注費算出パターン：」/ F3=パターン選択ドロップダウン を整備
    SetupOutsourcePatternSelector wsWelding

    ' パック工種(5000番台)が構成工種の行を参照できるよう、整理番号->行 を構築
    Set mSeiriRowMap = BuildSeiriRowMapWUP(wsWelding, lastRow)

    Dim outsourceHeaderText As String
    Dim railHeaderText As String
    outsourceHeaderText = BuildWeldingUnitPriceHeaderText(wsInfo, True)
    railHeaderText = BuildWeldingUnitPriceHeaderText(wsInfo, False)

    Dim vendorUnitPriceNameMap As Object
    Set vendorUnitPriceNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)

    ' --- 溶接施工会社ブロック (G/H列固定) ---
    ' 会社名は業者マスタ(B6支店シート)のB列一致→A列値へ解決し、複数社は「・」で結合して表示
    If weldingBlock.valueColumn > 0 And weldingBlock.hasRatio Then
        ApplyWeldingVendorBlock wsWelding, lastRow, WUP_WELDING_DAY_COL, weldingBlock, True, _
                                outsourceHeaderText, vendorUnitPriceNameMap, temotoMap, missingSeiriMap, _
                                BuildWeldingDisplayName(weldingNameSources, vendorUnitPriceNameMap)
    Else
        ClearWeldingVendorBlock wsWelding, lastRow, WUP_WELDING_DAY_COL
    End If

    ' --- 軌道会社ブロック (I/J列から1社2列ずつ) ---
    Dim railIndex As Long
    For railIndex = 1 To railBlockCount
        Dim railDayCol As Long
        railDayCol = WUP_FIRST_RAIL_DAY_COL + ((railIndex - 1) * 2)
        If railBlocks(railIndex).hasRatio Then
            ApplyWeldingVendorBlock wsWelding, lastRow, railDayCol, railBlocks(railIndex), False, _
                                    railHeaderText, vendorUnitPriceNameMap, temotoMap, missingSeiriMap
        Else
            ClearWeldingVendorBlock wsWelding, lastRow, railDayCol
        End If
    Next railIndex

    ' --- 余剰スロットのクリア(業者数減少時の残骸対策) ---
    For railIndex = railBlockCount + 1 To MAX_VENDOR_BLOCK_COUNT
        ClearWeldingVendorBlock wsWelding, lastRow, WUP_FIRST_RAIL_DAY_COL + ((railIndex - 1) * 2)
    Next railIndex

    LogWUP "展開完了 sheet=[" & wsWelding.Name & "] 軌道会社数=" & CStr(railBlockCount)
End Sub

Private Sub ApplyWeldingVendorBlock(ByVal wsWelding As Worksheet, _
                                    ByVal lastRow As Long, _
                                    ByVal dayCol As Long, _
                                    ByRef block As WeldingVendorBlock, _
                                    ByVal isWeldingVendor As Boolean, _
                                    ByVal headerText As String, _
                                    ByVal vendorUnitPriceNameMap As Object, _
                                    ByVal temotoMap As Object, _
                                    ByVal missingSeiriMap As Object, _
                                    Optional ByVal displayNameOverride As String = "")
    Dim nightCol As Long
    nightCol = dayCol + 1

    ' 再実行に備えて一旦初期化
    ClearWeldingVendorBlock wsWelding, lastRow, dayCol

    ' 列幅 = E/F列に合わせる(既存ロジック同様)
    wsWelding.Columns(dayCol).ColumnWidth = wsWelding.Columns(WUP_JR_DAY_COL).ColumnWidth
    wsWelding.Columns(nightCol).ColumnWidth = wsWelding.Columns(WUP_JR_NIGHT_COL).ColumnWidth

    ' 1行目: 溶接施工会社=外注比率表示 / 軌道会社=「外注比率＝」＋基本情報31行目(F31等)参照
    If isWeldingVendor Then
        ApplyOutsourceRatioRow wsWelding, dayCol, nightCol, block.ratioPercent
    Else
        ApplyRailMarkupRatioRow wsWelding, dayCol, nightCol, block.ratioAddress
    End If

    ' 4行目: 結合ヘッダー / 5行目: 結合会社名 / 6行目: 昼間・夜間
    Dim displayName As String
    displayName = displayNameOverride
    If Len(displayName) = 0 Then
        displayName = ResolveVendorUnitPriceNameWUP(vendorUnitPriceNameMap, block.vendorName)
    End If

    ApplyMergedCell wsWelding, WUP_HEADER_ROW, dayCol, nightCol, headerText, False
    ApplyMergedCell wsWelding, WUP_NAME_ROW, dayCol, nightCol, displayName, True
    With wsWelding.Cells(WUP_LABEL_ROW, dayCol)
        .Value = DayLabelText()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With wsWelding.Cells(WUP_LABEL_ROW, nightCol)
        .Value = NightLabelText()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' データ行
    Dim rowIndex As Long
    For rowIndex = WUP_DATA_START_ROW To lastRow
        ApplyWeldingVendorRow wsWelding, rowIndex, dayCol, nightCol, block, _
                              isWeldingVendor, temotoMap, missingSeiriMap
    Next rowIndex

    ' フォント(BIZ UDゴシック)と罫線(細格子+太外枠)
    ApplyWeldingVendorFont wsWelding, lastRow, dayCol, nightCol
    ApplyWeldingVendorBorders wsWelding, lastRow, dayCol, nightCol
End Sub

Private Sub ApplyWeldingVendorRow(ByVal wsWelding As Worksheet, _
                                  ByVal rowIndex As Long, _
                                  ByVal dayCol As Long, _
                                  ByVal nightCol As Long, _
                                  ByRef block As WeldingVendorBlock, _
                                  ByVal isWeldingVendor As Boolean, _
                                  ByVal temotoMap As Object, _
                                  ByVal missingSeiriMap As Object)
    Dim seiriText As String
    seiriText = Trim$(CStr(wsWelding.Cells(rowIndex, WUP_SEIRI_COL).Value))
    If Len(seiriText) = 0 Then
        ApplyGreyFill wsWelding.Cells(rowIndex, dayCol)
        ApplyGreyFill wsWelding.Cells(rowIndex, nightCol)
        Exit Sub
    End If

    Dim seiriNumber As Long
    seiriNumber = CLng(Val(StrConv(seiriText, vbNarrow)))

    ' パック工種(5000番台): 構成工種の同列セル×数量の合計を算出
    '   溶接(G/H)=3工種(J/K,L/M,N/O) / 軌道(I/J等)=1工種(J/K)のみ
    If seiriNumber >= WUP_PACK_SEIRI_MIN Then
        ApplyPackRowForBlock wsWelding, rowIndex, dayCol, nightCol, CStr(seiriNumber), isWeldingVendor
        Exit Sub
    End If

    ' 産廃処理は対象外(既存ロジック踏襲)
    Dim workTypeName As String
    workTypeName = NormalizeMatchTextWUP(CStr(wsWelding.Cells(rowIndex, WUP_WORK_NAME_COL).Value))
    If InStr(1, workTypeName, WasteDisposalKeywordText(), vbTextCompare) > 0 Then
        ApplyGreyFill wsWelding.Cells(rowIndex, dayCol)
        ApplyGreyFill wsWelding.Cells(rowIndex, nightCol)
        Exit Sub
    End If

    ' 手元割合を整理番号で参照
    ' (溶接単価シートB列の整理番号 = マスタ溶接手元割合シートA列の整理番号 で照合)
    Dim seiriKey As String
    seiriKey = CStr(seiriNumber)
    If Not temotoMap.Exists(seiriKey) Then
        ApplyGreyFill wsWelding.Cells(rowIndex, dayCol)
        ApplyGreyFill wsWelding.Cells(rowIndex, nightCol)
        If Not missingSeiriMap.Exists(seiriKey) Then missingSeiriMap.Add seiriKey, True
        Exit Sub
    End If

    Dim temotoPair As Variant
    temotoPair = temotoMap(seiriKey)   ' Array(昼E, 夜F) ※未設定はEmpty

    If isWeldingVendor Then
        ApplyWeldingVendorCell wsWelding.Cells(rowIndex, dayCol), wsWelding, rowIndex, _
                               WUP_JR_DAY_COL, temotoPair(0), block.ratioAddress, isWeldingVendor
        ApplyWeldingVendorCell wsWelding.Cells(rowIndex, nightCol), wsWelding, rowIndex, _
                               WUP_JR_NIGHT_COL, temotoPair(1), block.ratioAddress, isWeldingVendor
    Else
        ' 軌道会社: (JR×100/100.7)×手元割合(昼E/夜F)×軌道外注比率(基本情報31行)
        ApplyRailMarkupCell wsWelding.Cells(rowIndex, dayCol), wsWelding, rowIndex, _
                            WUP_JR_DAY_COL, temotoPair(0), block.ratioAddress
        ApplyRailMarkupCell wsWelding.Cells(rowIndex, nightCol), wsWelding, rowIndex, _
                            WUP_JR_NIGHT_COL, temotoPair(1), block.ratioAddress
    End If
End Sub

Private Sub ApplyWeldingVendorCell(ByVal targetCell As Range, _
                                   ByVal wsWelding As Worksheet, _
                                   ByVal rowIndex As Long, _
                                   ByVal sourceCol As Long, _
                                   ByVal temotoRatio As Variant, _
                                   ByVal ratioAddress As String, _
                                   ByVal isWeldingVendor As Boolean)
    With targetCell
        .ShrinkToFit = False
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' JR単価が空欄 or 手元比率未設定 -> グレー塗り
    If Len(Trim$(CStr(wsWelding.Cells(rowIndex, sourceCol).Value))) = 0 Then
        ApplyGreyFill targetCell
        Exit Sub
    End If
    If Not IsNumeric(temotoRatio) Then
        ApplyGreyFill targetCell
        Exit Sub
    End If

    targetCell.Formula = BuildWeldingVendorFormula(wsWelding, rowIndex, sourceCol, _
                                                   CDbl(temotoRatio), ratioAddress, isWeldingVendor)
    targetCell.NumberFormat = WUP_NUMBER_FORMAT
End Sub

' 数式組み立て:
'   溶接会社: =IFERROR(ROUND(E7*(1-0.0791)*('基本情報'!$L$31),-INT(LOG10(E7*(1-0.0791)*('基本情報'!$L$31)))+2),0)
'   軌道会社: =IFERROR(ROUND(E7*0.0791*('基本情報'!$F$29),-INT(LOG10(E7*0.0791*('基本情報'!$F$29)))+2),0)
' ※有効3桁丸め。手元比率0の工種(付帯作業等)で軌道側が0になるためIFERRORで0を返す。
Private Function BuildWeldingVendorFormula(ByVal wsWelding As Worksheet, _
                                           ByVal rowIndex As Long, _
                                           ByVal sourceCol As Long, _
                                           ByVal temotoRatio As Double, _
                                           ByVal ratioAddress As String, _
                                           ByVal isWeldingVendor As Boolean) As String
    Dim unitCellRef As String
    unitCellRef = wsWelding.Cells(rowIndex, sourceCol).Address(False, False)

    Dim factorText As String
    If isWeldingVendor Then
        factorText = "(1-" & RatioLiteralText(temotoRatio) & ")"
    Else
        factorText = RatioLiteralText(temotoRatio)
    End If

    Dim exprText As String
    exprText = unitCellRef & "*" & factorText & "*(" & ratioAddress & ")"

    BuildWeldingVendorFormula = "=IFERROR(ROUND(" & exprText & _
                                ",-INT(LOG10(" & exprText & "))+2),0)"
End Function

' Double を Excel 数式用の数値リテラルへ(小数点は常にピリオド)
Private Function RatioLiteralText(ByVal value As Double) As String
    Dim s As String
    s = Trim$(Str$(value))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    RatioLiteralText = s
End Function

' =====================================================================
' 軌道会社列(markup方式)の数式・セル書き込み
' =====================================================================

' 軌道会社セルへ新単価数式を書き込む(溶接会社セルと同様にグレー塗り判定を行う)。
'   sourceCol  : JR単価列(昼=E/夜=F)
'   temotoRatio: マスタ溶接手元割合の値(昼E/夜F) ※リテラル
'   ratioAddress: 軌道外注比率(基本情報31行 1社目F31/2社目I31…)のセル参照
Private Sub ApplyRailMarkupCell(ByVal targetCell As Range, _
                                ByVal wsWelding As Worksheet, _
                                ByVal rowIndex As Long, _
                                ByVal sourceCol As Long, _
                                ByVal temotoRatio As Variant, _
                                ByVal ratioAddress As String)
    With targetCell
        .ShrinkToFit = False
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' JR単価が空欄 or 手元割合未設定 -> グレー塗り
    If Len(Trim$(CStr(wsWelding.Cells(rowIndex, sourceCol).Value))) = 0 Then
        ApplyGreyFill targetCell
        Exit Sub
    End If
    If Not IsNumeric(temotoRatio) Then
        ApplyGreyFill targetCell
        Exit Sub
    End If

    targetCell.Formula = BuildRailMarkupFormula(wsWelding, rowIndex, sourceCol, _
                                                CDbl(temotoRatio), ratioAddress)
    targetCell.NumberFormat = WUP_NUMBER_FORMAT
End Sub

' 軌道会社単価の数式。F3(外注費算出パターン)の選択値で計算方式を切替える。
'   共通: jr=JR単価(E/F) / lit=手元割合(マスタ昼E/夜F リテラル) / R=軌道外注比率(基本情報31行)
'   ■物価指数適用 : v=(jr×(100/100.7)×lit)×R を 整数部4桁以上は上位3桁＋0埋め(切り捨て3桁)、
'                   3桁以下はROUNDDOWNで整数化。
'   ■外注比率適用 : Y=lit×R, v=Y×jr を 桁数3以下はそのまま、4桁以上は上位4桁ROUND(,-1)で桁戻し。
'                   (ユーザー指定式: =IF(J="","",IF(LEN(TEXT(Y*J,"#"))<=3,Y*J,
'                     ROUND(VALUE(LEFT(TEXT(Y*J,"#"),4)),-1)*10^VALUE(LEN(TEXT(Y*J,"#"))-4))))
'   ■前年度単価適用: 未定義のため空欄("")。
' F3を切り替えるとExcelの再計算でそのまま単価が切り替わる。
Private Function BuildRailMarkupFormula(ByVal wsWelding As Worksheet, _
                                        ByVal rowIndex As Long, _
                                        ByVal sourceCol As Long, _
                                        ByVal temotoRatio As Double, _
                                        ByVal ratioAddress As String) As String
    Dim q As String
    q = Chr$(34)

    Dim jrRef As String
    jrRef = wsWelding.Cells(rowIndex, sourceCol).Address(False, False)
    Dim lit As String
    lit = RatioLiteralText(temotoRatio)

    Dim f3Ref As String
    f3Ref = wsWelding.Cells(WUP_PATTERN_ROW, WUP_PATTERN_SELECT_COL).Address(True, True)  ' $F$3

    ' --- 物価指数適用 ---
    Dim coreBukka As String
    coreBukka = "(" & jrRef & "*" & WUP_RAIL_JR_FACTOR & "*" & lit & ")*" & ratioAddress
    Dim txBukka As String
    txBukka = "TEXT((" & coreBukka & "),0)"
    Dim exprBukka As String
    exprBukka = "IF(VALUE(LEN(" & txBukka & "))>3," & _
                "VALUE(MID(" & txBukka & ",1,3)&REPT(0,LEN(" & txBukka & ")-3))," & _
                "(ROUNDDOWN(" & coreBukka & ",0)))"

    ' --- 外注比率適用 ---  Y=lit×R, v=Y×jr
    Dim prodGaibu As String
    prodGaibu = "(" & lit & "*" & ratioAddress & ")*" & jrRef
    Dim txGaibu As String
    txGaibu = "TEXT(" & prodGaibu & "," & q & "#" & q & ")"
    Dim exprGaibu As String
    exprGaibu = "IF(" & jrRef & "=" & q & q & "," & q & q & "," & _
                "IF(LEN(" & txGaibu & ")<=3," & prodGaibu & "," & _
                "ROUND(VALUE(LEFT(" & txGaibu & ",4)),-1)*10^VALUE(LEN(" & txGaibu & ")-4)))"

    ' --- F3で分岐(前年度単価適用は未定義のため空欄) ---
    BuildRailMarkupFormula = _
        "=IF(" & f3Ref & "=" & q & PatternOutsourceRatioText() & q & "," & exprGaibu & "," & _
        "IF(" & f3Ref & "=" & q & PatternPriceIndexText() & q & "," & exprBukka & "," & _
        q & q & "))"
End Function

' =====================================================================
' パック工種(5000番台)の計算
' =====================================================================

' 当該シートの 整理番号 -> 行番号 マップ(パック構成工種の行参照用)
Private Function BuildSeiriRowMapWUP(ByVal wsWelding As Worksheet, ByVal lastRow As Long) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare

    Dim r As Long, t As String
    For r = WUP_DATA_START_ROW To lastRow
        t = Trim$(StrConv(CStr(wsWelding.Cells(r, WUP_SEIRI_COL).Value), vbNarrow))
        If Len(t) > 0 And IsNumeric(t) Then
            Dim k As String
            k = CStr(CLng(Val(t)))
            If Not m.Exists(k) Then m.Add k, r
        End If
    Next r
    Set BuildSeiriRowMapWUP = m
End Function

' パック行(dayCol/nightCol)へ構成工種の合計式を書き込む。
' 構成や行が解決できない場合はグレー塗り。
Private Sub ApplyPackRowForBlock(ByVal wsWelding As Worksheet, _
                                 ByVal rowIndex As Long, _
                                 ByVal dayCol As Long, _
                                 ByVal nightCol As Long, _
                                 ByVal packKey As String, _
                                 ByVal isWeldingVendor As Boolean)
    If mPackMap Is Nothing Or mSeiriRowMap Is Nothing Or Not mPackMap.Exists(packKey) Then
        ApplyGreyFill wsWelding.Cells(rowIndex, dayCol)
        ApplyGreyFill wsWelding.Cells(rowIndex, nightCol)
        Exit Sub
    End If

    Dim comps As Collection
    Set comps = mPackMap(packKey)

    ApplyPackCell wsWelding.Cells(rowIndex, dayCol), _
                  BuildPackSumFormula(wsWelding, dayCol, comps, isWeldingVendor)
    ApplyPackCell wsWelding.Cells(rowIndex, nightCol), _
                  BuildPackSumFormula(wsWelding, nightCol, comps, isWeldingVendor)
End Sub

' 構成工種の同列セル×数量 の合計式。溶接=全構成、軌道=先頭1工種(J/K)のみ。
Private Function BuildPackSumFormula(ByVal wsWelding As Worksheet, _
                                     ByVal col As Long, _
                                     ByVal comps As Collection, _
                                     ByVal isWeldingVendor As Boolean) As String
    Dim maxComp As Long
    maxComp = comps.Count
    If Not isWeldingVendor Then maxComp = 1          ' 軌道はJ列(1工種)のみ
    If maxComp > comps.Count Then maxComp = comps.Count

    Dim terms As String
    Dim i As Long
    For i = 1 To maxComp
        Dim comp As Variant
        comp = comps(i)                              ' Array(構成整理番号, 数量)
        Dim compKey As String
        compKey = CStr(comp(0))
        If mSeiriRowMap.Exists(compKey) Then
            Dim compRow As Long
            compRow = CLng(mSeiriRowMap(compKey))
            Dim cellRef As String
            cellRef = wsWelding.Cells(compRow, col).Address(False, False)
            If Len(terms) > 0 Then terms = terms & "+"
            terms = terms & cellRef & "*" & RatioLiteralText(CDbl(comp(1)))
        End If
    Next i

    If Len(terms) = 0 Then
        BuildPackSumFormula = ""
    Else
        BuildPackSumFormula = "=" & terms
    End If
End Function

Private Sub ApplyPackCell(ByVal targetCell As Range, ByVal formulaText As String)
    targetCell.Interior.ColorIndex = xlColorIndexNone
    targetCell.ShrinkToFit = False
    If Len(formulaText) = 0 Then
        ApplyGreyFill targetCell                     ' 構成行が解決できない -> グレー
    Else
        targetCell.Formula = formulaText
        targetCell.NumberFormat = WUP_NUMBER_FORMAT
    End If
End Sub

' =====================================================================
' 書式・罫線・クリア(mod_VendorMaster の単価シート作成ロジックと同一仕様)
' =====================================================================

Private Sub ApplyOutsourceRatioRow(ByVal wsWelding As Worksheet, _
                                   ByVal dayCol As Long, _
                                   ByVal nightCol As Long, _
                                   ByVal ratioPercent As Variant)
    With wsWelding.Cells(WUP_RATIO_ROW, dayCol)
        .Formula = ""
        .Value = OutsourceRatioLabelText()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    With wsWelding.Cells(WUP_RATIO_ROW, nightCol)
        .Formula = ""
        If IsNumeric(ratioPercent) Then .Value = CDbl(ratioPercent)
        .NumberFormat = WUP_RATIO_NUMBER_FORMAT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsWelding.Range(wsWelding.Cells(WUP_RATIO_ROW, dayCol), _
                         wsWelding.Cells(WUP_RATIO_ROW, nightCol)).Font
        .Name = WeldingUnitPriceFontNameText()
        .Size = WUP_RATIO_FONT_SIZE
        On Error Resume Next
        .NameFarEast = WeldingUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

' 軌道会社ブロックの1行目: 昼列(I/K…)=マスタH1のラベル文字列「上昇率 β =」/
'                          夜列(J/L…)=マスタI1の率の数値(0.0%表示)。
' データ行の数式は昼夜とも夜列(率数値セル)を参照する。
' 軌道会社ブロックの1行目: 昼列(I/K…)=「外注比率＝」ラベル /
'                          夜列(J/L…)=基本情報シートの当社列31行目(F31/I31…)への参照。
' 会社名(列)と整合するよう ratioAddress(=当ブロックの基本情報列$31) をそのまま参照する。
Private Sub ApplyRailMarkupRatioRow(ByVal wsWelding As Worksheet, _
                                    ByVal dayCol As Long, _
                                    ByVal nightCol As Long, _
                                    ByVal ratioAddress As String)
    ' 昼列1行目: 「外注比率＝」ラベル(溶接会社列と同じ表記)
    With wsWelding.Cells(WUP_RATIO_ROW, dayCol)
        .Formula = ""
        .Value = OutsourceRatioLabelText()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    ' 夜列1行目: 基本情報の当社列31行目への参照(例 ='基本情報'!$F$31)。60.1%等を表示。
    With wsWelding.Cells(WUP_RATIO_ROW, nightCol)
        If Len(ratioAddress) > 0 Then
            .Formula = "=" & ratioAddress
        Else
            .ClearContents
        End If
        .NumberFormat = WUP_RATIO_NUMBER_FORMAT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsWelding.Range(wsWelding.Cells(WUP_RATIO_ROW, dayCol), _
                         wsWelding.Cells(WUP_RATIO_ROW, nightCol)).Font
        .Name = WeldingUnitPriceFontNameText()
        .Size = WUP_RATIO_FONT_SIZE
        On Error Resume Next
        .NameFarEast = WeldingUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Private Sub ApplyMergedCell(ByVal wsWelding As Worksheet, _
                            ByVal rowIndex As Long, _
                            ByVal dayCol As Long, _
                            ByVal nightCol As Long, _
                            ByVal cellText As String, _
                            ByVal useShrinkToFit As Boolean)
    Dim mergeRange As Range
    Set mergeRange = wsWelding.Range(wsWelding.Cells(rowIndex, dayCol), _
                                     wsWelding.Cells(rowIndex, nightCol))

    SafeUnmergeRangeWUP mergeRange
    mergeRange.Merge
    With mergeRange
        .Value = cellText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = useShrinkToFit
        .WrapText = False
    End With
End Sub

' D3:E3=「外注費算出パターン：」(右詰・縮小) / F3=パターン選択ドロップダウン(左詰・縮小)
' F3の選択値で軌道会社の単価計算が切り替わる(各セルの数式がF3を参照して分岐する)。
Private Sub SetupOutsourcePatternSelector(ByVal wsWelding As Worksheet)
    If wsWelding Is Nothing Then Exit Sub
    On Error GoTo CleanupErr

    ' --- D3:E3 ラベル(右詰・縮小表示) ---
    Dim labelRange As Range
    Set labelRange = wsWelding.Range( _
        wsWelding.Cells(WUP_PATTERN_ROW, WUP_PATTERN_LABEL_COL_FIRST), _
        wsWelding.Cells(WUP_PATTERN_ROW, WUP_PATTERN_LABEL_COL_LAST))
    SafeUnmergeRangeWUP labelRange
    labelRange.Merge
    With labelRange
        .Value = OutsourcePatternLabelText()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
        .Font.Name = WeldingUnitPriceFontNameText()
        On Error Resume Next
        .Font.NameFarEast = WeldingUnitPriceFontNameText()
        On Error GoTo CleanupErr
    End With

    ' --- F3 ドロップダウン(左詰・縮小表示) ---
    Dim selectCell As Range
    Set selectCell = wsWelding.Cells(WUP_PATTERN_ROW, WUP_PATTERN_SELECT_COL)
    With selectCell
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
        .Font.Name = WeldingUnitPriceFontNameText()
        On Error Resume Next
        .Font.NameFarEast = WeldingUnitPriceFontNameText()
        On Error GoTo CleanupErr
    End With

    With selectCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:=PatternPrevYearText() & "," & PatternOutsourceRatioText() & "," & PatternPriceIndexText()
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

    ' 既定値: 空欄なら現行ロジック=物価指数適用 を設定
    If Len(Trim$(CStr(selectCell.Value))) = 0 Then
        selectCell.Value = PatternPriceIndexText()
    End If
    Exit Sub

CleanupErr:
    LogWUP "SetupOutsourcePatternSelector: 失敗 sheet=[" & wsWelding.Name & "] err=" & CStr(Err.Number)
End Sub

' 外注費算出パターンのラベル・選択肢(ドロップダウンと数式の比較で同一文字列を使用)
Private Function OutsourcePatternLabelText() As String
    Static cached As String
    If cached = "" Then cached = "外注費算出パターン："
    OutsourcePatternLabelText = cached
End Function

Private Function PatternPrevYearText() As String
    Static cached As String
    If cached = "" Then cached = "前年度単価適用"
    PatternPrevYearText = cached
End Function

Private Function PatternOutsourceRatioText() As String
    Static cached As String
    If cached = "" Then cached = "外注比率適用"
    PatternOutsourceRatioText = cached
End Function

Private Function PatternPriceIndexText() As String
    Static cached As String
    If cached = "" Then cached = "物価指数適用"
    PatternPriceIndexText = cached
End Function

Private Sub ApplyWeldingVendorFont(ByVal wsWelding As Worksheet, _
                                   ByVal lastRow As Long, _
                                   ByVal dayCol As Long, _
                                   ByVal nightCol As Long)
    Dim fontRange As Range
    Set fontRange = wsWelding.Range(wsWelding.Cells(WUP_HEADER_ROW, dayCol), _
                                    wsWelding.Cells(lastRow, nightCol))
    With fontRange.Font
        .Name = WeldingUnitPriceFontNameText()
        On Error Resume Next
        .NameFarEast = WeldingUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Private Sub ApplyWeldingVendorBorders(ByVal wsWelding As Worksheet, _
                                      ByVal lastRow As Long, _
                                      ByVal dayCol As Long, _
                                      ByVal nightCol As Long)
    Dim borderRange As Range
    Set borderRange = wsWelding.Range(wsWelding.Cells(WUP_HEADER_ROW, dayCol), _
                                      wsWelding.Cells(lastRow, nightCol))
    With borderRange
        .Borders.LineStyle = xlNone
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Private Sub ClearWeldingVendorBlock(ByVal wsWelding As Worksheet, _
                                    ByVal lastRow As Long, _
                                    ByVal dayCol As Long)
    Dim nightCol As Long
    nightCol = dayCol + 1

    With wsWelding.Cells(WUP_RATIO_ROW, dayCol)
        .ClearContents
        .HorizontalAlignment = xlGeneral
    End With
    With wsWelding.Cells(WUP_RATIO_ROW, nightCol)
        .ClearContents
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
    End With

    Dim clearLastRow As Long
    clearLastRow = lastRow
    If clearLastRow < WUP_DATA_START_ROW Then clearLastRow = WUP_DATA_START_ROW + 200

    Dim clearRange As Range
    Set clearRange = wsWelding.Range(wsWelding.Cells(WUP_HEADER_ROW, dayCol), _
                                     wsWelding.Cells(clearLastRow, nightCol))
    SafeUnmergeRangeWUP clearRange
    clearRange.ClearContents
    clearRange.NumberFormat = "General"
    clearRange.Interior.ColorIndex = xlColorIndexNone
    clearRange.ShrinkToFit = False
    clearRange.Borders.LineStyle = xlNone
End Sub

Private Sub ApplyGreyFill(ByVal targetCell As Range)
    With targetCell
        .ClearContents
        .NumberFormat = "General"
        .Interior.Color = RGB(WUP_FILL_COLOR_R, WUP_FILL_COLOR_G, WUP_FILL_COLOR_B)
    End With
End Sub

Private Sub SafeUnmergeRangeWUP(ByVal targetRange As Range)
    On Error Resume Next
    If targetRange.MergeCells Then targetRange.UnMerge
    On Error GoTo 0
End Sub

' =====================================================================
' 基本情報シートの業者ブロック走査
' =====================================================================

' F10から3列おきに工事種別(10行目)を確認し、溶接施工会社ブロックと軌道工事ブロックへ分類する
'
' G/H列(溶接施工会社)の比率参照ブロックの優先順位:
'   1. preferredRatioColumn (31行目が変更された列) のブロック ※軌道工事ブロックでも可
'   2. 10行目=「溶接工事」のブロックの31行目
'   3. 10行目=「軌道工事」で31行目に値を持つ最初のブロック
'
' G5:H5の会社名ソース:
'   採用ブロックが軌道工事 -> 31行目に値を持つ全軌道会社の11行目会社名(後段でA列値へ解決し「・」結合)
'   採用ブロックが溶接工事 -> そのブロックの11行目会社名
Private Sub ScanVendorBlocks(ByVal wsInfo As Worksheet, _
                             ByRef weldingBlock As WeldingVendorBlock, _
                             ByVal weldingNameSources As Collection, _
                             ByRef railBlocks() As WeldingVendorBlock, _
                             ByRef railBlockCount As Long, _
                             ByVal preferredRatioColumn As Long)
    Dim emptyBlock As WeldingVendorBlock
    weldingBlock = emptyBlock
    railBlockCount = 0
    ReDim railBlocks(1 To MAX_VENDOR_BLOCK_COUNT)

    Dim dedicatedBlock As WeldingVendorBlock          ' 10行目=溶接工事のブロック(31行目参照)
    Dim railWeldingBlocks() As WeldingVendorBlock     ' 10行目=軌道工事で31行目に値を持つブロック
    Dim railWeldingCount As Long
    ReDim railWeldingBlocks(1 To MAX_VENDOR_BLOCK_COUNT)

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCountWUP(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = BASIC_INFO_VENDOR_BLOCK_VALUE_COL + ((i - 1) * BASIC_INFO_VENDOR_BLOCK_STEP_COLS)

        Dim workTypeText As String
        workTypeText = NormalizeMatchTextWUP(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueColumn).Value))

        If StrComp(workTypeText, NormalizeMatchTextWUP(WeldingWorkTypeText()), vbTextCompare) = 0 Then
            ' 溶接工事ブロック(最初の1件のみ採用)
            If dedicatedBlock.valueColumn = 0 Then
                dedicatedBlock = BuildVendorBlock(wsInfo, valueColumn, BASIC_INFO_WELDING_RATIO_ROW)
                LogWUP "溶接工事ブロック col=" & CStr(valueColumn) & _
                       " 会社=[" & dedicatedBlock.vendorName & "] 比率有=" & CStr(dedicatedBlock.hasRatio)
            Else
                LogWUP "溶接工事ブロックが複数あります col=" & CStr(valueColumn) & " -> 無視"
            End If
        ElseIf StrComp(workTypeText, NormalizeMatchTextWUP(RailWorkTypeText()), vbTextCompare) = 0 Then
            ' 軌道工事ブロック(基本情報の並び順を維持) -> I/J列以降の手元分単価へ
            railBlockCount = railBlockCount + 1
            railBlocks(railBlockCount) = BuildVendorBlock(wsInfo, valueColumn, BASIC_INFO_RAIL_RATIO_ROW)
            LogWUP "軌道工事ブロック#" & CStr(railBlockCount) & " col=" & CStr(valueColumn) & _
                   " 会社=[" & railBlocks(railBlockCount).vendorName & _
                   "] 比率有=" & CStr(railBlocks(railBlockCount).hasRatio)

            ' 31行目(溶接外注比率)を持つ軌道会社 -> 溶接施工会社の候補
            Dim railWeldingCandidate As WeldingVendorBlock
            railWeldingCandidate = BuildVendorBlock(wsInfo, valueColumn, BASIC_INFO_WELDING_RATIO_ROW)
            If railWeldingCandidate.hasRatio Then
                railWeldingCount = railWeldingCount + 1
                railWeldingBlocks(railWeldingCount) = railWeldingCandidate
                LogWUP "軌道会社の溶接比率あり col=" & CStr(valueColumn) & _
                       " 会社=[" & railWeldingCandidate.vendorName & "]"
            End If
        End If
    Next i

    ' --- G/H列の比率参照ブロックを決定 ---
    If preferredRatioColumn > 0 Then
        Dim preferredBlock As WeldingVendorBlock
        preferredBlock = BuildVendorBlock(wsInfo, preferredRatioColumn, BASIC_INFO_WELDING_RATIO_ROW)
        If preferredBlock.hasRatio Then
            weldingBlock = preferredBlock
            LogWUP "比率参照=変更列 col=" & CStr(preferredRatioColumn)
        End If
    End If
    If weldingBlock.valueColumn = 0 And dedicatedBlock.hasRatio Then
        weldingBlock = dedicatedBlock
        LogWUP "比率参照=溶接工事ブロック col=" & CStr(dedicatedBlock.valueColumn)
    End If
    If weldingBlock.valueColumn = 0 And railWeldingCount > 0 Then
        weldingBlock = railWeldingBlocks(1)
        LogWUP "比率参照=軌道工事ブロック(31行目) col=" & CStr(weldingBlock.valueColumn)
    End If
    If weldingBlock.valueColumn = 0 And dedicatedBlock.valueColumn > 0 Then
        weldingBlock = dedicatedBlock   ' 比率未入力でも警告文言判定用に保持
    End If

    ' --- G5:H5へ表示する会社名ソースを決定 ---
    If weldingBlock.valueColumn > 0 Then
        Dim weldingBlockWorkType As String
        weldingBlockWorkType = NormalizeMatchTextWUP( _
            CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, weldingBlock.valueColumn).Value))

        If StrComp(weldingBlockWorkType, NormalizeMatchTextWUP(RailWorkTypeText()), vbTextCompare) = 0 Then
            ' 軌道会社が溶接施工 -> 31行目に値を持つ全軌道会社名を結合表示
            Dim j As Long
            For j = 1 To railWeldingCount
                AddNameSourceIfMissing weldingNameSources, railWeldingBlocks(j).vendorName
            Next j
            AddNameSourceIfMissing weldingNameSources, weldingBlock.vendorName
        Else
            AddNameSourceIfMissing weldingNameSources, weldingBlock.vendorName
        End If
    End If
End Sub

' 会社名ソースへ重複・空文字を除いて追加
Private Sub AddNameSourceIfMissing(ByVal nameSources As Collection, ByVal vendorName As String)
    Dim normalizedName As String
    normalizedName = CommonNormalizeText(vendorName)
    If Len(normalizedName) = 0 Then Exit Sub

    Dim existingName As Variant
    For Each existingName In nameSources
        If StrComp(CommonNormalizeText(CStr(existingName)), normalizedName, vbTextCompare) = 0 Then Exit Sub
    Next existingName

    nameSources.Add vendorName
End Sub

' 会社名ソースを業者マスタ(B列一致→A列値)へ解決し、複数社は「・」で結合
Private Function BuildWeldingDisplayName(ByVal nameSources As Collection, _
                                         ByVal vendorUnitPriceNameMap As Object) As String
    Dim nameValue As Variant
    Dim result As String
    For Each nameValue In nameSources
        Dim resolvedName As String
        resolvedName = ResolveVendorUnitPriceNameWUP(vendorUnitPriceNameMap, CStr(nameValue))
        If Len(resolvedName) > 0 Then
            If Len(result) > 0 Then result = result & MiddleDotText()
            result = result & resolvedName
        End If
    Next nameValue
    BuildWeldingDisplayName = result
End Function

Private Function GetVendorBlockCountWUP(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CStr(wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).Value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > MAX_VENDOR_BLOCK_COUNT Then countValue = MAX_VENDOR_BLOCK_COUNT
    GetVendorBlockCountWUP = countValue
End Function

Private Function BuildVendorBlock(ByVal wsInfo As Worksheet, _
                                  ByVal valueColumn As Long, _
                                  ByVal ratioRow As Long) As WeldingVendorBlock
    Dim result As WeldingVendorBlock
    result.valueColumn = valueColumn
    result.vendorName = Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).Value))
    result.ratioAddress = "'" & Replace$(wsInfo.Name, "'", "''") & "'!" & _
                          wsInfo.Cells(ratioRow, valueColumn).Address(True, True)
    result.ratioPercent = NormalizeRatioValue(wsInfo.Cells(ratioRow, valueColumn).Value)
    result.hasRatio = IsNumeric(result.ratioPercent)
    BuildVendorBlock = result
End Function

' 比率を0～1へ正規化(67.1 や "67.1%" -> 0.671)。数値化できなければEmpty
Private Function NormalizeRatioValue(ByVal sourceValue As Variant) As Variant
    Dim textValue As String

    If IsError(sourceValue) Then Exit Function
    If IsNumeric(sourceValue) Then
        NormalizeRatioValue = NormalizePercentScale(CDbl(sourceValue))
        Exit Function
    End If

    textValue = Trim$(StrConv(CStr(CommonNzText(sourceValue)), vbNarrow))
    If Len(textValue) = 0 Then Exit Function

    If Right$(textValue, 1) = "%" Then
        textValue = Trim$(Left$(textValue, Len(textValue) - 1))
        If IsNumeric(textValue) Then NormalizeRatioValue = CDbl(textValue) / 100#
        Exit Function
    End If

    If IsNumeric(textValue) Then NormalizeRatioValue = NormalizePercentScale(CDbl(textValue))
End Function

Private Function NormalizePercentScale(ByVal value As Double) As Double
    If value > 1# And value <= 100# Then
        NormalizePercentScale = value / 100#
    Else
        NormalizePercentScale = value
    End If
End Function

' =====================================================================
' 手元比率マスタ(レール溶接_軌道会社外注費率一覧 / 溶接手元割合シート)
' =====================================================================

' Dictionary: key=整理番号(文字列) -> Array(手元比率昼, 手元比率夜) ※未設定はEmpty
Private Function LoadTemotoRatioMap(ByRef loadErrorText As String) As Object
    loadErrorText = ""

    Dim sourceFilePath As String
    sourceFilePath = ResolveTemotoMasterFilePath()
    If sourceFilePath = "" Then
        loadErrorText = "ファイルが見つかりません: マスタデータ\" & TemotoMasterFilePatternText()
        Exit Function
    End If
    LogWUP "手元比率マスタ path=[" & sourceFilePath & "]"

    Dim cn As Object
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then
        loadErrorText = "ADO接続に失敗しました: " & sourceFilePath
        Exit Function
    End If

    On Error GoTo Cleanup

    Dim adoSheetName As String
    adoSheetName = FindAdoSheetNameWUP(cn, TemotoMasterSheetNameText())
    If adoSheetName = "" Then
        loadErrorText = "シート「" & TemotoMasterSheetNameText() & "」が見つかりません: " & sourceFilePath
        GoTo Cleanup
    End If

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & adoSheetName & "$]", cn, 0, 1  ' adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        loadErrorText = "シート「" & TemotoMasterSheetNameText() & "」にデータがありません。"
        CommonCloseAdoRecordset rs
        GoTo Cleanup
    End If

    Dim data As Variant
    data = rs.GetRows
    CommonCloseAdoRecordset rs

    Set LoadTemotoRatioMap = BuildTemotoMapFromData(data, loadErrorText)

Cleanup:
    CommonCloseAdoConnection cn
End Function

' 基本情報C23が「溶接工事あり」の場合のみ、F列以降3列ごとの業者ブロックのうち
' 10行目が「軌道工事」のブロックの31行目へ既定比率(60.1%)を入力する。
' ※31行目が空欄のときだけ入力する(既に値があれば手入力を尊重して上書きしない)。
' ※ScanVendorBlocks後に呼ぶこと(走査時の溶接ブロック判定へ影響させないため)。
' 戻り値: 1件でも入力したらTrue。
Private Function WriteRailOutsourceRatioToBasicInfo(ByVal wsInfo As Worksheet) As Boolean
    Dim flagText As String
    flagText = NormalizeMatchTextWUP(CStr(wsInfo.Cells(BASIC_INFO_WELDING_FLAG_ROW, _
                                                       BASIC_INFO_WELDING_FLAG_COL).Value))
    If InStr(1, flagText, NormalizeMatchTextWUP(WeldingWorkPresentKeywordText()), vbTextCompare) = 0 Then
        Exit Function   ' C23が「溶接工事あり」でない -> 何もしない
    End If

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCountWUP(wsInfo)

    Dim i As Long, valueColumn As Long, workTypeText As String
    For i = 1 To vendorCount
        valueColumn = BASIC_INFO_VENDOR_BLOCK_VALUE_COL + ((i - 1) * BASIC_INFO_VENDOR_BLOCK_STEP_COLS)
        workTypeText = NormalizeMatchTextWUP(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueColumn).Value))
        If StrComp(workTypeText, NormalizeMatchTextWUP(RailWorkTypeText()), vbTextCompare) = 0 Then
            Dim ratioCell As Range
            Set ratioCell = wsInfo.Cells(BASIC_INFO_WELDING_RATIO_ROW, valueColumn)
            ' 空欄のときだけ既定値を入れる。手入力済み(空欄でない)なら上書きしない。
            If Len(Trim$(CStr(ratioCell.Value))) = 0 Then
                ratioCell.Value = WUP_RAIL_FIXED_RATIO
                ratioCell.NumberFormat = WUP_RATIO_NUMBER_FORMAT
                WriteRailOutsourceRatioToBasicInfo = True
                LogWUP "軌道工事ブロック col=" & CStr(valueColumn) & " 31行目へ既定値" & _
                       Format$(WUP_RAIL_FIXED_RATIO, "0.0%") & "を入力(空欄のため)"
            Else
                LogWUP "軌道工事ブロック col=" & CStr(valueColumn) & " 31行目は既存値[" & _
                       CStr(ratioCell.Value) & "]のため上書きせず"
            End If
        End If
    Next i
End Function

Private Function BuildTemotoMapFromData(ByVal data As Variant, _
                                        ByRef loadErrorText As String) As Object
    Dim fieldCount As Long
    Dim recordCount As Long
    fieldCount = UBound(data, 1) + 1
    recordCount = UBound(data, 2) + 1

    ' --- ヘッダー行の検出(「整理番号」を含むセルを探す) ---
    Dim headerRecord As Long
    Dim seiriField As Long
    headerRecord = -1
    seiriField = -1

    Dim r As Long, f As Long
    For r = 0 To Application.WorksheetFunction.Min(recordCount - 1, 19)
        For f = 0 To fieldCount - 1
            If InStr(1, NormalizeMatchTextWUP(CommonNzText(data(f, r))), _
                     SeiriHeaderKeywordText(), vbTextCompare) > 0 Then
                headerRecord = r
                seiriField = f
                Exit For
            End If
        Next f
        If headerRecord >= 0 Then Exit For
    Next r

    If headerRecord < 0 Then
        loadErrorText = "ヘッダー行(整理番号)を検出できませんでした。"
        Exit Function
    End If

    ' --- 昼/夜の手元比率列の検出 ---
    Dim dayField As Long
    Dim nightField As Long
    dayField = -1
    nightField = -1
    For f = 0 To fieldCount - 1
        Dim headerCellText As String
        headerCellText = NormalizeMatchTextWUP(CommonNzText(data(f, headerRecord)))
        If f <> seiriField And Len(headerCellText) > 0 Then
            If dayField < 0 And InStr(1, headerCellText, DayKeywordText(), vbTextCompare) > 0 Then
                dayField = f
            ElseIf nightField < 0 And InStr(1, headerCellText, NightKeywordText(), vbTextCompare) > 0 Then
                nightField = f
            End If
        End If
    Next f

    ' フォールバック: 整理番号列の右隣を昼、その右を夜とみなす
    If dayField < 0 Then dayField = seiriField + 1
    If nightField < 0 Then nightField = dayField + 1
    If dayField > fieldCount - 1 Then
        loadErrorText = "手元比率(昼)の列を検出できませんでした。"
        Exit Function
    End If
    LogWUP "手元比率マスタ headerRow=" & CStr(headerRecord + 1) & _
           " 整理番号列=" & CStr(seiriField + 1) & _
           " 昼列=" & CStr(dayField + 1) & " 夜列=" & CStr(nightField + 1)

    ' --- 整理番号 -> Array(昼, 夜) の辞書を構築 / 5000番台はパック構成も登録 ---
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Set mPackMap = CreateObject("Scripting.Dictionary")
    mPackMap.CompareMode = vbTextCompare

    For r = headerRecord + 1 To recordCount - 1
        Dim seiriText As String
        seiriText = Trim$(StrConv(CommonNzText(data(seiriField, r)), vbNarrow))
        If Len(seiriText) > 0 And IsNumeric(seiriText) Then
            Dim seiriKey As String
            seiriKey = CStr(CLng(Val(seiriText)))

            Dim dayRatio As Variant
            Dim nightRatio As Variant
            dayRatio = NormalizeRatioValue(data(dayField, r))
            If nightField <= fieldCount - 1 Then
                nightRatio = NormalizeRatioValue(data(nightField, r))
            End If

            If Not result.Exists(seiriKey) Then
                result.Add seiriKey, Array(dayRatio, nightRatio)
            End If

            ' パック工種(5000番台): J/K, L/M, N/O から構成(整理番号, 数量)を読む
            If CLng(Val(seiriText)) >= WUP_PACK_SEIRI_MIN And Not mPackMap.Exists(seiriKey) Then
                Dim comps As Collection
                Set comps = New Collection
                AddPackComponent comps, data, fieldCount, r, _
                                 WUP_MASTER_PACK_COMP1_FIELD, WUP_MASTER_PACK_QTY1_FIELD
                AddPackComponent comps, data, fieldCount, r, _
                                 WUP_MASTER_PACK_COMP2_FIELD, WUP_MASTER_PACK_QTY2_FIELD
                AddPackComponent comps, data, fieldCount, r, _
                                 WUP_MASTER_PACK_COMP3_FIELD, WUP_MASTER_PACK_QTY3_FIELD
                If comps.Count > 0 Then mPackMap.Add seiriKey, comps
            End If
        End If
    Next r

    LogWUP "パック構成マスタ 件数=" & CStr(mPackMap.Count)

    If result.Count = 0 Then
        loadErrorText = "手元比率データを1件も読み込めませんでした。"
        Exit Function
    End If

    Set BuildTemotoMapFromData = result
End Function

' パック構成1件を追加(構成整理番号・数量がともに数値のときのみ)
Private Sub AddPackComponent(ByVal comps As Collection, ByVal data As Variant, _
                             ByVal fieldCount As Long, ByVal r As Long, _
                             ByVal compField As Long, ByVal qtyField As Long)
    If compField > fieldCount - 1 Or qtyField > fieldCount - 1 Then Exit Sub
    Dim compText As String, qtyText As String
    compText = Trim$(StrConv(CommonNzText(data(compField, r)), vbNarrow))
    qtyText = Trim$(StrConv(CommonNzText(data(qtyField, r)), vbNarrow))
    If Len(compText) = 0 Or Not IsNumeric(compText) Then Exit Sub
    If Len(qtyText) = 0 Or Not IsNumeric(qtyText) Then Exit Sub
    comps.Add Array(CStr(CLng(Val(compText))), CDbl(qtyText))
End Sub

' マスタファイルのパス解決(業者マスタと同方式)
'   1) %USERPROFILE%\大鉄工業株式会社\線路出張所用_注文書_請求書アクセスサイト - ドキュメント\マスタデータ\
'   2) ThisWorkbook の親フォルダ\マスタデータ\
'   3) ThisWorkbook と同階層\マスタデータ\
' ファイル名は「レール溶接_軌道会社外注費率一覧*.xlsx」でワイルドカード検索(末尾の空白等を許容)
Private Function ResolveTemotoMasterFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderCandidates As Collection
    Set folderCandidates = New Collection

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(userProfilePath)) > 0 Then
        folderCandidates.Add userProfilePath & "\" & CommonCompanyNameText() & "\" & _
                             OrderInvoiceDocumentFolderTextWUP() & "\" & MasterDataFolderText()
    End If
    If Len(ThisWorkbook.Path) > 0 Then
        folderCandidates.Add fso.GetParentFolderName(ThisWorkbook.Path) & "\" & MasterDataFolderText()
        folderCandidates.Add ThisWorkbook.Path & "\" & MasterDataFolderText()
    End If

    Dim folderPath As Variant
    For Each folderPath In folderCandidates
        Dim foundName As String
        foundName = ""
        On Error Resume Next
        foundName = Dir(CStr(folderPath) & "\" & TemotoMasterFilePatternText(), vbNormal)
        On Error GoTo 0
        If Len(foundName) > 0 Then
            ResolveTemotoMasterFilePath = CStr(folderPath) & "\" & foundName
            Exit Function
        End If
    Next folderPath
End Function

Private Function FindAdoSheetNameWUP(ByVal cn As Object, ByVal targetSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)

    Dim normalizedTarget As String
    normalizedTarget = NormalizeMatchTextWUP(targetSheetName)

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(NormalizeMatchTextWUP(CStr(sheetName)), normalizedTarget, vbTextCompare) = 0 Then
            FindAdoSheetNameWUP = CStr(sheetName)
            Exit Function
        End If
    Next sheetName

    ' 完全一致がなければ部分一致で再検索
    For Each sheetName In sheetNames
        If InStr(1, NormalizeMatchTextWUP(CStr(sheetName)), normalizedTarget, vbTextCompare) > 0 Then
            FindAdoSheetNameWUP = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

' =====================================================================
' ヘルパー
' =====================================================================

Private Function CollectWeldingUnitPriceSheets(ByVal targetBook As Workbook) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In targetBook.Worksheets
        If IsWeldingUnitPriceSheet(ws) Then result.Add ws
    Next ws

    Set CollectWeldingUnitPriceSheets = result
End Function

' (回数)(年度4桁)外注単価 / 軌道会社列は(回数)(年度4桁)溶接手元単価
Private Function BuildWeldingUnitPriceHeaderText(ByVal wsInfo As Worksheet, _
                                                 Optional ByVal isWeldingVendor As Boolean = True) As String
    Dim labelText As String
    If isWeldingVendor Then
        labelText = OutsourceUnitPriceLabelText()
    Else
        labelText = WeldingTemotoUnitPriceLabelText()
    End If

    BuildWeldingUnitPriceHeaderText = Trim$(CStr(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).Value)) & _
                                      CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).Value)) & _
                                      labelText
End Function

Private Function ResolveVendorUnitPriceNameWUP(ByVal vendorUnitPriceNameMap As Object, _
                                               ByVal basicInfoVendorName As String) As String
    Dim vendorNameKey As String
    vendorNameKey = CommonNormalizeText(basicInfoVendorName)

    If Not vendorUnitPriceNameMap Is Nothing Then
        If vendorUnitPriceNameMap.Exists(vendorNameKey) Then
            ResolveVendorUnitPriceNameWUP = CStr(vendorUnitPriceNameMap(vendorNameKey))
            Exit Function
        End If
    End If

    ResolveVendorUnitPriceNameWUP = basicInfoVendorName
End Function

Private Function WarnWeldingBlockText(ByRef weldingBlock As WeldingVendorBlock) As String
    If weldingBlock.valueColumn = 0 Then
        WarnWeldingBlockText = "基本情報に「溶接工事」のブロックも、31行目に溶接外注比率を持つ「軌道工事」のブロックも見つかりません。" & vbCrLf & _
                               "溶接施工会社の単価列(G/H列)は作成されません。"
    Else
        WarnWeldingBlockText = "基本情報の溶接外注比率(31行目)が未入力です。" & vbCrLf & _
                               "溶接施工会社の単価列(G/H列)は作成されません。"
    End If
End Function

Private Function BuildMissingSeiriWarningText(ByVal missingSeiriMap As Object) As String
    Dim keys As Variant
    keys = missingSeiriMap.keys

    Dim displayText As String
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then displayText = displayText & ", "
        If i - LBound(keys) >= 10 Then
            displayText = displayText & "...(計" & CStr(missingSeiriMap.Count) & "件)"
            Exit For
        End If
        displayText = displayText & CStr(keys(i))
    Next i

    BuildMissingSeiriWarningText = "手元比率マスタに以下の整理番号が見つからないため、" & _
                                   "該当行はグレー塗りにしました:" & vbCrLf & displayText
End Function

Private Function JoinCollectionText(ByVal values As Collection, ByVal delimiter As String) As String
    Dim itemValue As Variant
    Dim result As String
    For Each itemValue In values
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(itemValue)
    Next itemValue
    JoinCollectionText = result
End Function

Private Function NormalizeMatchTextWUP(ByVal value As String) As String
    NormalizeMatchTextWUP = CommonRemoveAllSpaces(CommonNormalizeText(value))
End Function

Private Sub LogWUP(ByVal msg As String)
    On Error Resume Next
    mod_DebugLog.Log "[WeldingUP] " & msg
    On Error GoTo 0
End Sub

' =====================================================================
' 文字列定数(CP932互換のためChrW$で定義 / 既存モジュールの方式に準拠)
' =====================================================================

' "_レール溶接単価"
Private Function WeldingSheetSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = "_" & ChrW$(&H30EC) & ChrW$(&H30FC) & ChrW$(&H30EB) & _
                 ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H5358) & ChrW$(&H4FA1)
    End If
    WeldingSheetSuffixText = cached
End Function

' "溶接工事"
Private Function WeldingWorkTypeText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
    End If
    WeldingWorkTypeText = cached
End Function

' 「溶接工事あり」(基本情報C23の判定キーワード)
Private Function WeldingWorkPresentKeywordText() As String
    Static cached As String
    If cached = "" Then
        cached = WeldingWorkTypeText() & ChrW$(&H3042) & ChrW$(&H308A)   ' 溶接工事 + あり
    End If
    WeldingWorkPresentKeywordText = cached
End Function

' "軌道工事"
Private Function RailWorkTypeText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
    End If
    RailWorkTypeText = cached
End Function

' "溶接手元割合"
Private Function TemotoMasterSheetNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & _
                 ChrW$(&H5143) & ChrW$(&H5272) & ChrW$(&H5408)
    End If
    TemotoMasterSheetNameText = cached
End Function

' "レール溶接_軌道会社外注費率一覧*.xlsx"
Private Function TemotoMasterFilePatternText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30EC) & ChrW$(&H30FC) & ChrW$(&H30EB) & _
                 ChrW$(&H6EB6) & ChrW$(&H63A5) & "_" & _
                 ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                 ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H8CBB) & ChrW$(&H7387) & _
                 ChrW$(&H4E00) & ChrW$(&H89A7) & "*.xlsx"
    End If
    TemotoMasterFilePatternText = cached
End Function

' "マスタデータ"
Private Function MasterDataFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & _
                 ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF)
    End If
    MasterDataFolderText = cached
End Function

' "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
Private Function OrderInvoiceDocumentFolderTextWUP() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H7528) & _
                 "_" & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & "_" & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & _
                 ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & ChrW$(&H30B9) & _
                 ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & _
                 " - " & ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    OrderInvoiceDocumentFolderTextWUP = cached
End Function

' "外注単価"
Private Function OutsourceUnitPriceLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H5358) & ChrW$(&H4FA1)
    End If
    OutsourceUnitPriceLabelText = cached
End Function

' "溶接手元単価"
Private Function WeldingTemotoUnitPriceLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & ChrW$(&H5143) & _
                 ChrW$(&H5358) & ChrW$(&H4FA1)
    End If
    WeldingTemotoUnitPriceLabelText = cached
End Function

' "外注比率＝"
Private Function OutsourceRatioLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&HFF1D)
    End If
    OutsourceRatioLabelText = cached
End Function

' "・" (複数会社名の結合用)
Private Function MiddleDotText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H30FB)
    MiddleDotText = cached
End Function

' "昼間"
Private Function DayLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H663C) & ChrW$(&H9593)
    DayLabelText = cached
End Function

' "夜間"
Private Function NightLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H591C) & ChrW$(&H9593)
    NightLabelText = cached
End Function

' "産廃処理"
Private Function WasteDisposalKeywordText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7523) & ChrW$(&H5EC3) & ChrW$(&H51E6) & ChrW$(&H7406)
    End If
    WasteDisposalKeywordText = cached
End Function

' "整理番号"
Private Function SeiriHeaderKeywordText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H6574) & ChrW$(&H7406) & ChrW$(&H756A) & ChrW$(&H53F7)
    End If
    SeiriHeaderKeywordText = cached
End Function

' "昼"
Private Function DayKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H663C)
    DayKeywordText = cached
End Function

' "夜"
Private Function NightKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H591C)
    NightKeywordText = cached
End Function

' "BIZ UDゴシック"
Private Function WeldingUnitPriceFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    WeldingUnitPriceFontNameText = cached
End Function
