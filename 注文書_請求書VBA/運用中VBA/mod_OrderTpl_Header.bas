Option Explicit

' 注文書テンプレート各シートへの基本情報ヘッダー転記。
' 生成時に加え、基本情報シートの転記元セル変更時にも再転記される(ライブ反映)。
' 転記対象シートの追加は ApplyVendorSheetHeaders のディスパッチへ実装を差し込む。
' 改修履歴: CHANGELOG.md 参照

' 基本情報のヘッダー転記元セル(全社共通分)。ブロック列(16/27行目)は動的に組み立てる
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' 受注者用シート転記で参照する施工会社ブロック行(基本情報)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35
' 施工会社ブロック(基本情報)から参照する項目: 代表者名(12行目)・住所(14行目)
Private Const CONTRACTOR_REPRESENTATIVE_ROW As Long = 12
Private Const CONTRACTOR_ADDRESS_ROW As Long = 14
Private Const CONDITION_CHECKBOX_D_COL As Long = 4
Private Const CONDITION_CHECKBOX_X_COL As Long = 24
Private Const CONDITION_CHECKBOX_E_COL As Long = 5
Private Const CONDITION_CHECKBOX_LEFT_MAX_COL As Long = 12
Private Const CONDITION_CHECKBOX_RIGHT_MIN_COL As Long = 20
Private Const CONDITION_ROW38_BAND_MIN_ROW As Long = 37
Private Const CONDITION_ROW38_BAND_MAX_ROW As Long = 39
Private Const CONDITION_E_PAIR_MIN_ROW As Long = 34
Private Const CONDITION_E_PAIR_MAX_ROW As Long = 35
' チェックボックスはTopLeftCell.Rowでテンプレート内シートのアンカー配置に依存する。
' Excelの仕様上、35～40 または 36～41 のいずれかになりうる。
' どちらでも動作するように範囲を35～41に拡張して吸収する。
Private Const ATTACHMENT3_CHECK_ROW_MIN As Long = 35
Private Const ATTACHMENT3_CHECK_ROW_MAX As Long = 41
Private Const ATTACHMENT3_COL_F As Long = 6
Private Const ATTACHMENT3_COL_H As Long = 8
Private Const ATTACHMENT3_COL_J As Long = 10
Private Const ATTACHMENT3_COL_L As Long = 12
Public Const ATTACHMENT3_SANPAI_ROW_MIN As Long = 49
Public Const ATTACHMENT3_SANPAI_ROW_MAX As Long = 52
' 産廃施設選択セルは名称側(4列目起点)と値側(10列目起点)の2つの結合セルで構成される。
' ダブルクリック判定は名称側セルの列(4列目)を基準に行う。
Public Const ATTACHMENT3_SANPAI_NAME_COL As Long = 4
Public Const ATTACHMENT3_SANPAI_VALUE_COL As Long = 10
Private Const SANPAI_FACILITY_MASTER_START_ROW As Long = 59

' 指定ブロックの施工会社に対応するテンプレート5シートへヘッダーを転記する(ディスパッチャ)
Public Sub ApplyVendorSheetHeaders(ByVal wsInfo As Worksheet, _
                                   ByVal vendorIndex As Long, _
                                   ByVal aliasText As String)
    If wsInfo Is Nothing Then Exit Sub
    If Len(aliasText) = 0 Then Exit Sub

    Dim baseNames As Variant
    baseNames = mod_OrderTpl_Shared.OrderTplTemplateSheetBaseNames()

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        Dim sheetName As String
        sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(i)), aliasText)
        If mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then
            Dim wsTarget As Worksheet
            Set wsTarget = ThisWorkbook.Worksheets(sheetName)

            ApplyVendorSheetTabColor wsInfo, wsTarget, vendorIndex

            Select Case i - LBound(baseNames)
                Case 0: ApplyBreakdownHeader wsInfo, wsTarget, vendorIndex
                Case 1: ApplyContractorHeader wsInfo, wsTarget, vendorIndex
                Case 2: ApplyAcceptanceHeader wsInfo, wsTarget, vendorIndex
                Case 3: ApplyBranchCopyHeader wsInfo, wsTarget, vendorIndex
                Case 4: ApplyAttachment3Header wsInfo, wsTarget, vendorIndex
            End Select
        End If
    Next i

    Dim condSheetName As String
    condSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
        mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetName) Then
        ApplyVendorSheetTabColor wsInfo, ThisWorkbook.Worksheets(condSheetName), vendorIndex
        SetupConditionCheckboxExclusivity ThisWorkbook.Worksheets(condSheetName)
    End If
End Sub

' 全確定会社のテンプレートシートへヘッダーを再転記する
Public Sub RefreshAllVendorSheetHeaders(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim indexes As Collection
    Set indexes = New Collection
    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)
    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        indexes.Add vendorIndex
    Next vendorIndex
    RefreshVendorSheetHeadersForIndexes wsInfo, indexes
End Sub

Public Sub RefreshVendorSheetHeadersForIndexes(ByVal wsInfo As Worksheet, ByVal vendorIndexes As Collection)
    If wsInfo Is Nothing Then Exit Sub
    If vendorIndexes Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim t0 As Double
    t0 = mod_Construction_Import_Shared.LogCIStart()

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim idx As Variant
    For Each idx In vendorIndexes
        Dim vendorIndex As Long
        vendorIndex = CLng(idx)
        Dim companyName As String
        companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
        If companyName <> "" Then
            Dim vendorName As String
            Dim aliasText As String
            Dim workText As String
            If mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then
                ApplyVendorSheetHeaders wsInfo, vendorIndex, aliasText

                Dim condSheetName As String
                condSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
                If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetName) Then
                    ' チェックボックス排他は ApplyVendorSheetHeaders 内で設定済みのため再実行しない
                    ApplyConditionSheetHeader wsInfo, ThisWorkbook.Worksheets(condSheetName), vendorIndex
                End If
            End If
        End If
    Next idx

    mod_Construction_Import_Shared.LogCIElapsed "RefreshVendorSheetHeadersForIndexes count=" & vendorIndexes.Count, t0
    Exit Sub

Quiet:
    mod_OrderTpl_Shared.OrderTplLog "RefreshVendorSheetHeadersForIndexes error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Sheet1(基本情報)のWorksheet_Changeから呼ばれる入口。
' ヘッダー転記元セル(B6/C6/C2/C9/C10/C15:C16/F6、各ブロックの16/27行目)の変更を各社シートへ反映する
Public Sub HandleBasicInfoHeaderSourceChange(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim sourceRange As Range
    Set sourceRange = BuildHeaderSourceRange(wsInfo)
    If sourceRange Is Nothing Then Exit Sub
    If Intersect(target, sourceRange) Is Nothing Then Exit Sub

    Dim headerT0 As Double
    headerT0 = mod_Construction_Import_Shared.LogCIStart()

    ' 施工会社名(11行目)のみの変更は RunScheduled 側で ApplyVendorSheetHeaders される。
    ' ここで全社分を再転記すると数十秒掛かり、シート生成までブロックする。
    If IsVendorNameRowOnlyHeaderChange(wsInfo, target) Then
        mod_Construction_Import_Shared.LogCI "HandleBasicInfoHeaderSourceChange skip(all) name-row-only"
        Exit Sub
    End If

    Dim affectedIndexes As Collection
    Set affectedIndexes = CollectAffectedVendorIndexesFromHeaderChange(wsInfo, target)
    If affectedIndexes Is Nothing Then
        RefreshAllVendorSheetHeaders wsInfo
        mod_Construction_Import_Shared.LogCIElapsed "HandleBasicInfoHeaderSourceChange all", headerT0
    ElseIf affectedIndexes.Count = 0 Then
        Exit Sub
    Else
        RefreshVendorSheetHeadersForIndexes wsInfo, affectedIndexes
        mod_Construction_Import_Shared.LogCIElapsed "HandleBasicInfoHeaderSourceChange partial count=" & affectedIndexes.Count, headerT0
    End If
    Exit Sub

Quiet:
    Err.Clear
End Sub

' ヘッダー転記元セルの監視範囲を返す公開ラッパー(Sheet1の変更ゲート構築用)
Private Function IsVendorNameRowOnlyHeaderChange(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function
    If wsInfo Is Nothing Then Exit Function

    Dim hit As Range
    Set hit = Intersect(target, BuildHeaderSourceRange(wsInfo))
    If hit Is Nothing Then Exit Function

    Dim cell As Range
    For Each cell In hit.Cells
        If cell.Row <> BASIC_INFO_VENDOR_NAME_ROW Then Exit Function
        If mod_VendorUnitPrice.GetVendorIndexFromValueColumn(cell.Column) < 1 Then Exit Function
    Next cell
    IsVendorNameRowOnlyHeaderChange = True
End Function

' 変更セルに対応する施工会社ブロック番号だけを返す。共通セル(B6等)を含む場合は Nothing(=全件)。
Private Function CollectAffectedVendorIndexesFromHeaderChange(ByVal wsInfo As Worksheet, _
                                                              ByVal target As Range) As Collection
    If target Is Nothing Then Exit Function
    If wsInfo Is Nothing Then Exit Function

    Dim commonHit As Range
    Set commonHit = Intersect(target, wsInfo.Range(HEADER_SOURCE_COMMON_CELLS))
    If Not commonHit Is Nothing Then
        Set CollectAffectedVendorIndexesFromHeaderChange = Nothing
        Exit Function
    End If

    Dim headerHit As Range
    Set headerHit = Intersect(target, BuildHeaderSourceRange(wsInfo))
    If headerHit Is Nothing Then
        Set CollectAffectedVendorIndexesFromHeaderChange = New Collection
        Exit Function
    End If

    Dim result As Collection
    Set result = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In headerHit.Cells
        Dim vendorIndex As Long
        vendorIndex = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(cell.Column)
        If vendorIndex >= 1 Then
            If Not seen.Exists(vendorIndex) Then
                seen.Add vendorIndex, True
                result.Add vendorIndex
            End If
        End If
    Next cell

    Set CollectAffectedVendorIndexesFromHeaderChange = result
End Function

Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' ヘッダー転記元セルの監視範囲(共通セル + 各ブロックの業者コード16行目/注文番号27行目)
Private Function BuildHeaderSourceRange(ByVal wsInfo As Worksheet) As Range
    Dim result As Range
    Set result = wsInfo.Range(HEADER_SOURCE_COMMON_CELLS)

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
        Set result = Union(result, _
                           wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueColumn), _
                           wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn), _
                           wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn), _
                           wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn), _
                           wsInfo.Cells(38, valueColumn), _
                           wsInfo.Cells(39, valueColumn), _
                           wsInfo.Cells(40, valueColumn), _
                           wsInfo.Cells(41, valueColumn), _
                           wsInfo.Cells(42, valueColumn))
    Next vendorIndex

    Set BuildHeaderSourceRange = result
End Function

' シート見出し(タブ)色: 工事区分(基本情報10行目)セルの塗りつぶし色を適用する
Private Sub ApplyVendorSheetTabColor(ByVal wsInfo As Worksheet, _
                                     ByVal wsTarget As Worksheet, _
                                     ByVal vendorIndex As Long)
    On Error Resume Next
    Dim sourceCell As Range
    Set sourceCell = wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, _
                                  mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex))
    If sourceCell.Interior.ColorIndex = xlColorIndexNone Then
        wsTarget.Tab.ColorIndex = xlColorIndexNone
    Else
        wsTarget.Tab.Color = sourceCell.Interior.Color
    End If
    On Error GoTo 0
End Sub


' 内訳明細ヘッダー部へ基本情報シートの内容を転記する
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: 部店コード(出張所別_単価適用線区の単価適用線区シート B/C列照合 → G列)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: 注文番号(施工会社ブロック27行目)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: 工事番号(基本情報C9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' D6: 工事名称(基本情報C10、結合セルのため中央揃え)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: 工期 自/至(基本情報C15/C16、和暦表示)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: 作成日(基本情報C2、和暦表示)
    ' P2:Q2 created date from C2. O2 keeps template label.
    WriteHeaderDate wsBreakdown.Range("P2"), wsInfo.Range("C2").value
    wsBreakdown.Range("P2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE

    ' O3: 出張所長(基本情報F6)
    ' P3:Q3 branch chief from F6. O3 keeps template label.
    WriteHeaderValue wsBreakdown.Range("P3"), wsInfo.Range("F6").value, True

    ' P5: 外注会社名(施工会社ブロック11行目)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: 業者コード(施工会社ブロック16行目)
    WriteHeaderValue wsBreakdown.Range("P6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' 複数ページ印刷用にヘッダー行(7:10行目)をタイトル行に設定する
    On Error Resume Next
    wsBreakdown.PageSetup.PrintTitleRows = ORDER_TPL_PRINT_TITLE_ROWS
    On Error GoTo ErrorHandler

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader done: " & wsBreakdown.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Rail-work type text for H25:U25 copy (rail condition sheet only)
Private Function ConditionRailWorkTypeText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H8ECC, &H9053, &H5DE5, &H4E8B)
    End If
    ConditionRailWorkTypeText = cached
End Function

' Copy basic info / vendor fields into condition sheet.
' Ten shared fields + rail sheet H25:U25 from vendor row41.
Public Sub ApplyConditionSheetHeader(ByVal wsInfo As Worksheet, _
                                     ByVal wsCondition As Worksheet, _
                                     ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsCondition Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' S1:X1 created date from C2 (Gregorian yyyy/m/d, centered)
    WriteHeaderDateGregorian wsCondition.Range("S1"), wsInfo.Range("C2").value
    wsCondition.Range("S1").MergeArea.Cells(1, 1).HorizontalAlignment = xlCenter

    ' P3:S3 branch B6 / T3:X3 office C6 / T4:X4 chief F6 (bottom aligned)
    WriteHeaderValueBottom wsCondition.Range("P3"), wsInfo.Range("B6").value
    WriteHeaderValueBottom wsCondition.Range("T3"), wsInfo.Range("C6").value
    WriteHeaderValueBottom wsCondition.Range("T4"), wsInfo.Range("F6").value

    ' B5:C6 vendor name row11 (bottom aligned)
    WriteHeaderValueBottom wsCondition.Range("B5"), _
        wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' F9:I9 project no C9 / L9:X9 project name C10 (centered)
    WriteHeaderValue wsCondition.Range("F9"), wsInfo.Range("C9").value, True
    WriteHeaderValue wsCondition.Range("L9"), wsInfo.Range("C10").value, True

    ' F11:L11 start C15 / R11:X11 end C16 (Gregorian yyyy/m/d)
    WriteHeaderDateGregorian wsCondition.Range("F11"), wsInfo.Range("C15").value
    WriteHeaderDateGregorian wsCondition.Range("R11"), wsInfo.Range("C16").value

    ' D16:X16 site C13 (centered)
    WriteHeaderValue wsCondition.Range("D16"), wsInfo.Range("C13").value, True

    ' Rail-work condition sheet only: H25:U25 <- vendor row41 (left aligned)
    Dim workType As String
    workType = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueColumn).value))
    If StrComp(workType, ConditionRailWorkTypeText(), vbTextCompare) = 0 Then
        WriteHeaderValueLeft wsCondition.Range("H25"), wsInfo.Cells(41, valueColumn).value
    End If

    mod_OrderTpl_Shared.OrderTplLog "ApplyConditionSheetHeader done: " & wsCondition.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyConditionSheetHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' ===== Condition sheet checkbox exclusivity =====
' D/X rows (10,18-33,38) and E34/E35 pairs keep exactly one side ON.
' Turning one ON turns the other OFF; turning one OFF turns the other ON.


' Exclusive target rows (D/X: 10,18-33,38 / E vertical pair: 34,35)
Private Function IsConditionDxExclusiveRow(ByVal rowIndex As Long) As Boolean
    IsConditionDxExclusiveRow = (rowIndex = 10) Or (rowIndex >= 18 And rowIndex <= 33) Or (rowIndex = 38)
End Function

Private Function IsConditionEVerticalPairRow(ByVal rowIndex As Long) As Boolean
    IsConditionEVerticalPairRow = (rowIndex = 34) Or (rowIndex = 35)
End Function

Private Function IsConditionDxLeftSideColumn(ByVal colIndex As Long) As Boolean
    IsConditionDxLeftSideColumn = (colIndex > 0 And colIndex <= CONDITION_CHECKBOX_LEFT_MAX_COL And _
                                   colIndex <> CONDITION_CHECKBOX_E_COL)
End Function

Private Function IsConditionDxRightSideColumn(ByVal colIndex As Long) As Boolean
    IsConditionDxRightSideColumn = (colIndex >= CONDITION_CHECKBOX_RIGHT_MIN_COL)
End Function

Private Function IsConditionDxSideCheckbox(ByVal cb As Object) As Boolean
    Dim colIndex As Long
    colIndex = cb.TopLeftCell.Column
    IsConditionDxSideCheckbox = IsConditionDxLeftSideColumn(colIndex) Or _
                                IsConditionDxRightSideColumn(colIndex)
End Function

Private Function ConditionCheckboxOverlapsCell(ByVal cb As Object, _
                                               ByVal ws As Worksheet, _
                                               ByVal rowIndex As Long, _
                                               ByVal colIndex As Long) As Boolean
    On Error GoTo Fail
    Dim cell As Range
    Set cell = ws.Cells(rowIndex, colIndex)
    Dim centerX As Double
    Dim centerY As Double
    centerX = cb.Left + (cb.Width / 2#)
    centerY = cb.Top + (cb.Height / 2#)
    ConditionCheckboxOverlapsCell = (centerX >= cell.Left And centerX < (cell.Left + cell.Width) And _
                                     centerY >= cell.Top And centerY < (cell.Top + cell.Height))
    Exit Function
Fail:
    ConditionCheckboxOverlapsCell = False
End Function

Private Function ConditionEVerticalCheckboxRow(ByVal cb As Object, ByVal ws As Worksheet) As Long
    If ConditionCheckboxOverlapsCell(cb, ws, CONDITION_E_PAIR_MAX_ROW, CONDITION_CHECKBOX_E_COL) Then
        ConditionEVerticalCheckboxRow = CONDITION_E_PAIR_MAX_ROW
        Exit Function
    End If
    If ConditionCheckboxOverlapsCell(cb, ws, CONDITION_E_PAIR_MIN_ROW, CONDITION_CHECKBOX_E_COL) Then
        ConditionEVerticalCheckboxRow = CONDITION_E_PAIR_MIN_ROW
        Exit Function
    End If
    If cb.TopLeftCell.Column = CONDITION_CHECKBOX_E_COL Then
        If cb.TopLeftCell.Row >= CONDITION_E_PAIR_MIN_ROW And _
           cb.TopLeftCell.Row <= CONDITION_E_PAIR_MAX_ROW Then
            ConditionEVerticalCheckboxRow = cb.TopLeftCell.Row
        End If
    End If
End Function

Private Function IsConditionEVerticalCheckbox(ByVal cb As Object, ByVal ws As Worksheet) As Boolean
    IsConditionEVerticalCheckbox = (ConditionEVerticalCheckboxRow(cb, ws) > 0)
End Function

Private Function IsConditionRow38BandRow(ByVal rowIndex As Long) As Boolean
    IsConditionRow38BandRow = (rowIndex >= CONDITION_ROW38_BAND_MIN_ROW And _
                               rowIndex <= CONDITION_ROW38_BAND_MAX_ROW)
End Function

Private Function IsConditionExclusiveCheckbox(ByVal cb As Object, ByVal ws As Worksheet) As Boolean
    Dim rowIndex As Long
    rowIndex = cb.TopLeftCell.Row

    If IsConditionEVerticalCheckbox(cb, ws) Then
        IsConditionExclusiveCheckbox = True
        Exit Function
    End If

    If Not IsConditionDxSideCheckbox(cb) Then Exit Function

    If IsConditionDxExclusiveRow(rowIndex) Then
        IsConditionExclusiveCheckbox = True
    ElseIf IsConditionRow38BandRow(rowIndex) Then
        ' Row-38 controls may anchor to rows 37-39.
        IsConditionExclusiveCheckbox = True
    End If
End Function

' Exclusive target rows (D/X: 10,18-33,38 / E vertical pair: 34,35)
Private Function IsConditionExclusiveRow(ByVal r As Long) As Boolean
    IsConditionExclusiveRow = IsConditionDxExclusiveRow(r) Or IsConditionEVerticalPairRow(r)
End Function

' Assign exclusive-click macro to condition-sheet checkboxes
Public Sub SetupConditionCheckboxExclusivity(ByVal wsCondition As Worksheet)
    If wsCondition Is Nothing Then Exit Sub
    FixConditionE34E35Pair wsCondition
    On Error Resume Next
    Dim cb As Object
    For Each cb In wsCondition.CheckBoxes
        If IsConditionExclusiveCheckbox(cb, wsCondition) Then
            cb.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
        End If
    Next cb
    On Error GoTo 0
    NormalizeConditionCheckboxPairs wsCondition
    FixConditionE34E35Pair wsCondition
End Sub

' Checkbox click handler: flip paired control to opposite state.
Public Sub ConditionCheckboxExclusiveClick()
    On Error GoTo Done
    Dim callerName As String
    callerName = Application.Caller

    ' クリックされたチェックボックスは必ずアクティブシート上にある。
    ' 複数の条件書(コピー)間で同じ名前のチェックボックスが混在するため、全シートから
    ' 先頭一致で別シートに誤反応するのを防ぎ、アクティブシートに限定する。
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo Done
    If ws Is Nothing Then Exit Sub

    Dim clicked As Object
    On Error Resume Next
    Set clicked = ws.CheckBoxes(callerName)
    On Error GoTo Done
    If clicked Is Nothing Then Exit Sub

    Dim pair As Object
    Set pair = FindConditionCheckboxPair(ws, clicked)
    If pair Is Nothing Then Exit Sub

    Application.EnableEvents = False
    If clicked.Value = xlOn Then
        pair.Value = xlOff
    Else
        pair.Value = xlOn
    End If
    Application.EnableEvents = True

Done:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Return paired checkbox (D/X: opposite col + nearest Top / E34-E35: other row)
Private Function FindConditionCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim rowIndex As Long
    rowIndex = clicked.TopLeftCell.Row

    If IsConditionEVerticalCheckbox(clicked, ws) Then
        Set FindConditionCheckboxPair = FindConditionEVerticalCheckboxPair(ws, clicked)
        Exit Function
    End If

    If IsConditionDxSideCheckbox(clicked) Then
        If IsConditionDxExclusiveRow(rowIndex) Or IsConditionRow38BandRow(rowIndex) Then
            Set FindConditionCheckboxPair = FindConditionDxCheckboxPair(ws, clicked)
        End If
    End If
End Function

Private Function FindConditionEVerticalCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim clickedRow As Long
    clickedRow = ConditionEVerticalCheckboxRow(clicked, ws)
    If clickedRow = 0 Then Exit Function

    Dim targetRow As Long
    If clickedRow = CONDITION_E_PAIR_MIN_ROW Then
        targetRow = CONDITION_E_PAIR_MAX_ROW
    Else
        targetRow = CONDITION_E_PAIR_MIN_ROW
    End If

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If ConditionEVerticalCheckboxRow(cb, ws) = targetRow Then
                Set FindConditionEVerticalCheckboxPair = cb
                Exit Function
            End If
        End If
    Next cb
End Function

Private Function FindConditionDxCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim clickedCol As Long
    Dim clickedRow As Long
    Dim clickedTop As Double
    clickedCol = clicked.TopLeftCell.Column
    clickedRow = clicked.TopLeftCell.Row
    clickedTop = clicked.Top

    Dim clickedIsLeft As Boolean
    If IsConditionDxLeftSideColumn(clickedCol) Then
        clickedIsLeft = True
    ElseIf IsConditionDxRightSideColumn(clickedCol) Then
        clickedIsLeft = False
    Else
        Exit Function
    End If

    Dim rowMin As Long
    Dim rowMax As Long
    If clickedRow = 38 Or clickedRow = 39 Then
        rowMin = CONDITION_ROW38_BAND_MIN_ROW
        rowMax = CONDITION_ROW38_BAND_MAX_ROW
    Else
        rowMin = clickedRow
        rowMax = clickedRow
    End If

    Dim bestCb As Object
    Dim bestTopDist As Double
    bestTopDist = -1

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If cb.TopLeftCell.Row >= rowMin And cb.TopLeftCell.Row <= rowMax Then
                Dim cbCol As Long
                cbCol = cb.TopLeftCell.Column
                Dim cbIsLeft As Boolean
                Dim hasSide As Boolean
                hasSide = False
                If IsConditionDxLeftSideColumn(cbCol) Then
                    cbIsLeft = True
                    hasSide = True
                ElseIf IsConditionDxRightSideColumn(cbCol) Then
                    cbIsLeft = False
                    hasSide = True
                End If

                If hasSide And (cbIsLeft <> clickedIsLeft) Then
                    Dim topDist As Double
                    topDist = Abs(cb.Top - clickedTop)
                    If bestCb Is Nothing Or topDist < bestTopDist Then
                        Set bestCb = cb
                        bestTopDist = topDist
                    End If
                End If
            End If
        End If
    Next cb

    Set FindConditionDxCheckboxPair = bestCb
End Function

Private Sub NormalizeConditionCheckboxPairs(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = vbTextCompare

    Application.EnableEvents = False
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If IsConditionExclusiveCheckbox(cb, ws) Then
            If Not processed.Exists(cb.Name) Then
                Dim pair As Object
                Set pair = FindConditionCheckboxPair(ws, cb)
                If Not pair Is Nothing Then
                    processed.Add cb.Name, True
                    processed.Add pair.Name, True

                    If cb.Value = xlOn And pair.Value = xlOn Then
                        pair.Value = xlOff
                    ElseIf cb.Value <> xlOn And pair.Value <> xlOn Then
                        If IsConditionEVerticalCheckbox(cb, ws) Or _
                           IsConditionDxLeftSideColumn(cb.TopLeftCell.Column) Then
                            cb.Value = xlOn
                        Else
                            pair.Value = xlOn
                        End If
                    End If
                End If
            End If
        End If
    Next cb
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub FixConditionE34E35Pair(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim cb34 As Object
    Dim cb35 As Object
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        Select Case ConditionEVerticalCheckboxRow(cb, ws)
            Case CONDITION_E_PAIR_MIN_ROW
                Set cb34 = cb
            Case CONDITION_E_PAIR_MAX_ROW
                Set cb35 = cb
        End Select
    Next cb
    If cb34 Is Nothing Or cb35 Is Nothing Then Exit Sub

    On Error Resume Next
    cb34.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
    cb35.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
    On Error GoTo 0

    Application.EnableEvents = False
    If cb34.Value = xlOn And cb35.Value = xlOn Then
        cb35.Value = xlOff
    ElseIf cb34.Value <> xlOn And cb35.Value <> xlOn Then
        cb34.Value = xlOn
    End If
    Application.EnableEvents = True
End Sub

' 受注者用シートへのヘッダー転記。共通項目(S1/Q2/行20-34など)は
' ApplyContractorStyleCommonFields で3シート(受注者用/注文請書/支店控)共通に適用し、
' 受注者用固有の項目(E9=業者コード、A13=会社名)をここで追加転記する。
' 代表者・住所などは ApplyOfficeChiefBlock で処理する。
' 併せて排他選択(mod_BasicInfoExclusiveChoice)・支給貸与(mod_BasicInfoSupplyLoan)の再適用も行う。
' 改修時は3シート(受注者用/注文請書/支店控)すべてに影響することに留意する。
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    ' E9: 業者コード(施工会社ブロック16行目), 中央
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: 会社名(施工会社ブロック11行), 中央
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' M10/M11/M12: 代表者住所・氏名・電話番号(受注者用シートのみ)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    ' 行38-42(排他選択チェック・支給貸与品など)は受注者用シートにのみ存在する。
    ' 施工会社・内訳明細・支店控へ再転記(Reapplyは対象シートすべてへ適用される)。
    mod_BasicInfoExclusiveChoice.ReapplyExclusiveChoices wsInfo, vendorIndex
    mod_BasicInfoSupplyLoan.ReapplySupplyLoan wsInfo, vendorIndex

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' 受注者用/注文請書/支店控で共通の転記(S1:U1 注文番号(27) / Q2:V2 作成日(和暦) /
' 行20-34: E20 工事名称・E22 施工期間・G24/G26 工期(和暦)・Q22-24 契約金額系を参照)。
' 各3シートとも共通の項目なので、この関数にまとめて実装し重複を避ける。
Private Sub ApplyContractorStyleCommonFields(ByVal wsInfo As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal valueColumn As Long)
    ' S1:U1: 注文番号(施工会社ブロック27行目), 中央
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2:V2: 作成日(基本情報C2, 和暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E20: 工事名称(基本情報C10), 中央
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: 施工期間(基本情報C13), 中央
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: 工期 自(基本情報C15, 和暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: 工期 至(基本情報C16, 和暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22/Q23/Q24: 内訳明細シートが存在すればその「計」行(Q列)を数式で参照する。
    ' 内訳明細が存在しない場合は基本情報側の値をそのまま転記する。
    Dim breakdownName As String
    breakdownName = ResolveBreakdownSheetNameFromTarget(wsTarget)
    If Len(breakdownName) > 0 Then
        WriteHeaderFormulaRight wsTarget.Range("Q22"), BuildBreakdownQFormula(breakdownName, ContractorGrandTotalLabelText())
        WriteHeaderFormulaRight wsTarget.Range("Q23"), BuildBreakdownQFormula(breakdownName, ContractorNetTotalLabelText())
        WriteHeaderFormulaRight wsTarget.Range("Q24"), BuildBreakdownQFormula(breakdownName, ContractorTaxLabelText())
    Else
        WriteHeaderValueRight wsTarget.Range("Q22"), wsInfo.Cells(CONTRACTOR_CONTRACT_TOTAL_ROW, valueColumn).value
        WriteHeaderValueRight wsTarget.Range("Q23"), wsInfo.Cells(CONTRACTOR_CONTRACT_AMOUNT_ROW, valueColumn).value
        WriteHeaderValueRight wsTarget.Range("Q24"), wsInfo.Cells(CONTRACTOR_CONSUMPTION_TAX_ROW, valueColumn).value
    End If
End Sub

' wsTarget(処理対象シート)と同一グループの受注者用(○○)シートを解決して返す。
Private Function ResolveContractorSheetFromTarget(ByVal wsTarget As Worksheet) As Worksheet
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then Set ResolveContractorSheetFromTarget = ThisWorkbook.Worksheets(nm)
End Function

' 受注者用シートの値を参照専用で取得するヘルパー(値取得のみ、書込は行わない)
Private Function MirroredContractorText(ByVal wsContractor As Worksheet, ByVal address As String) As String
    If wsContractor Is Nothing Then Exit Function
    MirroredContractorText = CommonNzText(wsContractor.Range(address).MergeArea.Cells(1, 1).value)
End Function

' 受注者用 M10:V 各行と同じ書式でM:V列を結合するヘルパー(代表者ブロック用)
Private Sub EnsureOfficeChiefRowMerge(ByVal wsTarget As Worksheet, ByVal rowNo As Long)
    Const OFFICE_COL_START As Long = 13  ' M
    Const OFFICE_COL_END As Long = 22    ' V

    Dim mergeRange As Range
    Set mergeRange = wsTarget.Range(wsTarget.Cells(rowNo, OFFICE_COL_START), _
                                    wsTarget.Cells(rowNo, OFFICE_COL_END))
    mod_VendorBlockLayout.SafeUnmergeRange mergeRange
    On Error Resume Next
    mergeRange.Merge
    On Error GoTo 0
    With mergeRange.Cells(1, 1)
        .ShrinkToFit = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

' M10:代表者氏名 / M11:代表者住所+電話番号+誓約書 / M12:(受注者用のみ)住所・電話番号
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
    ' M10:V の各行(10～12行目)を結合し直し、代表者ブロックの表示崩れを防ぐ。
    Dim mergeRow As Long
    For mergeRow = 10 To 12
        EnsureOfficeChiefRowMerge wsTarget, mergeRow
    Next mergeRow

    Dim branchName As String, officeName As String
    branchName = CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value)
    officeName = CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value)

    Dim addr As String, title As String, chiefName As String
    Dim coreOffice As String, matchedOffice As String
    If Not mod_FillManagerName.GetOfficeChiefInfo(branchName, officeName, _
             addr, title, chiefName, coreOffice, matchedOffice) Then
        WriteHeaderValue wsTarget.Range("M10"), "", False
        WriteHeaderValue wsTarget.Range("M11"), "", False
        WriteHeaderValue wsTarget.Range("M12"), "", False
        Exit Sub
    End If

    Dim fw As String
    fw = ChrW$(&H3000)

    ' M10: 住所
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: 大鉄工業株式会社 + 全角空白 + 基幹出張所
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: 出張所=基幹出張所なら「役職 氏名」、不一致なら「出張所 役職 氏名」
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValueRight wsTarget.Range("M12"), m12
End Sub

' 右詰の値転記(BizUDゴシック適用)
Private Sub WriteHeaderValueRight(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' 左詰の値転記(BizUDゴシック適用)
Private Sub WriteHeaderValueLeft(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlLeft
    writeCell.VerticalAlignment = xlCenter
End Sub

' 下揃え・中央揃えの値転記(BizUDゴシック適用)
Private Sub WriteHeaderValueBottom(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If

    target.MergeArea.Font.Name = BASIC_INFO_REF_FONT_NAME
    target.MergeArea.HorizontalAlignment = xlCenter
    target.MergeArea.VerticalAlignment = xlBottom
End Sub

Private Sub WriteHeaderDateGregorian(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
        If IsDate(value) Then
            writeCell.NumberFormatLocal = "yyyy" & ChrW$(&H5E74) & "m" & ChrW$(&H6708) & "d" & ChrW$(&H65E5)
        End If
    End If
    ApplyHeaderCellFormat target, True
End Sub

' 注文請書シートへの転記。共通項目(S1:U1/Q2:V2/行20-34)は ApplyContractorStyleCommonFields。
'   G8:K9 業者コード(受注者用E9参照) / B12:K12 住所(基本情報14行目・左詰) /
'   B14:K14 会社名(基本情報11行目・左詰) / C15:I16 代表者名(基本情報12行目・右詰) /
'   M9:V9 住所(受注者用M10参照・左詰) / M10:V10 代表者氏名+電話番号+誓約書(受注者用M11参照・左詰) /
'   M11:V11 代表者(受注者用M12参照)+「殿」(右詰)
Private Sub ApplyAcceptanceHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    Dim wsContractor As Worksheet
    Set wsContractor = ResolveContractorSheetFromTarget(wsTarget)

    ' G8:K9: 業者コード(受注者用E9参照), 左詰
    WriteHeaderValue wsTarget.Range("G8"), MirroredContractorText(wsContractor, "E9"), True

    ' B12:K12: 住所(基本情報 施工会社ブロック14行目), 左詰
    WriteHeaderValueLeft wsTarget.Range("B12"), wsInfo.Cells(CONTRACTOR_ADDRESS_ROW, valueColumn).value

    ' B14:K14: 会社名(基本情報 施工会社ブロック11行目), 左詰
    WriteHeaderValueLeft wsTarget.Range("B14"), wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' C15:I16: 代表者名(基本情報 施工会社ブロック12行目), 右詰
    WriteHeaderValueRight wsTarget.Range("C15"), wsInfo.Cells(CONTRACTOR_REPRESENTATIVE_ROW, valueColumn).value

    ' M9:V9: 住所(受注者用M10参照), 左詰
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M10")

    ' M10:V10: 代表者氏名+電話番号(受注者用M11参照), 左詰
    WriteHeaderValueLeft wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M11")

    ' M11:V11: 代表者(受注者用M12参照) + 「殿」, 右詰
    Dim chiefText As String
    chiefText = MirroredContractorText(wsContractor, "M12")
    If Len(chiefText) > 0 Then chiefText = chiefText & ContractorHonorificSuffixText()
    WriteHeaderValueRight wsTarget.Range("M11"), chiefText

    mod_OrderTpl_Shared.OrderTplLog "ApplyAcceptanceHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyAcceptanceHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' 支店控シートへの転記。共通項目(S1:U1/Q2:V2/行20-34)は ApplyContractorStyleCommonFields。
'   E9:I10 業者コード(受注者用E9参照) / A13:I15 会社名(受注者用A13参照) /
'   M8:V8 住所(受注者用M10参照・左詰) / M9:V9 代表者氏名+電話番号(受注者用M11参照・左詰) /
'   M10:V10 代表者(受注者用M12参照・右詰)
Private Sub ApplyBranchCopyHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    Dim wsContractor As Worksheet
    Set wsContractor = ResolveContractorSheetFromTarget(wsTarget)

    ' E9:I10: 業者コード(受注者用E9参照), 左詰
    WriteHeaderValue wsTarget.Range("E9"), MirroredContractorText(wsContractor, "E9"), True

    ' A13:I15: 会社名(受注者用A13参照), 左詰
    WriteHeaderValue wsTarget.Range("A13"), MirroredContractorText(wsContractor, "A13"), True

    ' M8:V8: 住所(受注者用M10参照), 左詰
    WriteHeaderValueLeft wsTarget.Range("M8"), MirroredContractorText(wsContractor, "M10")

    ' M9:V9: 代表者氏名+電話番号(受注者用M11参照), 左詰
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M11")

    ' M10:V10: 代表者(受注者用M12参照), 右詰
    WriteHeaderValueRight wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M12")

    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' 別紙Ⅲシートへの転記(転記仕様が確定したらここへ実装する)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    On Error GoTo ErrorHandler
    ' O1: 施工会社名(基本情報 施工会社ブロック11行目)を右詰め・BIZ UDゴシックで転記
    Dim vendorCol As Long
    vendorCol = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    Dim vendorName As String
    vendorName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, vendorCol).Address)
    WriteAttachment3VendorName wsTarget.Range("O1"), vendorName

    ' 36～41行目の F/H列・J/L列の排他チェックボックスを設定
    SetupAttachment3CheckboxExclusivity wsTarget

    ' E49:I49～E52:I52 ダブルクリックによる産業廃棄物処理施設選択
    SetupAttachment3SanpaiFacilityDoubleClickHint wsTarget

    ' J54:M54 産廃行JR金額合計(工事シートに紐付く施工会社のみに反映)
    RefreshAttachment3SanpaiJrTotal wsTarget, wsInfo, vendorIndex
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyAttachment3Header error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' O1(施工会社名)の転記(右詰め・BIZ UDゴシック)
Private Sub WriteAttachment3VendorName(ByVal target As Range, ByVal vendorName As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    writeCell.NumberFormat = "@"
    If Len(Trim$(vendorName)) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = vendorName
    End If

    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' ===== 別紙Ⅲ 36～41行目 F/H列・J/L列 排他チェックボックス =====
' F列=有 / H列=無 のペアと、J列・L列のペアをそれぞれ排他制御する。
' さらに H列(無)がチェックされている行は J列・L列のどちらもチェックできない。
Public Sub SetupAttachment3CheckboxExclusivity(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        Dim r As Long, c As Long
        r = cb.TopLeftCell.Row
        c = cb.TopLeftCell.Column
        If r >= ATTACHMENT3_CHECK_ROW_MIN And r <= ATTACHMENT3_CHECK_ROW_MAX Then
            If c = ATTACHMENT3_COL_F Or c = ATTACHMENT3_COL_H Or _
               c = ATTACHMENT3_COL_J Or c = ATTACHMENT3_COL_L Then
                cb.OnAction = "'" & ThisWorkbook.Name & "'!Attachment3CheckboxClick"
            End If
        End If
    Next cb
    On Error GoTo 0

    NormalizeAttachment3CheckboxRows ws
End Sub

' チェックボックスクリック時の排他制御本体(各チェックボックスの OnAction から呼ばれる)
Public Sub Attachment3CheckboxClick()
    On Error GoTo Done

    Dim callerName As String
    callerName = Application.Caller

    ' クリックされたチェックボックスは必ずアクティブシート上にある。
    ' 複数の別紙Ⅲ(コピー)間で同じチェックボックスが混在するため、全シートから
    ' 先頭一致で他シートに誤反応するのを防ぎ、アクティブシートに限定する。
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo Done
    If ws Is Nothing Then Exit Sub
    If Not IsAttachment3Sheet(ws) Then Exit Sub

    Dim clicked As Object
    On Error Resume Next
    Set clicked = ws.CheckBoxes(callerName)
    On Error GoTo Done
    If clicked Is Nothing Then Exit Sub

    Dim rowIndex As Long, colIndex As Long
    rowIndex = clicked.TopLeftCell.Row
    colIndex = clicked.TopLeftCell.Column

    Application.EnableEvents = False

    Select Case colIndex
        Case ATTACHMENT3_COL_F
            If clicked.value = xlOn Then
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_H, False
            End If
        Case ATTACHMENT3_COL_H
            If clicked.value = xlOn Then
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_F, False
                ' H(無)をチェックした行は J列・L列のどちらもチェックさせない
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_J, False
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_L, False
            End If
        Case ATTACHMENT3_COL_J
            If clicked.value = xlOn Then
                If IsAttachment3CheckboxOn(ws, rowIndex, ATTACHMENT3_COL_H) Then
                    clicked.value = xlOff
                Else
                    SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_L, False
                End If
            End If
        Case ATTACHMENT3_COL_L
            If clicked.value = xlOn Then
                If IsAttachment3CheckboxOn(ws, rowIndex, ATTACHMENT3_COL_H) Then
                    clicked.value = xlOff
                Else
                    SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_J, False
                End If
            End If
    End Select

    Application.EnableEvents = True
    Exit Sub

Done:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Function FindAttachment3Checkbox(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Object
    Dim cb As Object
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.TopLeftCell.Row = rowIndex And cb.TopLeftCell.Column = colIndex Then
            Set FindAttachment3Checkbox = cb
            Exit Function
        End If
    Next cb
    On Error GoTo 0
End Function

Private Sub SetAttachment3Checkbox(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal turnOn As Boolean)
    Dim cb As Object
    Set cb = FindAttachment3Checkbox(ws, rowIndex, colIndex)
    If cb Is Nothing Then Exit Sub
    If turnOn Then
        cb.value = xlOn
    Else
        cb.value = xlOff
    End If
End Sub

Private Function IsAttachment3CheckboxOn(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    Dim cb As Object
    Set cb = FindAttachment3Checkbox(ws, rowIndex, colIndex)
    If cb Is Nothing Then Exit Function
    IsAttachment3CheckboxOn = (cb.value = xlOn)
End Function

' 既存の選択内容に矛盾(F/Hが両方チェック、H+J/Lの併用、J/L両方チェック等)がある場合のみ補正する。
' 新規に既定値を付与することはしない。
Private Sub NormalizeAttachment3CheckboxRows(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next

    Dim r As Long
    For r = ATTACHMENT3_CHECK_ROW_MIN To ATTACHMENT3_CHECK_ROW_MAX
        If IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_F) And _
           IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_H) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_H, False
        End If

        If IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_H) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_J, False
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_L, False
        ElseIf IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_J) And _
               IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_L) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_L, False
        End If
    Next r

    On Error GoTo 0
    Application.EnableEvents = prevEvents
End Sub

' ===== 別紙Ⅲ E49:I49～E52:I52 産業廃棄物処理施設選択 =====
' ダブルクリック対象範囲は既存のシート構成のままで動作するため、現時点では特別な設定は不要。
' (対象判定・選択処理は ThisWorkbook 側から mod_OrderTpl_Header の関数を直接呼び出す)
Private Sub SetupAttachment3SanpaiFacilityDoubleClickHint(ByVal ws As Worksheet)
    ' 予約フック(将来、追加設定が必要になった場合のために用意している)
End Sub

' Attachment3(別紙Ⅲ)シートかどうかを判定する
Public Function IsAttachment3Sheet(ByVal ws As Worksheet) As Boolean
    IsAttachment3Sheet = False
    If ws Is Nothing Then Exit Function
    Dim baseName As String, aliasText As String
    If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then
        IsAttachment3Sheet = (StrComp(baseName, mod_OrderTpl_Shared.OrderTplBaseNameAttachment3Text(), vbTextCompare) = 0)
    End If
End Function

' E49:I49～E52:I52 (産業廃棄物処理施設選択セル)かどうかを判定する
Public Function IsAttachment3SanpaiFacilityTarget(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    IsAttachment3SanpaiFacilityTarget = False
    If ws Is Nothing Or target Is Nothing Then Exit Function
    If Not IsAttachment3Sheet(ws) Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)

    If topLeft.Row < ATTACHMENT3_SANPAI_ROW_MIN Or topLeft.Row > ATTACHMENT3_SANPAI_ROW_MAX Then Exit Function
    IsAttachment3SanpaiFacilityTarget = (topLeft.Column = ATTACHMENT3_SANPAI_NAME_COL)
End Function

' frmSubconSelector.frm を再利用して産業廃棄物処理施設を選択させ、
' E列(名称)・J列(対応値、業者マスタB列)へ書き込む。
Public Sub RequestAttachment3SanpaiFacilitySelection(ByVal ws As Worksheet, ByVal target As Range)
    If ws Is Nothing Or target Is Nothing Then Exit Sub

    Dim rowIndex As Long
    rowIndex = target.Row

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim branchName As String
    branchName = GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL)

    Dim names As Collection
    Dim companionMap As Object
    Set names = New Collection
    Set companionMap = CreateObject("Scripting.Dictionary")
    companionMap.CompareMode = vbTextCompare

    If Not LoadSanpaiFacilityMasterRows(branchName, names, companionMap) Then
        MsgBox "産業廃棄物処理施設のマスタを取得できませんでした。", vbExclamation
        Exit Sub
    End If
    If names.Count = 0 Then
        MsgBox "産業廃棄物処理施設のマスタにデータがありません。", vbExclamation
        Exit Sub
    End If

    Dim arr() As String
    ReDim arr(1 To names.Count)
    Dim i As Long
    For i = 1 To names.Count
        arr(i) = CStr(names(i))
    Next i

    Dim f As New frmSubconSelector
    f.Caption = Attachment3SanpaiFacilityCaptionText()
    f.SetCompanies arr
    f.Show vbModal

    Dim confirmed As Boolean, chosen As String
    confirmed = f.confirmed
    chosen = f.SelectedCompany
    Unload f
    If Not confirmed Or chosen = "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    ws.Cells(rowIndex, ATTACHMENT3_SANPAI_NAME_COL).MergeArea.Cells(1, 1).value = chosen

    Dim companionValue As String
    If companionMap.Exists(chosen) Then companionValue = CStr(companionMap(chosen))
    ws.Cells(rowIndex, ATTACHMENT3_SANPAI_VALUE_COL).MergeArea.Cells(1, 1).value = companionValue

    Application.EnableEvents = prevEvents
End Sub

' 業者マスタ(全社版).xlsx の支店名シートから A59以降(名称)・B59以降(対応値)を読み込む。
Private Function LoadSanpaiFacilityMasterRows(ByVal branchName As String, _
                                              ByVal names As Collection, _
                                              ByVal companionMap As Object) As Boolean
    Dim connection As Object
    Dim recordset As Object
    On Error GoTo Cleanup

    If Trim$(branchName) = "" Then GoTo Cleanup

    Dim masterPath As String
    masterPath = mod_Construction_BasicTotals.ResolveVendorMasterPath()
    If masterPath = "" Then GoTo Cleanup

    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then GoTo Cleanup

    Dim actualSheetName As String
    actualSheetName = mod_Construction_OutputLayout.FindAdoWorksheetName(connection, branchName)
    If actualSheetName = "" Then GoTo Cleanup

    Dim tableRangeName As String
    tableRangeName = "[" & Replace$(actualSheetName, "]", "]]") & "$A" & _
                      SANPAI_FACILITY_MASTER_START_ROW & ":B10000]"

    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT F1, F2 FROM " & tableRangeName, connection, 0, 1, 1

    Do Until recordset.EOF
        Dim nameVal As String, companionVal As String
        nameVal = Trim$(CommonNzText(recordset.fields(0).value))
        companionVal = Trim$(CommonNzText(recordset.fields(1).value))
        If nameVal <> "" Then
            names.Add nameVal
            If Not companionMap.Exists(nameVal) Then companionMap.Add nameVal, companionVal
        End If
        recordset.MoveNext
    Loop

    LoadSanpaiFacilityMasterRows = True

Cleanup:
    If Err.Number <> 0 Then Err.Clear
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
End Function

' "産業廃棄物処理施設選択"
Private Function Attachment3SanpaiFacilityCaptionText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H7523, &H696D, &H5EC3, &H68C4, &H7269, _
                                      &H51E6, &H7406, &H65BD, &H8A2D, &H9078, &H629E)
    End If
    Attachment3SanpaiFacilityCaptionText = cached
End Function

' ===== 別紙Ⅲ J54:M54 産廃行JR金額合計 =====
' 施行指示書(工事)/施行通知書(工事)シートの産廃行(施工会社が当初選択できなかった行)の
' JR金額の合計を、対応する別紙Ⅲの J54:M54(結合セル)へ転記する。
Private Sub RefreshAttachment3SanpaiJrTotal(ByVal ws As Worksheet, _
                                            ByVal wsInfo As Worksheet, _
                                            ByVal vendorIndex As Long)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Done

    Dim jrTotalArea As Range
    Set jrTotalArea = ws.Range("J54").MergeArea      ' J54:M54(結合セル)
    Dim writeCell As Range
    Set writeCell = jrTotalArea.Cells(1, 1)

    ' 施行指示書(工事)/施行通知書(工事)いずれにも紐付いていない施工会社の別紙Ⅲには
    ' 産廃JR金額合計を反映しない(結合セル全体をクリア)。
    ' この結合セルの一部(Cells(1,1))に対してのみ ClearContents すると 1004
    '   「複数結合セルの一部だけには行えません」になるため、MergeArea 全体をクリアする。
    If Not IsAttachment3VendorLinkedToWorks(wsInfo, vendorIndex) Then
        jrTotalArea.ClearContents
        Exit Sub
    End If

    Dim total As Double
    total = mod_Construction_BasicTotals.SumSanpaiJrAmount()

    writeCell.value = Int(total)   ' 桁切り(端数切り捨て)で入力する
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
    Exit Sub

Done:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAttachment3SanpaiJrTotal error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' 当該別紙IIIの業者が施行指示書(工事)/通知書(工事)のA列に紐付いているか
Private Function IsAttachment3VendorLinkedToWorks(ByVal wsInfo As Worksheet, _
                                                  ByVal vendorIndex As Long) As Boolean
    On Error GoTo Done
    If wsInfo Is Nothing Then Exit Function

    Dim branchName As String
    branchName = GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL)

    Dim vendorCol As Long
    vendorCol = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    Dim companyName As String
    companyName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, vendorCol).Address)
    If companyName = "" Then Exit Function

    IsAttachment3VendorLinkedToWorks = _
        mod_Construction_BasicTotals.IsVendorSelectedOnWorksSheet(branchName, companyName)

Done:
End Function

' 通常セルの値の転記(フォント適用、必要に応じて上下・左右中央揃え)
Private Sub WriteHeaderValue(ByVal target As Range, ByVal value As Variant, ByVal centered As Boolean)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If

    ApplyHeaderCellFormat target, centered
End Sub

' 文字列としての転記(部店コード等、日付誤変換を防ぐ)
Private Sub WriteHeaderText(ByVal target As Range, ByVal textValue As String, ByVal centered As Boolean)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    writeCell.NumberFormat = "@"
    If Len(Trim$(textValue)) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = textValue
    End If

    ApplyHeaderCellFormat target, centered
End Sub

' 日付の転記(和暦表示形式、結合セルのため中央揃え)
Private Sub WriteHeaderDate(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
        If IsDate(value) Then
            writeCell.NumberFormat = mod_OrderTpl_Shared.OrderTplEraDateNumberFormatText()
        End If
    End If

    ApplyHeaderCellFormat target, True
End Sub

Private Sub ApplyHeaderCellFormat(ByVal target As Range, ByVal centered As Boolean)
    target.MergeArea.Font.Name = BASIC_INFO_REF_FONT_NAME
    If centered Then
        target.MergeArea.HorizontalAlignment = xlCenter
        target.MergeArea.VerticalAlignment = xlCenter
    End If
End Sub

' 受注者用シートwsTargetと同一グループの内訳明細シートを解決して返す
Private Function ResolveBreakdownSheetNameFromTarget(ByVal wsTarget As Worksheet) As String
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then ResolveBreakdownSheetNameFromTarget = nm
End Function

' 内訳明細シートから、指定ラベル行のQ列を参照する数式を組み立てる(該当なければ空欄)
Private Function BuildBreakdownQFormula(ByVal sheetName As String, ByVal labelText As String) As String
    Dim q As String
    q = "'" & Replace$(sheetName, "'", "''") & "'"
    BuildBreakdownQFormula = "=IFERROR(INDEX(" & q & "!Q:Q,MATCH(""" & labelText & """," & q & "!A:A,0)),"""")"
End Function

' 数式(右詰)の転記(BizUDゴシック・桁区切り書式)
Private Sub WriteHeaderFormulaRight(ByVal target As Range, ByVal formulaText As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    writeCell.Formula = formulaText
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' "合計"
Private Function ContractorGrandTotalLabelText() As String
    ContractorGrandTotalLabelText = ChrW$(&H5408) & ChrW$(&H8A08)
End Function

' "計"
Private Function ContractorNetTotalLabelText() As String
    ContractorNetTotalLabelText = ChrW$(&H8A08)
End Function

' "消費税"
Private Function ContractorTaxLabelText() As String
    ContractorTaxLabelText = ChrW$(&H6D88) & ChrW$(&H8CBB) & ChrW$(&H7A0E)
End Function

' "　　殿"(全角空白2つ + 殿。M11:V11 の末尾へ付与する)
Private Function ContractorHonorificSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H6BBF)
    End If
    ContractorHonorificSuffixText = cached
End Function
