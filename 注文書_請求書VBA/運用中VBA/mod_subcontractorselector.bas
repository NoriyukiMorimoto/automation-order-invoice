Option Explicit

Private Const COL_VENDOR As Long = 1
Private Const COL_SEIRI As Long = 2
Private Const DATA_START_ROW As Long = 2

Private Const SANPAI_KEYWORD As String = "産廃処理"
Private Const COL_TYPE As Long = 3

Private Const WELD_COL_WELDING_VENDOR As Long = 1
Private Const WELD_COL_TRACK_VENDOR As Long = 2

Private Const VENDOR_NAME_ROW As Long = 11
Private Const VENDOR_WORK_TYPE_ROW As Long = 10
Private Const VENDOR_FIRST_COL As Long = 6
Private Const VENDOR_STEP_COLS As Long = 3
Private Const VENDOR_MAX_BLOCKS As Long = 20
Private Const VENDOR_COUNT_CELL As String = "F9"

Private Const WELDING_WORK_TYPE_KEYWORD As String = "溶接工事"
Private Const TRACK_WORK_TYPE_KEYWORD As String = "軌道工事"

Private Const HELPER_COL As Long = 100
Private Const LOG_TAG As String = "[SubconSelector]"

Private mSelectionScheduled As Boolean
Private mLastTrigger As String
Private mCurrentStep As String
Private mLogSessionStarted As Boolean

Private mPendingWs As Worksheet
Private mPendingVendorCol As Long
Private mPendingRows As Collection
Private mPendingContextReady As Boolean

Public Sub SetSubcontractorTrigger(ByVal triggerName As String)
    mLastTrigger = Trim$(triggerName)
End Sub

Public Sub OpenSubcontractorSelectionLogFile()
    SubconLog "OpenSubcontractorSelectionLogFile"
    On Error Resume Next
    Shell "notepad.exe " & Chr$(34) & mod_DebugLog.GetPersistedLogFilePath() & Chr$(34), vbNormalFocus
    On Error GoTo 0
End Sub

Private Sub SubconLog(ByVal msg As String)
    If Not mLogSessionStarted Then
        mLogSessionStarted = True
        mod_DebugLog.LogPersist LOG_TAG & " logFile=" & mod_DebugLog.GetPersistedLogFilePath()
        mod_DebugLog.LogPersist LOG_TAG & " workbook=" & ThisWorkbook.Name
    End If
    mod_DebugLog.LogPersist LOG_TAG & " " & msg
End Sub

Private Sub SubconLogErr(Optional ByVal stepName As String = "")
    Dim stepText As String
    If Len(stepName) > 0 Then
        stepText = stepName
    Else
        stepText = mCurrentStep
    End If
    SubconLog "ERR step=" & stepText & " Err=" & Err.Number & " " & Err.Description
End Sub

Private Function DescribeRowCollection(ByVal rows As Collection) As String
    If rows Is Nothing Then
        DescribeRowCollection = "(nothing)"
        Exit Function
    End If
    If rows.Count = 0 Then
        DescribeRowCollection = "(empty)"
        Exit Function
    End If

    Dim parts As String
    Dim rowRef As Variant
    For Each rowRef In rows
        parts = parts & IIf(parts = "", "", ",") & CStr(rowRef)
    Next rowRef
    DescribeRowCollection = parts
End Function

Public Sub ClearSubcontractorSelectionContext()
    Set mPendingWs = Nothing
    Set mPendingRows = Nothing
    mPendingVendorCol = 0
    mPendingContextReady = False
End Sub

'  イベント発火中に Target / Selection を確定させ、OnTime 後も同じ行を使う。
Public Sub PrepareSubcontractorSelection(ByVal ws As Worksheet, _
                                         ByVal Target As Range, _
                                         Optional ByVal multiSel As Range = Nothing)
    mCurrentStep = "Prepare"
    SubconLog "Prepare start trigger=[" & mLastTrigger & "] sheet=" & ws.Name & _
              " target=" & Target.Address(False, False) & _
              " filter=" & CStr(ws.AutoFilterMode)

    ClearSubcontractorSelectionContext
    If ws Is Nothing Or Target Is Nothing Then
        SubconLog "Prepare abort: ws or Target is Nothing"
        Exit Sub
    End If

    Dim seiriCol As Long
    seiriCol = SeiriColumn(ws)

    Dim lastRow As Long
    lastRow = GetSheetDataLastRow(ws)
    SubconLog "Prepare lastRow=" & lastRow & " seiriCol=" & seiriCol
    If lastRow < DATA_START_ROW Then
        SubconLog "Prepare abort: no data rows"
        Exit Sub
    End If

    mPendingVendorCol = ResolveTargetVendorColumnFromCell(ws, Target)
    If mPendingVendorCol = 0 Then
        SubconLog "Prepare abort: vendor column unresolved col=" & Target.Column
        Exit Sub
    End If

    Set mPendingWs = ws
    Set mPendingRows = New Collection

    If Not multiSel Is Nothing Then
        If multiSel.Worksheet Is ws And multiSel.CountLarge > 1 Then
            Dim multiHit As Range
            On Error Resume Next
            Set multiHit = Application.Intersect(multiSel, VendorSelectionColumnsRange(ws))
            On Error GoTo 0
            If Not multiHit Is Nothing Then
                AddEligibleVendorRowsFromRange ws, multiHit, seiriCol, lastRow, mPendingRows
            End If
            SubconLog "Prepare multiSel rows=" & DescribeRowCollection(mPendingRows)
        End If
    End If

    If mPendingRows.Count = 0 Then
        If Target.CountLarge = 1 And IsVendorSelectionColumn(ws, Target.Column) Then
            SubconLog "Prepare singleCell row=" & Target.Row
            AddEligibleVendorRowDirect ws, Target.Row, seiriCol, lastRow, mPendingRows, True
        Else
            BuildPendingRowsFromTarget ws, Target, seiriCol, lastRow
        End If
    End If

    mPendingContextReady = (mPendingRows.Count > 0)
    SubconLog "Prepare done vendorCol=" & mPendingVendorCol & _
              " rows=" & DescribeRowCollection(mPendingRows) & _
              " ready=" & CStr(mPendingContextReady)
End Sub

'  ダブルクリック・右クリックメニューからの呼び出し用エントリーポイント。
'  イベント/メニューの最中にモーダルフォームを同期表示するとハングするため、
'  Application.OnTime で UI 解放後に起動を遅延させる。
Public Sub RequestSubcontractorSelection()
    mCurrentStep = "Request"
    If mSelectionScheduled Then
        SubconLog "Request skip: already scheduled trigger=[" & mLastTrigger & "]"
        Exit Sub
    End If
    mSelectionScheduled = True

    Dim procPrimary As String
    Dim procFallback As String
    procPrimary = "'" & ThisWorkbook.Name & "'!RunScheduledSubcontractorSelection"
    procFallback = "mod_subcontractorselector.RunScheduledSubcontractorSelection"
    SubconLog "Request schedule trigger=[" & mLastTrigger & "] proc=" & procPrimary

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:=procPrimary, _
                       Schedule:=True
    If Err.Number <> 0 Then
        SubconLog "Request OnTime primary failed Err=" & Err.Number & " " & Err.Description
        Err.Clear
        SubconLog "Request schedule fallback proc=" & procFallback
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:=procFallback, _
                           Schedule:=True
        If Err.Number <> 0 Then
            SubconLog "Request OnTime fallback failed Err=" & Err.Number & " " & Err.Description
            Err.Clear
            mSelectionScheduled = False
            ClearSubcontractorSelectionContext
        End If
    Else
        SubconLog "Request OnTime primary scheduled"
    End If
    On Error GoTo 0
End Sub

Public Sub RunScheduledSubcontractorSelection()
    mCurrentStep = "RunScheduled"
    SubconLog "RunScheduled start trigger=[" & mLastTrigger & "]"
    mSelectionScheduled = False
    DoEvents
    SubconLog "RunScheduled after DoEvents"
    SelectSubcontractorForSelection
    SubconLog "RunScheduled end"
End Sub

Public Function GetSubcontractorList() As Variant
    GetSubcontractorList = GetSubcontractorListByWorkType("")
End Function

Public Function GetSubcontractorListByWorkType(ByVal workTypeKeyword As String, _
                                               Optional ByVal excludeWorkTypeKeyword As String = "") As Variant
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim names As Collection
    Set names = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim vendorNameMap As Object
    Set vendorNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)
    If vendorNameMap Is Nothing Then Exit Function

    Dim blockIndex As Long
    For blockIndex = 1 To VENDOR_MAX_BLOCKS
        Dim col As Long
        col = VENDOR_FIRST_COL + ((blockIndex - 1) * VENDOR_STEP_COLS)

        If workTypeKeyword <> "" Then
            If Not VendorBlockMatchesWorkType(wsInfo, col, workTypeKeyword) Then GoTo NextVendorBlock
        Else
            Dim vendorCount As Long
            vendorCount = CLng(Val(StrConv(CStr(CommonNzText(wsInfo.Range(VENDOR_COUNT_CELL).value)), vbNarrow)))
            If vendorCount < 1 Then vendorCount = 1
            If vendorCount > VENDOR_MAX_BLOCKS Then vendorCount = VENDOR_MAX_BLOCKS
            If blockIndex > vendorCount Then Exit For
            If excludeWorkTypeKeyword <> "" Then
                If VendorBlockMatchesWorkType(wsInfo, col, excludeWorkTypeKeyword) Then GoTo NextVendorBlock
            End If
        End If

        Dim nm As String
        Dim mappedName As String
        nm = Trim$(CommonNzText(wsInfo.Cells(VENDOR_NAME_ROW, col).value))
        If nm <> "" Then
            nm = CommonNormalizeText(nm)
            If vendorNameMap.Exists(nm) Then
                mappedName = Trim$(CommonNzText(vendorNameMap(nm)))
                If mappedName <> "" And Not seen.Exists(mappedName) Then
                    seen.Add mappedName, True
                    names.Add mappedName
                End If
            End If
        End If
NextVendorBlock:
    Next blockIndex

    If names.Count = 0 Then Exit Function

    Dim arr() As String
    Dim i As Long
    ReDim arr(1 To names.Count)
    For i = 1 To names.Count
        arr(i) = CStr(names(i))
    Next i
    GetSubcontractorListByWorkType = arr
End Function

Private Function GetSubcontractorListForVendorSheet(ByVal ws As Worksheet) As Variant
    If mod_Construction_Order_Import.IsConstructionOrderWorksOutputSheet(ws) Then
        GetSubcontractorListForVendorSheet = GetSubcontractorListByWorkType( _
            "", WELDING_WORK_TYPE_KEYWORD)
    Else
        GetSubcontractorListForVendorSheet = GetSubcontractorList()
    End If
End Function

Private Function VendorBlockMatchesWorkType(ByVal wsInfo As Worksheet, _
                                            ByVal valueCol As Long, _
                                            ByVal workTypeKeyword As String) As Boolean
    Dim workTypeText As String
    workTypeText = CommonRemoveAllSpaces(CommonNormalizeText( _
        CommonNzText(wsInfo.Cells(VENDOR_WORK_TYPE_ROW, valueCol).value)))
    VendorBlockMatchesWorkType = (workTypeText <> "") And _
        (InStr(1, workTypeText, workTypeKeyword, vbTextCompare) > 0)
End Function

Public Sub ApplySubcontractorDropdowns(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = GetSheetDataLastRow(ws)
    If lastRow < DATA_START_ROW Then Exit Sub

    Dim prevEvents As Boolean
    Dim errorDescription As String
    prevEvents = Application.EnableEvents
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        ApplyVendorColumnDropdown ws, WELD_COL_WELDING_VENDOR, WELDING_WORK_TYPE_KEYWORD, lastRow
        ApplyVendorColumnDropdown ws, WELD_COL_TRACK_VENDOR, TRACK_WORK_TYPE_KEYWORD, lastRow
    Else
        Dim names As Variant
        names = GetSubcontractorListForVendorSheet(ws)
        If IsArray(names) Then
            ApplyVendorColumnDropdownWithNames ws, COL_VENDOR, names, lastRow
        End If
    End If

Cleanup:
    Application.EnableEvents = prevEvents
    If errorDescription <> "" Then
        MsgBox "施工会社のドロップダウンを設定できませんでした。" & vbCrLf & _
               errorDescription, vbExclamation
    End If
    Exit Sub

ErrorHandler:
    errorDescription = Err.Description
    Resume Cleanup
End Sub

Private Sub ApplyVendorColumnDropdown(ByVal ws As Worksheet, _
                                      ByVal vendorColumn As Long, _
                                      ByVal workTypeKeyword As String, _
                                      ByVal lastRow As Long)
    Dim names As Variant
    names = GetSubcontractorListByWorkType(workTypeKeyword)
    If Not IsArray(names) Then Exit Sub
    ApplyVendorColumnDropdownWithNames ws, vendorColumn, names, lastRow
End Sub

Private Sub ApplyVendorColumnDropdownWithNames(ByVal ws As Worksheet, _
                                               ByVal vendorColumn As Long, _
                                               ByVal names As Variant, _
                                               ByVal lastRow As Long)
    Dim listFormula As String
    listFormula = BuildValidationListFormula(ws, names)
    If listFormula = "" Then Exit Sub

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If Trim$(CommonNzText(ws.Cells(r, SeiriColumn(ws)).value)) <> "" Then
            If Not IsSanpaiRow(ws, r) Then
                With ws.Cells(r, vendorColumn).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:=listFormula
                    .IgnoreBlank = True
                    .InCellDropdown = False
                    .ShowError = False
                End With
            End If
        End If
    Next r
End Sub

Private Function BuildValidationListFormula(ByVal ws As Worksheet, ByVal names As Variant) As String
    Dim joined As String, i As Long, hasComma As Boolean
    For i = LBound(names) To UBound(names)
        joined = joined & IIf(joined = "", "", ",") & CStr(names(i))
        If InStr(1, CStr(names(i)), ",") > 0 Then hasComma = True
    Next i

    If Len(joined) <= 255 And Not hasComma Then
        BuildValidationListFormula = joined
        Exit Function
    End If

    Dim n As Long
    n = UBound(names) - LBound(names) + 1

    Dim helper As Range
    Set helper = ws.Range(ws.Cells(1, HELPER_COL), ws.Cells(n, HELPER_COL))
    helper.ClearContents

    Dim arr2() As Variant
    ReDim arr2(1 To n, 1 To 1)
    For i = 1 To n
        arr2(i, 1) = CStr(names(LBound(names) + i - 1))
    Next i
    helper.value = arr2
    ws.Columns(HELPER_COL).Hidden = True

    BuildValidationListFormula = "=" & helper.Address(True, True)
End Function

Public Sub SelectSubcontractorForSelection()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    On Error GoTo ErrorHandler

    mCurrentStep = "SelectStart"
    SubconLog "Select start trigger=[" & mLastTrigger & "] pending=" & CStr(mPendingContextReady)

    Dim ws As Worksheet
    Dim seiriCol As Long
    Dim lastRow As Long
    Dim targetColumn As Long
    Dim targetRows As Collection
    Dim usePendingContext As Boolean

    usePendingContext = mPendingContextReady And Not mPendingWs Is Nothing
    If usePendingContext Then
        Set ws = mPendingWs
        Set targetRows = mPendingRows
        targetColumn = mPendingVendorCol
        seiriCol = SeiriColumn(ws)
        lastRow = GetSheetDataLastRow(ws)
        ClearSubcontractorSelectionContext
        SubconLog "Select use pending sheet=" & ws.Name & " rows=" & DescribeRowCollection(targetRows)
    Else
        Set ws = ActiveSheet
        seiriCol = SeiriColumn(ws)
        lastRow = GetSheetDataLastRow(ws)
        Set targetRows = New Collection
        SubconLog "Select use active sheet=" & ws.Name & " lastRow=" & lastRow
    End If

    If lastRow < DATA_START_ROW Then
        SubconLog "Select abort: no data"
        MsgBox "対象データがありません。", vbExclamation
        GoTo CleanExit
    End If

    If Not usePendingContext Then
        mCurrentStep = "ResolveTargetColumn"
        targetColumn = ResolveTargetVendorColumn(ws, seiriCol, lastRow)
        If targetColumn = 0 Then
            SubconLog "Select abort: target column unresolved"
            If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
                MsgBox "溶接会社(A列)または軌道手元会社(B列)を選択した状態で実行してください。", vbExclamation
            Else
                MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
            End If
            GoTo CleanExit
        End If

        mCurrentStep = "CollectTargetRows"
        If TypeName(Selection) = "Range" Then
            Dim vendorHit As Range
            On Error Resume Next
            Set vendorHit = Application.Intersect(Selection, VendorSelectionColumnsRange(ws))
            On Error GoTo ErrorHandler

            If Not vendorHit Is Nothing Then
                AddEligibleVendorRowsFromRange ws, vendorHit, seiriCol, lastRow, targetRows
            End If
        End If

        If targetRows.Count = 0 Then
            If ActiveCell.Row >= DATA_START_ROW And ActiveCell.Row <= lastRow Then
                AddEligibleVendorRowDirect ws, ActiveCell.Row, seiriCol, lastRow, targetRows, True
            End If
        End If
        SubconLog "Select collected rows=" & DescribeRowCollection(targetRows)
    ElseIf targetColumn = 0 Then
        SubconLog "Select abort: pending vendor column missing"
        MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
        GoTo CleanExit
    End If

    If targetRows.Count = 0 Then
        SubconLog "Select abort: no target rows"
        If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
            MsgBox "会社を設定する行を選択してください(C列に整理番号がある行)。", vbExclamation
        Else
            MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
        End If
        GoTo CleanExit
    End If

    mCurrentStep = "GetSubcontractorList"
    Dim names As Variant
    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        If targetColumn = WELD_COL_WELDING_VENDOR Then
            names = GetSubcontractorListByWorkType(WELDING_WORK_TYPE_KEYWORD)
        Else
            names = GetSubcontractorListByWorkType(TRACK_WORK_TYPE_KEYWORD)
        End If
    Else
        names = GetSubcontractorListForVendorSheet(ws)
    End If

    If Not IsArray(names) Then
        SubconLog "Select abort: subcontractor list empty"
        MsgBox "基本情報の施工会社に対応する業者マスタA列の候補が見つかりません。", vbExclamation
        GoTo CleanExit
    End If
    SubconLog "Select companyCount=" & (UBound(names) - LBound(names) + 1)

    mCurrentStep = "ShowForm"
    Dim f As frmSubconSelector
    Set f = New frmSubconSelector
    f.SetCompanies names
    SubconLog "Select before ShowForm"
    f.Show vbModal
    SubconLog "Select after ShowForm"

    Dim confirmed As Boolean, chosen As String
    confirmed = f.confirmed
    chosen = f.SelectedCompany
    Unload f
    If Not confirmed Or chosen = "" Then
        SubconLog "Select cancelled confirmed=" & CStr(confirmed)
        GoTo CleanExit
    End If
    SubconLog "Select chosen=[" & chosen & "]"

    Application.EnableEvents = False

    mCurrentStep = "WriteVendor"
    Dim rIdx As Variant
    For Each rIdx In targetRows
        ws.Cells(CLng(rIdx), targetColumn).value = chosen
    Next rIdx
    SubconLog "Select vendor written rows=" & DescribeRowCollection(targetRows)

    mCurrentStep = "RefreshPrices"
    If ws.AutoFilterMode Then
        mod_Construction_Order_Import.RefreshSubcontractorPriceColumns ws
    Else
        mod_Construction_Order_Import.RefreshSubcontractorPriceColumns ws, targetRows
    End If
    SubconLog "Select refresh done filter=" & CStr(ws.AutoFilterMode)

    ws.Columns(targetColumn).AutoFit

    Application.EnableEvents = prevEvents

    SubconLog "Select complete rows=" & targetRows.Count & " company=[" & chosen & "]"

    ' 施工会社割当てを各社の内訳明細へ自動反映
    mCurrentStep = "RefreshOrderDetails"
    mod_OrderTpl_Generate.RefreshAllVendorOrderDetailsSilent
    MsgBox "選択された" & targetRows.Count & "行に「" & chosen & "」の施工単価を設定しました。", _
           vbInformation
    GoTo CleanExit

ErrorHandler:
    SubconLogErr mCurrentStep
    ClearSubcontractorSelectionContext
    Application.EnableEvents = prevEvents
    MsgBox "施工会社別の単価・金額列を更新できませんでした。" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           ChrW$(&H8A73) & ChrW$(&H7D30) & ChrW$(&H306F) & " " & mod_DebugLog.GetPersistedLogFilePath() & ChrW$(&H3092) & ChrW$(&H78BA) & ChrW$(&H8A8D) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002), vbExclamation
    Exit Sub

CleanExit:
    Application.EnableEvents = prevEvents
    Exit Sub
End Sub

Private Function ResolveTargetVendorColumnFromCell(ByVal ws As Worksheet, ByVal Target As Range) As Long
    If Target Is Nothing Then Exit Function

    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        If Target.Column = WELD_COL_WELDING_VENDOR Then
            ResolveTargetVendorColumnFromCell = WELD_COL_WELDING_VENDOR
            Exit Function
        End If
        If Target.Column = WELD_COL_TRACK_VENDOR Then
            ResolveTargetVendorColumnFromCell = WELD_COL_TRACK_VENDOR
            Exit Function
        End If
        Exit Function
    End If

    If IsVendorSelectionColumn(ws, Target.Column) Then
        ResolveTargetVendorColumnFromCell = COL_VENDOR
    End If
End Function

Private Function IsVendorSelectionColumn(ByVal ws As Worksheet, ByVal columnIndex As Long) As Boolean
    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        IsVendorSelectionColumn = (columnIndex = WELD_COL_WELDING_VENDOR Or columnIndex = WELD_COL_TRACK_VENDOR)
    Else
        IsVendorSelectionColumn = (columnIndex = COL_VENDOR)
    End If
End Function

Private Sub BuildPendingRowsFromTarget(ByVal ws As Worksheet, _
                                       ByVal Target As Range, _
                                       ByVal seiriCol As Long, _
                                       ByVal lastRow As Long)
    If Not IsVendorSelectionColumn(ws, Target.Column) Then Exit Sub

    Dim c As Range
    For Each c In Target.Cells
        AddEligibleVendorRowDirect ws, c.Row, seiriCol, lastRow, mPendingRows, False
    Next c
End Sub

Private Function VendorSelectionColumnsRange(ByVal ws As Worksheet) As Range
    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        Set VendorSelectionColumnsRange = ws.Columns("A:B")
    Else
        Set VendorSelectionColumnsRange = ws.Columns(COL_VENDOR)
    End If
End Function

Private Function IsDataRowVisible(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    On Error GoTo NotVisible

    If ws.Rows(rowIndex).Hidden Then GoTo NotVisible

    If ws.AutoFilterMode Then
        If ws.Rows(rowIndex).RowHeight <= 0# Then GoTo NotVisible
    End If

    IsDataRowVisible = True
    Exit Function

NotVisible:
    IsDataRowVisible = False
End Function

Private Sub AddEligibleVendorRowDirect(ByVal ws As Worksheet, _
                                       ByVal rowIndex As Long, _
                                       ByVal seiriCol As Long, _
                                       ByVal lastRow As Long, _
                                       ByVal targetRows As Collection, _
                                       ByVal assumeVisible As Boolean)
    If rowIndex < DATA_START_ROW Or rowIndex > lastRow Then Exit Sub
    If Not assumeVisible Then
        If Not IsDataRowVisible(ws, rowIndex) Then Exit Sub
    End If

    Dim i As Long
    For i = 1 To targetRows.Count
        If CLng(targetRows(i)) = rowIndex Then Exit Sub
    Next i

    If Trim$(CommonNzText(ws.Cells(rowIndex, seiriCol).value)) = "" Then Exit Sub
    If IsSanpaiRow(ws, rowIndex) Then Exit Sub
    targetRows.Add rowIndex
End Sub

Private Sub AddEligibleVendorRowsFromRange(ByVal ws As Worksheet, _
                                           ByVal rng As Range, _
                                           ByVal seiriCol As Long, _
                                           ByVal lastRow As Long, _
                                           ByVal targetRows As Collection)
    If rng Is Nothing Then Exit Sub

    Dim startRow As Long
    Dim endRow As Long
    startRow = rng.Row
    endRow = rng.Row + rng.Rows.Count - 1

    Dim rowIndex As Long
    For rowIndex = startRow To endRow
        AddEligibleVendorRowDirect ws, rowIndex, seiriCol, lastRow, targetRows, False
    Next rowIndex
End Sub

Private Function ResolveTargetVendorColumn(ByVal ws As Worksheet, _
                                           ByVal seiriCol As Long, _
                                           ByVal lastRow As Long) As Long
    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        If ActiveCell.Column = WELD_COL_WELDING_VENDOR Then
            ResolveTargetVendorColumn = WELD_COL_WELDING_VENDOR
            Exit Function
        End If
        If ActiveCell.Column = WELD_COL_TRACK_VENDOR Then
            ResolveTargetVendorColumn = WELD_COL_TRACK_VENDOR
            Exit Function
        End If
        If TypeName(Selection) = "Range" Then
            Dim vendorHit As Range
            On Error Resume Next
            Set vendorHit = Application.Intersect(Selection, ws.Columns(WELD_COL_WELDING_VENDOR))
            On Error GoTo 0
            If Not vendorHit Is Nothing Then
                ResolveTargetVendorColumn = WELD_COL_WELDING_VENDOR
                Exit Function
            End If
            On Error Resume Next
            Set vendorHit = Application.Intersect(Selection, ws.Columns(WELD_COL_TRACK_VENDOR))
            On Error GoTo 0
            If Not vendorHit Is Nothing Then
                ResolveTargetVendorColumn = WELD_COL_TRACK_VENDOR
                Exit Function
            End If
        End If
        Exit Function
    End If

    ResolveTargetVendorColumn = COL_VENDOR
End Function

Private Function GetSheetDataLastRow(ByVal ws As Worksheet) As Long
    GetSheetDataLastRow = mod_Construction_LineMapping.GetLastDataRow(ws, SeiriColumn(ws))
End Function

Private Function SeiriColumn(ByVal ws As Worksheet) As Long
    SeiriColumn = mod_Construction_Order_Import.OutputSheetSeiriColumn(ws)
End Function

Private Function TypeColumn(ByVal ws As Worksheet) As Long
    TypeColumn = mod_Construction_Order_Import.OutputSheetCol(ws, COL_TYPE)
End Function

Private Function IsSanpaiRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    IsSanpaiRow = (InStr(1, CommonRemoveAllSpaces(CommonNzText(ws.Cells(rowIndex, TypeColumn(ws)).value)), _
                         SANPAI_KEYWORD, vbTextCompare) > 0)
End Function
