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

Public Function GetSubcontractorList() As Variant
    GetSubcontractorList = GetSubcontractorListByWorkType("")
End Function

Public Function GetSubcontractorListByWorkType(ByVal workTypeKeyword As String) As Variant
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

Private Function VendorBlockMatchesWorkType(ByVal wsInfo As Worksheet, _
                                            ByVal valueCol As Long, _
                                            ByVal workTypeKeyword As String) As Boolean
    Dim workTypeText As String
    workTypeText = CommonRemoveAllSpaces(CommonNormalizeText( _
        CommonNzText(wsInfo.Cells(VENDOR_WORK_TYPE_ROW, valueCol).value)))
    VendorBlockMatchesWorkType = (workTypeText <> "") And _
        (InStr(1, workTypeText, workTypeKeyword, vbTextCompare) > 0)
End Function

Public Sub ApplySubcontractorDropdownsToActiveSheet()
    ApplySubcontractorDropdowns ActiveSheet
End Sub

Public Sub ApplySubcontractorDropdowns(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, SeiriColumn(ws)).End(xlUp).Row
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
        names = GetSubcontractorList()
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
                    .InCellDropdown = True
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

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim seiriCol As Long
    seiriCol = SeiriColumn(ws)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, seiriCol).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "対象データがありません。", vbExclamation
        Exit Sub
    End If

    Dim targetColumn As Long
    targetColumn = ResolveTargetVendorColumn(ws, seiriCol, lastRow)
    If targetColumn = 0 Then
        If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
            MsgBox "溶接会社(A列)または軌道手元会社(B列)を選択した状態で実行してください。", vbExclamation
        Else
            MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
        End If
        Exit Sub
    End If

    Dim targetRows As Collection
    Set targetRows = New Collection

    If TypeName(Selection) = "Range" Then
        Dim usedRange As Range
        Set usedRange = ws.Range(ws.Cells(DATA_START_ROW, seiriCol), ws.Cells(lastRow, seiriCol))

        Dim hit As Range
        On Error Resume Next
        Set hit = Application.Intersect(Selection.EntireRow, usedRange)
        On Error GoTo ErrorHandler

        If Not hit Is Nothing Then
            AddEligibleVendorRowsFromRange ws, hit, seiriCol, targetRows
        End If
    End If

    If targetRows.Count = 0 Then
        If ActiveCell.Row >= DATA_START_ROW And ActiveCell.Row <= lastRow Then
            If IsRowVisibleInRange(ws.Cells(ActiveCell.Row, seiriCol)) Then
                If Trim$(CommonNzText(ws.Cells(ActiveCell.Row, seiriCol).value)) <> "" Then
                    If Not IsSanpaiRow(ws, ActiveCell.Row) Then targetRows.Add ActiveCell.Row
                End If
            End If
        End If
    End If

    If targetRows.Count = 0 Then
        If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
            MsgBox "会社を設定する行を選択してください(C列に整理番号がある行)。", vbExclamation
        Else
            MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
        End If
        Exit Sub
    End If

    Dim names As Variant
    If mod_Construction_Order_Import.IsWeldingOutputSheet(ws) Then
        If targetColumn = WELD_COL_WELDING_VENDOR Then
            names = GetSubcontractorListByWorkType(WELDING_WORK_TYPE_KEYWORD)
        Else
            names = GetSubcontractorListByWorkType(TRACK_WORK_TYPE_KEYWORD)
        End If
    Else
        names = GetSubcontractorList()
    End If

    If Not IsArray(names) Then
        MsgBox "基本情報の施工会社に対応する業者マスタA列の候補が見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim f As frmSubconSelector
    Set f = New frmSubconSelector
    f.SetCompanies names
    f.Show

    Dim confirmed As Boolean, chosen As String
    confirmed = f.confirmed
    chosen = f.SelectedCompany
    Unload f
    If Not confirmed Or chosen = "" Then Exit Sub

    Application.EnableEvents = False

    Dim rIdx As Variant
    For Each rIdx In targetRows
        ws.Cells(CLng(rIdx), targetColumn).value = chosen
    Next rIdx

    mod_Construction_Order_Import.RefreshSubcontractorPriceColumns ws

    ws.Columns(targetColumn).AutoFit

    Application.EnableEvents = prevEvents

    MsgBox targetRows.Count & " 行に「" & chosen & "」を設定しました。", vbInformation
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    MsgBox "施工会社別の単価・金額列を更新できませんでした。" & vbCrLf & _
           Err.Description, vbExclamation
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

Private Function GetVisibleCellsInRange(ByVal rng As Range) As Range
    If rng Is Nothing Then Exit Function
    On Error Resume Next
    Set GetVisibleCellsInRange = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
End Function

Private Function IsRowVisibleInRange(ByVal rng As Range) As Boolean
    IsRowVisibleInRange = Not GetVisibleCellsInRange(rng) Is Nothing
End Function

Private Sub AddEligibleVendorRowsFromRange(ByVal ws As Worksheet, _
                                           ByVal rng As Range, _
                                           ByVal seiriCol As Long, _
                                           ByVal targetRows As Collection)
    Dim visibleRng As Range
    Set visibleRng = GetVisibleCellsInRange(rng)
    If visibleRng Is Nothing Then Exit Sub

    Dim seenRows As Object
    Set seenRows = CreateObject("Scripting.Dictionary")

    Dim c As Range
    For Each c In visibleRng.Cells
        Dim rowIndex As Long
        rowIndex = c.Row
        If Not seenRows.Exists(CStr(rowIndex)) Then
            If Trim$(CommonNzText(ws.Cells(rowIndex, seiriCol).value)) <> "" Then
                If Not IsSanpaiRow(ws, rowIndex) Then
                    seenRows.Add CStr(rowIndex), True
                    targetRows.Add rowIndex
                End If
            End If
        End If
    Next c
End Sub
