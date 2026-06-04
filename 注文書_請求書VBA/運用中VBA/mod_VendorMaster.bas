Option Explicit

'==========================================================================
'  業者マスタ参照／業者選択モジュール
'    改修内容：
'      #18: 業者情報8行目に2行挿入のため、F列セル参照を+2行シフト。
'           BASIC_INFO_VENDOR_NAME_CELL : F9 → F11
'           BASIC_INFO_VENDOR_CLEAR_RANGES : F9:F14,F16:F21 → F11:F16,F18:F23
'           ApplyVendorRowToBasicInfo : F9?F21 → F11?F23
'           FitVendorComboBoxToF9 → FitVendorComboBoxToF11
'      #7 : WriteVendorValidationList の Dictionary→セル単位ループを
'           Variant 2次元配列＋Range 一括代入へ置換。
'      #8 : LoadVendorComboBoxItems のセル単位走査を配列読込ループへ。
'      #9 : NormalizeText / CommonGetBasicInfoWorksheet / 日本語名生成 /
'           ADO 接続生成は mod_Common に集約。重複定義を撤去。
'==========================================================================

Private Const VENDOR_LIST_COL As String = "AD"
Private Const VENDOR_LIST_START_ROW As Long = 2
Private Const BASIC_INFO_VENDOR_NAME_CELL As String = "F11"
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_LABEL_COL As Long = 5
Private Const BASIC_INFO_VENDOR_BLOCK_VALUE_COL As Long = 6
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW As Long = 27
Private Const BASIC_INFO_VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_LABEL_COL_WIDTH As Double = 26.38
Private Const BASIC_INFO_VENDOR_VALUE_COL_WIDTH As Double = 42.5
Private Const BASIC_INFO_VENDOR_SPACER_COL_WIDTH As Double = 0.92
Private Const BASIC_INFO_VENDOR_PERCENT_ROW As Long = 25
Private Const BASIC_INFO_VENDOR_EXTRA_CLEAR_ROW As Long = 27
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const MAX_VENDOR_BLOCK_COUNT As Long = 20
Private Const VENDOR_SOURCE_START_ROW As Long = 2
Private Const VENDOR_SOURCE_END_ROW As Long = 36
Private Const VENDOR_COMBO_NAME As String = "ComboBoxVendor"
Private mVendorPromptTime As Date
Private mVendorTargetAddress As String

Public Sub RefreshVendorListForBasicInfo(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    DeleteVendorComboBox wsInfo

    Dim BranchName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").value))
    If BranchName = "" Then
        ClearVendorList wsInfo
        Exit Sub
    End If

    Dim vendorRows As Collection
    Set vendorRows = LoadVendorRows(BranchName)
    If Not HasVendorRows(vendorRows) Then
        ClearVendorList wsInfo
        Exit Sub
    End If

    WriteVendorValidationList wsInfo, vendorRows
    Exit Sub

ErrorHandler:
    ClearVendorList wsInfo
End Sub

Public Sub FillVendorInfoToBasicInfo(Optional ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    Set targetCell = GetVendorTargetCell(wsInfo, targetCell)
    If targetCell Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim BranchName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").value))

    Dim selectedVendorName As String
    selectedVendorName = CommonNormalizeText(CStr(targetCell.value))
    If selectedVendorName = "" Then
        ClearVendorInfoBlock targetCell
        Exit Sub
    End If
    If BranchName = "" Then Exit Sub

    ApplyVendorSelection BranchName, selectedVendorName, wsInfo, targetCell
    Exit Sub

ErrorHandler:
    MsgBox VendorInfoFillErrorText() & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub ShowAllVendorSelection(Optional ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Set targetCell = GetVendorTargetCell(wsInfo, targetCell)
    If targetCell Is Nothing Then Exit Sub
    mVendorTargetAddress = targetCell.Address(False, False)

    DeleteVendorComboBox wsInfo
    AllVenderSelection.Show vbModal
End Sub

Public Sub ApplyVendorSelection(ByVal BranchName As String, ByVal vendorName As String, Optional ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    Set targetCell = GetVendorTargetCell(wsInfo, targetCell)
    If targetCell Is Nothing Then Exit Sub

    BranchName = CommonNormalizeText(BranchName)
    vendorName = CommonNormalizeText(vendorName)
    If BranchName = "" Or vendorName = "" Then Exit Sub

    Dim vendorRows As Collection
    Set vendorRows = LoadVendorRows(BranchName)
    If Not HasVendorRows(vendorRows) Then Exit Sub

    Dim rowData As Variant
    For Each rowData In vendorRows
        If StrComp(CommonNormalizeText(CStr(rowData(0))), vendorName, vbTextCompare) = 0 Then
            ApplyVendorRowSelection wsInfo, BranchName, rowData, targetCell
            Exit Sub
        End If
    Next rowData
End Sub

Private Sub ApplyVendorRowSelection(ByVal wsInfo As Worksheet, ByVal BranchName As String, ByVal rowData As Variant, ByVal targetCell As Range)
    Dim previousEnableEvents As Boolean
    previousEnableEvents = Application.EnableEvents

    Dim currentBranchName As String
    currentBranchName = CStr(wsInfo.Range("B6").value)

    Dim currentOfficeName As String
    currentOfficeName = CStr(wsInfo.Range("C6").value)

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ApplyVendorRowToBasicInfo targetCell, rowData
    wsInfo.Range("B6").value = currentBranchName
    wsInfo.Range("C6").value = currentOfficeName

ExitHandler:
    Application.EnableEvents = previousEnableEvents
End Sub

Public Function GetAllVendorSelectionData() As Variant
    Dim allRows As Collection
    Set allRows = LoadAllVendorRows()
    If Not HasVendorRows(allRows) Then Exit Function

    Dim result() As Variant
    ReDim result(1 To allRows.Count, 1 To 3)

    Dim i As Long
    Dim rowData As Variant
    For i = 1 To allRows.Count
        rowData = allRows(i)
        result(i, 1) = CStr(rowData(13))
        result(i, 2) = CStr(rowData(0))
        result(i, 3) = CStr(rowData(1))
    Next i

    GetAllVendorSelectionData = result
End Function

Public Sub ScheduleVendorListDropdown(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    If Intersect(ActiveCell, GetVendorNameRange(wsInfo)) Is Nothing Then Exit Sub
    mVendorTargetAddress = ActiveCell.Address(False, False)

    CancelScheduledVendorListDropdown
    mVendorPromptTime = Now + TimeSerial(0, 0, 1)
    On Error Resume Next
    Application.OnTime EarliestTime:=mVendorPromptTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!PromptVendorListDropdown", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        PromptVendorListDropdown wsInfo
    End If
    On Error GoTo 0
End Sub

Public Sub CancelScheduledVendorListDropdown()
    If mVendorPromptTime = 0 Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=mVendorPromptTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!PromptVendorListDropdown", _
                       Schedule:=False
    On Error GoTo 0
    mVendorPromptTime = 0
End Sub

Public Sub HideVendorComboBox(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    DeleteVendorComboBox wsInfo
End Sub

Public Sub PromptVendorListDropdown(Optional ByVal wsInfo As Worksheet)
    mVendorPromptTime = 0
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    If ActiveSheet Is Nothing Then Exit Sub
    If Not ActiveSheet Is wsInfo Then Exit Sub
    If Intersect(ActiveCell, GetVendorNameRange(wsInfo)) Is Nothing Then Exit Sub
    mVendorTargetAddress = ActiveCell.Address(False, False)

    ShowVendorComboBox wsInfo
End Sub

Public Sub CommitVendorComboBoxSelection(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ExitHandler

    Dim targetCell As Range
    Set targetCell = GetVendorTargetCell(wsInfo)
    If targetCell Is Nothing Then GoTo ExitHandler

    Dim selectedVendorName As String
    selectedVendorName = CommonNormalizeText(CStr(wsInfo.OLEObjects(VENDOR_COMBO_NAME).Object.value))
    If selectedVendorName <> "" Then
        targetCell.value = selectedVendorName
        FillVendorInfoToBasicInfo wsInfo, targetCell
    End If

ExitHandler:
    DeleteVendorComboBox wsInfo
    On Error Resume Next
    If Not targetCell Is Nothing Then targetCell.Select
    On Error GoTo 0
End Sub

Private Sub ShowVendorComboBox(ByVal wsInfo As Worksheet)
    On Error GoTo ExitHandler

    Dim targetCell As Range
    Set targetCell = GetVendorTargetCell(wsInfo)
    If targetCell Is Nothing Then Exit Sub

    Dim ole As OLEObject
    Set ole = GetVendorComboBox(wsInfo, targetCell)
    If ole Is Nothing Then Exit Sub

    LoadVendorComboBoxItems wsInfo, ole, targetCell
    FitVendorComboBoxToTarget targetCell, ole

    wsInfo.Activate
    targetCell.Select
    ole.Visible = True
    ole.Activate
    ole.Object.DropDown
    Exit Sub

ExitHandler:
    DeleteVendorComboBox wsInfo
End Sub

Private Function GetVendorComboBox(ByVal wsInfo As Worksheet, ByVal targetCell As Range) As OLEObject
    On Error Resume Next
    Set GetVendorComboBox = wsInfo.OLEObjects(VENDOR_COMBO_NAME)
    On Error GoTo 0
    If Not GetVendorComboBox Is Nothing Then Exit Function

    On Error Resume Next
    Set GetVendorComboBox = wsInfo.OLEObjects.Add(ClassType:="Forms.ComboBox.1", _
                                                  Link:=False, _
                                                  DisplayAsIcon:=False, _
                                                  Left:=targetCell.Left, _
                                                  Top:=targetCell.Top, _
                                                  Width:=targetCell.Width, _
                                                  Height:=targetCell.Height)
    If Not GetVendorComboBox Is Nothing Then
        GetVendorComboBox.Name = VENDOR_COMBO_NAME
        GetVendorComboBox.Visible = False
        FitVendorComboBoxToTarget targetCell, GetVendorComboBox
    End If
    On Error GoTo 0
End Function

Private Sub FitVendorComboBoxToTarget(ByVal targetCell As Range, ByVal ole As OLEObject)
    With ole
        .Left = targetCell.Left
        .Top = targetCell.Top
        .Width = targetCell.Width
        .Height = targetCell.Height
        .Placement = xlMoveAndSize
    End With
End Sub

'--------------------------------------------------------------------------
'  ComboBox 用 業者名アイテム読込
'  改修（#8）：従来は AD 列を 1 セルずつ Cells(...).Value で参照していたが
'              対象範囲を Variant 配列で一括取得してからループする。
'--------------------------------------------------------------------------
Private Sub LoadVendorComboBoxItems(ByVal wsInfo As Worksheet, ByVal ole As OLEObject, ByVal targetCell As Range)
    Dim lastRow As Long
    lastRow = wsInfo.Cells(wsInfo.rows.Count, VENDOR_LIST_COL).End(xlUp).Row
    If lastRow < VENDOR_LIST_START_ROW Then Exit Sub

    Dim arr As Variant
    arr = wsInfo.Range(wsInfo.Cells(VENDOR_LIST_START_ROW, VENDOR_LIST_COL), _
                       wsInfo.Cells(lastRow, VENDOR_LIST_COL)).value

    Dim currentValue As String
    currentValue = CStr(targetCell.value)

    With ole.Object
        .Clear
        Dim rr As Long
        If IsArray(arr) Then
            For rr = 1 To UBound(arr, 1)
                If CommonNormalizeText(CStr(arr(rr, 1))) <> "" Then
                    .AddItem CStr(arr(rr, 1))
                End If
            Next rr
        Else
            ' 範囲が 1 セルだけのとき Variant は配列にならない
            If CommonNormalizeText(CStr(arr)) <> "" Then .AddItem CStr(arr)
        End If
        .LinkedCell = ""
        .ListRows = Application.Max(1, Application.Min(12, .ListCount))
        .MatchRequired = False
        .value = currentValue
    End With
End Sub

'--------------------------------------------------------------------------
'  業者バリデーションリスト書込
'  改修（#7）：Dictionary.Keys を 1 度だけ取得→ 2 次元配列に詰めて
'              Range 一括代入する。
'--------------------------------------------------------------------------
Private Sub WriteVendorValidationList(ByVal wsInfo As Worksheet, ByVal vendorRows As Collection)
    wsInfo.Columns(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).Hidden = False
    wsInfo.Range(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).ClearContents

    Dim vendorList As Object
    Set vendorList = CreateObject("Scripting.Dictionary")
    vendorList.CompareMode = vbTextCompare

    Dim rowData As Variant
    Dim vendorName As String
    For Each rowData In vendorRows
        vendorName = CommonNormalizeText(CStr(rowData(0)))
        If vendorName <> "" Then
            If Not vendorList.Exists(vendorName) Then vendorList.Add vendorName, vendorName
        End If
    Next rowData

    If vendorList.Count > 0 Then
        Dim keysArr As Variant
        keysArr = vendorList.Keys
        Dim outArr() As Variant
        ReDim outArr(1 To vendorList.Count, 1 To 1)
        Dim i As Long
        For i = 0 To vendorList.Count - 1
            outArr(i + 1, 1) = keysArr(i)
        Next i
        wsInfo.Range(VENDOR_LIST_COL & VENDOR_LIST_START_ROW).Resize(vendorList.Count, 1).value = outArr
    End If

    Dim vendorCells As Range
    Set vendorCells = GetVendorNameRange(wsInfo)

    Dim vendorCell As Range
    For Each vendorCell In vendorCells.Cells
        ResetVendorValidation vendorCell, _
                              wsInfo.Range(VENDOR_LIST_COL & VENDOR_LIST_START_ROW).Resize(Application.Max(1, vendorList.Count, 1))
    Next vendorCell

    wsInfo.Columns(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).Hidden = True
End Sub

Private Sub ResetVendorValidation(ByVal targetCell As Range, ByVal listRange As Range)
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Sub ClearVendorList(ByVal wsInfo As Worksheet)
    On Error Resume Next
    wsInfo.Columns(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).Hidden = False
    wsInfo.Range(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).ClearContents
    GetVendorNameRange(wsInfo).Validation.Delete
    ClearAllVendorInfoBlocks wsInfo
    wsInfo.Columns(VENDOR_LIST_COL & ":" & VENDOR_LIST_COL).Hidden = True
    On Error GoTo 0
End Sub

Private Sub DeleteVendorComboBox(ByVal wsInfo As Worksheet)
    On Error Resume Next
    With wsInfo.OLEObjects(VENDOR_COMBO_NAME)
        .Visible = False
        .Delete
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyVendorRowToBasicInfo(ByVal targetCell As Range, ByVal rowData As Variant)
    With targetCell.Worksheet
        .Cells(11, targetCell.Column).value = rowData(0)
        .Cells(12, targetCell.Column).value = rowData(3)
        .Cells(13, targetCell.Column).value = rowData(11)
        .Cells(14, targetCell.Column).value = rowData(2)
        .Cells(15, targetCell.Column).value = rowData(4)
        .Cells(16, targetCell.Column).value = rowData(1)
        .Cells(18, targetCell.Column).value = rowData(5)
        .Cells(19, targetCell.Column).value = rowData(6)
        .Cells(20, targetCell.Column).value = rowData(10)
        .Cells(21, targetCell.Column).value = rowData(7)
        .Cells(22, targetCell.Column).value = rowData(8)
        .Cells(23, targetCell.Column).value = rowData(9)
        .Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column).NumberFormatLocal = "@"
        .Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column).value = "100" & ChrW$(&HFF05)
    End With
End Sub

Public Function GetVendorNameRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim i As Long
    Dim result As Range
    For i = 1 To vendorCount
        If result Is Nothing Then
            Set result = VendorNameCellByIndex(wsInfo, i)
        Else
            Set result = Union(result, VendorNameCellByIndex(wsInfo, i))
        End If
    Next i

    Set GetVendorNameRange = result
End Function

Public Sub SyncVendorBlocksFromCount(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ExitHandler

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, BASIC_INFO_VENDOR_BLOCK_LABEL_COL).value = VendorInfoHeaderText(1)
    ApplyVendorBlockColumnWidths wsInfo, vendorCount

    Dim sourceBlock As Range
    Set sourceBlock = wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, BASIC_INFO_VENDOR_BLOCK_LABEL_COL), _
                                   wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW, BASIC_INFO_VENDOR_BLOCK_VALUE_COL))

    Dim i As Long
    For i = 2 To vendorCount
        Dim destBlock As Range
        Set destBlock = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i)).Resize(sourceBlock.Rows.Count, sourceBlock.Columns.Count)
        sourceBlock.Copy Destination:=destBlock
        destBlock.Cells(1, 1).value = VendorInfoHeaderText(i)
        ClearVendorInfoBlock VendorNameCellByIndex(wsInfo, i)
    Next i

    ClearUnusedVendorBlocks wsInfo, vendorCount + 1
    Application.CutCopyMode = False
    RefreshVendorListForBasicInfo wsInfo
    Exit Sub

ExitHandler:
    Application.CutCopyMode = False
End Sub

Private Function GetVendorTargetCell(ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range) As Range
    If Not targetCell Is Nothing Then
        If Not Intersect(targetCell, GetVendorNameRange(wsInfo)) Is Nothing Then
            Set GetVendorTargetCell = targetCell.Cells(1, 1)
            mVendorTargetAddress = GetVendorTargetCell.Address(False, False)
            Exit Function
        End If
    End If

    If mVendorTargetAddress <> "" Then
        On Error Resume Next
        Set GetVendorTargetCell = wsInfo.Range(mVendorTargetAddress)
        On Error GoTo 0
        If Not GetVendorTargetCell Is Nothing Then
            If Intersect(GetVendorTargetCell, GetVendorNameRange(wsInfo)) Is Nothing Then Set GetVendorTargetCell = Nothing
        End If
    End If

    If GetVendorTargetCell Is Nothing Then
        If Not ActiveCell Is Nothing Then
            If ActiveCell.Worksheet Is wsInfo Then
                If Not Intersect(ActiveCell, GetVendorNameRange(wsInfo)) Is Nothing Then
                    Set GetVendorTargetCell = ActiveCell
                End If
            End If
        End If
    End If

    If GetVendorTargetCell Is Nothing Then Set GetVendorTargetCell = wsInfo.Range(BASIC_INFO_VENDOR_NAME_CELL)
    mVendorTargetAddress = GetVendorTargetCell.Address(False, False)
End Function

Private Function GetVendorBlockCount(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CStr(wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > MAX_VENDOR_BLOCK_COUNT Then countValue = MAX_VENDOR_BLOCK_COUNT
    GetVendorBlockCount = countValue
End Function

Private Function VendorNameCellByIndex(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Range
    Set VendorNameCellByIndex = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, VendorValueColumnByIndex(vendorIndex))
End Function

Private Function VendorLabelColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorLabelColumnByIndex = BASIC_INFO_VENDOR_BLOCK_LABEL_COL + ((vendorIndex - 1) * BASIC_INFO_VENDOR_BLOCK_STEP_COLS)
End Function

Private Function VendorValueColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorValueColumnByIndex = VendorLabelColumnByIndex(vendorIndex) + 1
End Function

Private Function VendorSpacerColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorSpacerColumnByIndex = VendorLabelColumnByIndex(vendorIndex) + 2
End Function

Private Sub ApplyVendorBlockColumnWidths(ByVal wsInfo As Worksheet, ByVal vendorCount As Long)
    Dim i As Long
    For i = 1 To vendorCount
        wsInfo.Columns(VendorLabelColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_LABEL_COL_WIDTH
        wsInfo.Columns(VendorValueColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_VALUE_COL_WIDTH
        wsInfo.Columns(VendorSpacerColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_SPACER_COL_WIDTH
    Next i
End Sub

Private Sub ClearVendorInfoBlock(ByVal targetCell As Range)
    With targetCell.Worksheet
        .Range(.Cells(11, targetCell.Column), .Cells(16, targetCell.Column)).ClearContents
        .Range(.Cells(18, targetCell.Column), .Cells(23, targetCell.Column)).ClearContents
        .Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column).ClearContents
        .Cells(BASIC_INFO_VENDOR_EXTRA_CLEAR_ROW, targetCell.Column).ClearContents
    End With
End Sub

Private Sub ClearAllVendorInfoBlocks(ByVal wsInfo As Worksheet)
    Dim vendorCells As Range
    Set vendorCells = GetVendorNameRange(wsInfo)

    Dim vendorCell As Range
    For Each vendorCell In vendorCells.Cells
        ClearVendorInfoBlock vendorCell
    Next vendorCell
End Sub

Private Sub ClearUnusedVendorBlocks(ByVal wsInfo As Worksheet, ByVal firstUnusedIndex As Long)
    Dim i As Long
    For i = firstUnusedIndex To MAX_VENDOR_BLOCK_COUNT
        Dim headerCell As Range
        Set headerCell = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i))
        If InStr(1, CStr(headerCell.value), VendorInfoHeaderPrefixText(), vbTextCompare) > 0 Then
            Dim clearRange As Range
            Set clearRange = wsInfo.Range(headerCell, headerCell.Offset(BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW - BASIC_INFO_VENDOR_BLOCK_TOP_ROW, 1))
            clearRange.ClearContents
            clearRange.Interior.Color = RGB(6, 17, 29)
            clearRange.Borders.LineStyle = xlNone
        End If
    Next i
End Sub

Private Function VendorInfoHeaderText(ByVal vendorIndex As Long) As String
    VendorInfoHeaderText = "  " & ChrW$(&H25B8) & "  " & VendorInfoHeaderPrefixText() & CStr(vendorIndex)
End Function

Private Function VendorInfoHeaderPrefixText() As String
    VendorInfoHeaderPrefixText = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H60C5) & ChrW$(&H5831) & "-"
End Function

Private Function LoadVendorRows(ByVal BranchName As String) As Collection
    Dim sourceFilePath As String
    sourceFilePath = GetVendorMasterFilePath()
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim sheetName As String
    sheetName = GetAdoWorksheetName(connection, BranchName)
    If sheetName = "" Then GoTo Cleanup

    Set LoadVendorRows = LoadVendorRowsFromAdoSheet(connection, sheetName)

Cleanup:
    CommonCloseAdoConnection connection
End Function

Private Function LoadAllVendorRows() As Collection
    Dim sourceFilePath As String
    sourceFilePath = GetVendorMasterFilePath()
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim rows As Collection
    Set rows = New Collection

    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim sheetName As Variant
    Dim sheetRows As Collection
    Dim rowData As Variant
    For Each sheetName In sheetNames
        Set sheetRows = LoadVendorRowsFromAdoSheet(connection, CStr(sheetName))
        If HasVendorRows(sheetRows) Then
            For Each rowData In sheetRows
                rows.Add rowData
            Next rowData
        End If
    Next sheetName

    If rows.Count > 0 Then Set LoadAllVendorRows = rows

Cleanup:
    CommonCloseAdoConnection connection
End Function

Private Function LoadVendorRowsFromAdoSheet(ByVal connection As Object, ByVal sheetName As String) As Collection
    Dim rows As Collection
    Set rows = New Collection

    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")

    On Error GoTo Cleanup
    recordset.Open "SELECT * FROM " & GetAdoVendorRangeName(sheetName), connection, 0, 1, 1

    Do Until recordset.EOF
        If CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 1))) <> "" Then
            rows.Add BuildVendorRowFromAdoRecord(recordset, sheetName)
        End If
        recordset.MoveNext
    Loop

    If rows.Count > 0 Then Set LoadVendorRowsFromAdoSheet = rows

Cleanup:
    CommonCloseAdoRecordset recordset
End Function

Private Function BuildVendorRowFromAdoRecord(ByVal recordset As Object, ByVal sheetName As String) As Variant
    BuildVendorRowFromAdoRecord = Array(CommonNzText(CommonGetAdoFieldValue(recordset, 1)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 2)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 3)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 4)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 5)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 6)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 7)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 8)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 9)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 10)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 11)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 12)), _
                                        CommonNzText(CommonGetAdoFieldValue(recordset, 0)), _
                                        CommonNormalizeText(sheetName))
End Function

' 業者マスタは指定支店名のシートを検索する。シート名の大文字小文字・
' 全角空白などのゆらぎを吸収するため CommonNormalizeText で比較する。
Private Function GetAdoWorksheetName(ByVal connection As Object, ByVal targetSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim sheetName As Variant
    Dim normalizedTargetName As String
    normalizedTargetName = CommonNormalizeText(targetSheetName)

    For Each sheetName In sheetNames
        If StrComp(CommonNormalizeText(CStr(sheetName)), normalizedTargetName, vbTextCompare) = 0 Then
            GetAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

Private Function GetAdoVendorRangeName(ByVal sheetName As String) As String
    GetAdoVendorRangeName = "[" & Replace$(sheetName, "]", "]]") & "$A" & VENDOR_SOURCE_START_ROW & ":M" & VENDOR_SOURCE_END_ROW & "]"
End Function

Private Function HasVendorRows(ByVal rows As Collection) As Boolean
    On Error Resume Next
    HasVendorRows = Not rows Is Nothing
    If HasVendorRows Then HasVendorRows = (rows.Count > 0)
    On Error GoTo 0
End Function

Private Function GetVendorMasterFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")

    Dim sharedPath As String
    If Len(Trim$(userProfilePath)) > 0 Then
        sharedPath = userProfilePath & Chr$(92) & CommonCompanyNameText() & Chr$(92) & OrderInvoiceDocumentFolderText() & _
                     Chr$(92) & VendorMasterFolderNameText() & Chr$(92) & VendorMasterFileNameText()
        If Len(Dir(sharedPath, vbNormal)) > 0 Then
            GetVendorMasterFilePath = sharedPath
            Exit Function
        End If
    End If

    Dim parentPath As String
    If Len(ThisWorkbook.Path) > 0 Then
        parentPath = fso.GetParentFolderName(ThisWorkbook.Path) & Chr$(92) & VendorMasterFolderNameText() & Chr$(92) & VendorMasterFileNameText()
        If Len(Dir(parentPath, vbNormal)) > 0 Then
            GetVendorMasterFilePath = parentPath
            Exit Function
        End If
    End If

    Dim localPath As String
    localPath = ThisWorkbook.Path & Chr$(92) & VendorMasterFolderNameText() & Chr$(92) & VendorMasterFileNameText()
    If Len(Dir(localPath, vbNormal)) > 0 Then GetVendorMasterFilePath = localPath
End Function

'--------------------------------------------------------------------------
'  本モジュール専用の日本語名
'    線路出張所用_注文書_請求書アクセスサイト - ドキュメント
'    業者マスタ
'    業者マスタ(全社版).xlsx
'    業者情報を入力できませんでした。
'--------------------------------------------------------------------------

Private Function OrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H7528) & _
                 "_" & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & "_" & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & _
                 ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & ChrW$(&H30B9) & _
                 ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & _
                 " - " & ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    OrderInvoiceDocumentFolderText = cached
End Function

Private Function VendorMasterFolderNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF)
    End If
    VendorMasterFolderNameText = cached
End Function

Private Function VendorMasterFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = VendorMasterFolderNameText() & "(" & ChrW$(&H5168) & ChrW$(&H793E) & ChrW$(&H7248) & ").xlsx"
    End If
    VendorMasterFileNameText = cached
End Function

Private Function VendorInfoFillErrorText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H60C5) & ChrW$(&H5831) & ChrW$(&H3092) & _
                 ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3067) & ChrW$(&H304D) & ChrW$(&H307E) & _
                 ChrW$(&H305B) & ChrW$(&H3093) & ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H3002)
    End If
    VendorInfoFillErrorText = cached
End Function


