Option Explicit

Private Const VENDOR_LIST_COL As String = "AD"
Private Const VENDOR_LIST_START_ROW As Long = 2
Private Const BASIC_INFO_VENDOR_NAME_CELL As String = "F11"
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_LABEL_COL As Long = 5
Private Const BASIC_INFO_VENDOR_BLOCK_VALUE_COL As Long = 6
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW As Long = 31
Private Const BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW As Long = 32
Private Const BASIC_INFO_VENDOR_TOTAL_ROW As Long = 33
Private Const BASIC_INFO_VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_LABEL_COL_WIDTH As Double = 26.38
Private Const BASIC_INFO_VENDOR_VALUE_COL_WIDTH As Double = 42.5
Private Const BASIC_INFO_VENDOR_SPACER_COL_WIDTH As Double = 0.92
Private Const BASIC_INFO_VENDOR_PERCENT_ROW As Long = 25
Private Const BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW As Long = 29
Private Const BASIC_INFO_VENDOR_RAIL_PATTERN_ROW As Long = 30
Private Const BASIC_INFO_VENDOR_WELDING_RATIO_ROW As Long = 31
Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C22"
Private Const VENDOR_UNIT_PRICE_JR_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_START As Long = 5
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_END As Long = 6
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW As Long = 1
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_FONT_SIZE As Long = 11
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_NUMBER_FORMAT As String = "0.0%"
Private Const VENDOR_UNIT_PRICE_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_NAME_ROW As Long = 5
Private Const VENDOR_UNIT_PRICE_LABEL_ROW As Long = 6
Private Const VENDOR_UNIT_PRICE_DATA_START_ROW As Long = 7
Private Const VENDOR_UNIT_PRICE_FIRST_DAY_COL As Long = 7
Private Const VENDOR_UNIT_PRICE_REF_UNIT_COL As Long = 5
Private Const VENDOR_UNIT_PRICE_REF_WIDTH_COL As Long = 6
Private Const VENDOR_UNIT_PRICE_WORK_TYPE_COL As Long = 3
Private Const VENDOR_UNIT_PRICE_LAST_ROW_COL As Long = 2
Private Const VENDOR_UNIT_PRICE_INITIAL_FILL_LAST_COL As Long = 10
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_R As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_G As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_B As Long = 128
Private Const VENDOR_UNIT_PRICE_NUMBER_FORMAT As String = "#,##0"

Private Const MAX_VENDOR_BLOCK_COUNT As Long = 20
Private Const VENDOR_SOURCE_START_ROW As Long = 2
Private Const VENDOR_SOURCE_END_ROW As Long = 500
Private Const VENDOR_ROW_NAME_INDEX As Long = 0
Private Const VENDOR_ROW_UNIT_PRICE_NAME_INDEX As Long = 12
Private Const VENDOR_MASTER_ADO_COLUMN_O_NAME As String = "F15"
Private Const VENDOR_MASTER_EXCEL_COLUMN_O As Long = 15
Private Const VENDOR_COMBO_NAME As String = "ComboBoxVendor"
Private mVendorPromptTime As Date
Private mVendorTargetAddress As String
Private mLastVendorBlockCount As Long
Private mVendorRowsCache As Object
Private mVendorNameIndexCache As Object
Private mSyncVendorBlocksInProgress As Boolean

Public Function IsSyncVendorBlocksInProgress() As Boolean
    IsSyncVendorBlocksInProgress = mSyncVendorBlocksInProgress
End Function

Public Sub ResetVendorBlockSyncState()
    mSyncVendorBlocksInProgress = False
End Sub

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

    Dim previousWorkType As String
    previousWorkType = CommonNormalizeText(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).value))

    Dim selectedVendorName As String
    selectedVendorName = CommonNormalizeText(CStr(targetCell.value))
    If selectedVendorName = "" Then
        ClearVendorInfoBlock targetCell
        RefreshVendorUnitPriceForValueColumn wsInfo, targetCell.Column
        NotifyVendorBasicInfoBlockChanged wsInfo, targetCell.Column, previousWorkType
        Exit Sub
    End If
    If BranchName = "" Then
        NotifyVendorBasicInfoBlockChanged wsInfo, targetCell.Column, previousWorkType
        Exit Sub
    End If

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

    Dim nameIndex As Object
    Set nameIndex = GetVendorNameIndex(BranchName, vendorRows)
    If nameIndex.Exists(vendorName) Then
        Dim rowData As Variant
        rowData = vendorRows(CLng(nameIndex(vendorName)))
        ApplyVendorRowSelection wsInfo, BranchName, rowData, targetCell
        Exit Sub
    End If

    RefreshVendorUnitPriceForValueColumn wsInfo, targetCell.Column
    NotifyVendorBasicInfoBlockChanged wsInfo, targetCell.Column
    mod_DebugLog.Log "[VendorMaster] ApplyVendorSelection: no match in vendor master. Branch=[" & _
                      BranchName & "] Vendor=[" & vendorName & "] Col=" & targetCell.Column & _
                      " RowCount=" & vendorRows.Count & " -> row12+ left blank"
End Sub

Public Sub NotifyVendorBasicInfoBlockChanged(ByVal wsInfo As Worksheet, _
                                              ByVal valueColumn As Long, _
                                              Optional ByVal previousWorkType As Variant)
    If wsInfo Is Nothing Then Exit Sub
    If valueColumn <= 0 Then Exit Sub

    On Error GoTo ErrorHandler

    If IsMissing(previousWorkType) Then
        mod_WeldingUnitPrice.UpdateWeldingVendorDisplayNamesForBasicInfo wsInfo, valueColumn
    Else
        Dim prevWorkType As String
        Dim currentWorkType As String
        prevWorkType = CommonNormalizeText(CStr(previousWorkType))
        currentWorkType = CommonNormalizeText(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueColumn).value))

        If StrComp(prevWorkType, currentWorkType, vbTextCompare) = 0 Then
            mod_WeldingUnitPrice.UpdateWeldingVendorDisplayNamesForBasicInfo wsInfo, valueColumn
        ElseIf Len(prevWorkType) = 0 And Len(currentWorkType) > 0 Then
            Dim firstFillCols As Collection
            Set firstFillCols = New Collection
            firstFillCols.Add valueColumn
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfoColumns wsInfo, firstFillCols, valueColumn
        Else
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, valueColumn
        End If
    End If

    Dim vendorIndex As Long
    vendorIndex = GetVendorIndexFromValueColumn(valueColumn)
    If vendorIndex > 0 Then
        mod_BasicInfoGuide.RefreshSingleVendorRowGuidePublic wsInfo, vendorIndex
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals vendorIndex
    End If
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Function GetVendorIndexFromValueColumnPublic(ByVal valueColumn As Long) As Long
    GetVendorIndexFromValueColumnPublic = GetVendorIndexFromValueColumn(valueColumn)
End Function

' F9 件数内の施工会社名セル（F11/I11/L11...）を target から特定する。
' Union レンジとの Intersect が不安定な場合でも列位置で判定する。
Public Function ResolveVendorNameChangeCell(ByVal wsInfo As Worksheet, ByVal target As Range) As Range
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim hitCell As Range
    For Each hitCell In target.Cells
        If hitCell.Row = BASIC_INFO_VENDOR_NAME_ROW Then
            Dim vendorIndex As Long
            vendorIndex = GetVendorIndexFromValueColumn(hitCell.Column)
            If vendorIndex >= 1 And vendorIndex <= vendorCount Then
                Set ResolveVendorNameChangeCell = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, hitCell.Column)
                Exit Function
            End If
        End If
    Next hitCell
End Function

Private Sub ApplyVendorRowSelection(ByVal wsInfo As Worksheet, ByVal BranchName As String, ByVal rowData As Variant, ByVal targetCell As Range)
    Dim previousEnableEvents As Boolean
    previousEnableEvents = Application.EnableEvents

    Dim currentBranchName As String
    currentBranchName = CStr(wsInfo.Range("B6").value)

    Dim currentOfficeName As String
    currentOfficeName = CStr(wsInfo.Range("C6").value)

    Dim previousWorkType As String
    previousWorkType = CommonNormalizeText(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).value))

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ApplyVendorRowToBasicInfo targetCell, rowData
    wsInfo.Range("B6").value = currentBranchName
    wsInfo.Range("C6").value = currentOfficeName

ExitHandler:
    Application.EnableEvents = previousEnableEvents
    Err.Clear

    NotifyVendorBasicInfoBlockChanged wsInfo, targetCell.Column, previousWorkType
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
    If ResolveVendorNameChangeCell(wsInfo, ActiveCell) Is Nothing Then Exit Sub
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
    If ResolveVendorNameChangeCell(wsInfo, ActiveCell) Is Nothing Then Exit Sub
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

    Dim comboDisplayText As String
    comboDisplayText = ole.Object.Text
    With ole.Object
        If .ListCount > 0 Then .ListIndex = 0
    End With

    wsInfo.Activate
    targetCell.Select
    ole.Visible = True
    ole.Activate
    ole.Object.DropDown
    If Len(Trim$(comboDisplayText)) > 0 Then ole.Object.Text = comboDisplayText
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
            If CommonNormalizeText(CStr(arr)) <> "" Then .AddItem CStr(arr)
        End If
        .LinkedCell = ""
        .ListRows = Application.Max(1, Application.Min(12, .ListCount))
        .MatchRequired = False
        .ListIndex = -1
        If Len(Trim$(currentValue)) > 0 Then
            .Text = currentValue
        Else
            .Text = ""
        End If
    End With
End Sub

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
    mod_DebugLog.Log "[VendorMaster] ApplyVendorRow F10=[" & CStr(rowData(14)) & "] vendor=[" & CStr(rowData(0)) & "]"
    With targetCell.Worksheet
        Dim workTypeCell As Range
        Set workTypeCell = VendorWritableValueCell(targetCell.Worksheet, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column)
        ApplyVendorRow10ValueCellFormat .Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column)
        workTypeCell.value = rowData(14)
        VendorWritableValueCell(targetCell.Worksheet, 11, targetCell.Column).value = rowData(0)
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

    RefreshVendorUnitPriceForValueColumn targetCell.Worksheet, targetCell.Column
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

Public Sub InitVendorBlockCountFromSheet(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    ResetVendorBlockSyncState

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim existingBlockCount As Long
    existingBlockCount = CountExistingVendorBlocks(wsInfo)

    ' F9件数に対し業者情報ブロックが不足しているときだけフル同期する。
    ' 毎回 SyncVendorBlocksFromCount を呼ぶと溶接単価展開等で Activate が固まる。
    If vendorCount > existingBlockCount Then
        mLastVendorBlockCount = existingBlockCount
        If mLastVendorBlockCount < 1 Then mLastVendorBlockCount = 1
        SyncVendorBlocksFromCount wsInfo
    Else
        mLastVendorBlockCount = vendorCount
        mod_VendorInfoColors.ApplyVendorInfoRow10Colors wsInfo
        ClearVendorWorkTypeWhenCompanyEmpty wsInfo, vendorCount
        RestoreVendorBlockValueColumnRightBorders wsInfo, vendorCount
    End If
End Sub

Public Sub SyncVendorBlocksFromCount(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If mSyncVendorBlocksInProgress Then Exit Sub

    mSyncVendorBlocksInProgress = True

    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ExitHandler

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim previousCount As Long
    previousCount = mLastVendorBlockCount
    If previousCount <= 0 Then
        previousCount = CountExistingVendorBlocks(wsInfo)
        If previousCount <= 0 Then previousCount = vendorCount
    End If

    wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, BASIC_INFO_VENDOR_BLOCK_LABEL_COL).value = VendorInfoHeaderText(1)
    ApplyVendorBlockColumnWidths wsInfo, vendorCount

    Dim i As Long
    Dim vendorBlocksEnsured As Boolean
    vendorBlocksEnsured = False
    For i = 2 To vendorCount
        If i > previousCount Or VendorBlockNeedsPresentationRestore(wsInfo, i) Then
            EnsureVendorBlockFromTemplate wsInfo, i
            vendorBlocksEnsured = True
        End If
    Next i

    For i = 2 To vendorCount
        wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i)).value = VendorInfoHeaderText(i)
    Next i

    Dim needWeldingRefresh As Boolean
    needWeldingRefresh = False
    Dim oldWeldingBlockCount As Long
    Dim oldRailBlockCount As Long
    oldWeldingBlockCount = 0
    oldRailBlockCount = 0
    If vendorCount < previousCount Then
        mod_WeldingUnitPrice.GetVendorBlockLayoutCountsForLimit wsInfo, previousCount, _
            oldWeldingBlockCount, oldRailBlockCount
        needWeldingRefresh = mod_WeldingUnitPrice.RemovedVendorIndicesRequireWeldingRefresh( _
            wsInfo, vendorCount + 1, previousCount)
    End If

    ClearUnusedVendorBlocks wsInfo, vendorCount + 1

    ClearVendorWorkTypeWhenCompanyEmpty wsInfo, vendorCount

    Dim formatIndex As Long
    For formatIndex = 1 To vendorCount
        ApplyVendorRow10ValueCellFormat wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorValueColumnByIndex(formatIndex))
    Next formatIndex

    Application.CutCopyMode = False
    RefreshVendorListForBasicInfo wsInfo
    SyncVendorUnitPriceBlocksAfterCountChange wsInfo, vendorCount, previousCount, True

    If vendorCount > previousCount Then
        ' 追加ブロックは空のため溶接単価の再展開は不要
    ElseIf vendorCount < previousCount Then
        mod_WeldingUnitPrice.RefreshWeldingAfterVendorCountDecrease wsInfo, vendorCount, _
            oldWeldingBlockCount, oldRailBlockCount, True
    ElseIf vendorBlocksEnsured Then
        mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, 0, True
    End If
    mLastVendorBlockCount = vendorCount

    If vendorBlocksEnsured Or vendorCount > previousCount Then
        Dim restoreIndex As Long
        For restoreIndex = 1 To vendorCount
            mod_BasicInfoGuide.RefreshSingleVendorRowGuidePublic wsInfo, restoreIndex
        Next restoreIndex
    ElseIf vendorCount <> previousCount Then
        mod_BasicInfoGuide.RefreshVendorGuidesForBasicInfo wsInfo
    End If

    RestoreVendorBlockValueColumnRightBorders wsInfo, vendorCount

    If vendorBlocksEnsured Or vendorCount <> previousCount Or needWeldingRefresh Then
        On Error Resume Next
        Application.Calculate
        On Error GoTo ExitHandler
    End If

    Exit Sub

ExitHandler:
    If Err.Number <> 0 Then
        mod_DebugLog.Log "[VendorMaster] SyncVendorBlocksFromCount Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
    Application.CutCopyMode = False
    mSyncVendorBlocksInProgress = False
End Sub

Public Function GetVendorUnitPriceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim result As Range
    Dim i As Long
    For i = 1 To vendorCount
        Dim valueCol As Long
        valueCol = VendorValueColumnByIndex(i)

        Dim blockRange As Range
        Set blockRange = Union(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_RAIL_PATTERN_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, valueCol))

        If result Is Nothing Then
            Set result = blockRange
        Else
            Set result = Union(result, blockRange)
        End If
    Next i

    Dim yearBillingRange As Range
    Set yearBillingRange = wsInfo.Range(BASIC_INFO_YEAR_CELL & "," & BASIC_INFO_BILLING_COUNT_CELL)
    If result Is Nothing Then
        Set result = yearBillingRange
    Else
        Set result = Union(result, yearBillingRange)
    End If
    Set GetVendorUnitPriceMonitorRange = result
End Function

Public Sub ApplyImportedUnitPriceJrHeadersForBasicInfo(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            ApplyVendorUnitPriceJrHeader wsUnitPrice, wsInfo
        End If
    Next wsUnitPrice
End Sub

Public Sub HandleVendorUnitPriceMonitorChange(ByVal wsInfo As Worksheet, ByVal changedRange As Range)
    If wsInfo Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If mSyncVendorBlocksInProgress Then Exit Sub
    If Intersect(changedRange, GetVendorUnitPriceMonitorRange(wsInfo)) Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ' 工事種別(10行目)・業者名(11行目)が変わった列はヘッダー/結合/罫線/数式を含めた全展開が必要
    Dim structuralCols As Collection
    Set structuralCols = CollectMonitorChangedValueColumns(wsInfo, changedRange, False, False)

    ' 外注比率(29行目)のみが変わった列は、単価数式が比率セルを直接参照しているため
    ' 表示欄(ラベル・値)だけ更新すれば良く、全展開は不要(速度改善)
    Dim ratioOnlyCols As Collection
    Set ratioOnlyCols = CollectOutsourceRatioOnlyChangedValueColumns(wsInfo, changedRange, structuralCols)

    Dim weldingCols As Collection
    Set weldingCols = CollectMonitorChangedValueColumns(wsInfo, changedRange, False, True)

    Dim preferredRatioColumn As Long
    preferredRatioColumn = GetPreferredWeldingRatioColumnFromChange(wsInfo, changedRange)

    Dim col As Variant
    For Each col In structuralCols
        mod_DebugLog.Log "[VendorMaster] HandleVendorUnitPriceMonitorChange: full refresh col=" & CLng(col)
        RefreshVendorUnitPriceForValueColumn wsInfo, CLng(col)
    Next col

    For Each col In ratioOnlyCols
        mod_DebugLog.Log "[VendorMaster] HandleVendorUnitPriceMonitorChange: ratio-only refresh col=" & CLng(col)
        RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn wsInfo, CLng(col)
    Next col

    If weldingCols.Count > 0 Then
        ' 10行目(工事種別)変更は溶接/軌道列の配置が変わるため全展開する
        If ChangedRangeIncludesVendorWorkTypeRow(wsInfo, changedRange) Then
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, preferredRatioColumn
        Else
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfoColumns _
                wsInfo, weldingCols, preferredRatioColumn
        End If
    End If

ExitHandler:
    Application.EnableEvents = prevEvents
End Sub

Private Function CollectMonitorChangedValueColumns(ByVal wsInfo As Worksheet, _
                                                   ByVal changedRange As Range, _
                                                   ByVal includeOutsourceRatioRow As Boolean, _
                                                   ByVal includeWeldingRatioRow As Boolean) As Collection
    Dim result As Collection
    Dim vendorCount As Long
    Dim i As Long
    Set result = New Collection

    If wsInfo Is Nothing Then
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If
    If changedRange Is Nothing Then
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If

    If Not Intersect(changedRange, wsInfo.Range(BASIC_INFO_YEAR_CELL & "," & BASIC_INFO_BILLING_COUNT_CELL)) Is Nothing Then
        vendorCount = GetVendorBlockCount(wsInfo)
        For i = 1 To vendorCount
            AddUniqueLongToCollection result, VendorValueColumnByIndex(i)
        Next i
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If

    vendorCount = GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        Dim valueCol As Long
        Dim monitorCell As Range
        valueCol = VendorValueColumnByIndex(i)

        Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol)
        If Not Intersect(changedRange, monitorCell) Is Nothing Then
            AddUniqueLongToCollection result, valueCol
            GoTo ContinueNextVendorColumn
        End If

        Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol)
        If Not Intersect(changedRange, monitorCell) Is Nothing Then
            AddUniqueLongToCollection result, valueCol
            GoTo ContinueNextVendorColumn
        End If

        If includeOutsourceRatioRow Then
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
                GoTo ContinueNextVendorColumn
            End If
        End If

        If includeWeldingRatioRow Then
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
                GoTo ContinueNextVendorColumn
            End If

            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_RAIL_PATTERN_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
            End If
        End If

ContinueNextVendorColumn:
    Next i

    Set CollectMonitorChangedValueColumns = result
End Function

Private Function ChangedRangeIncludesVendorWorkTypeRow(ByVal wsInfo As Worksheet, _
                                                       ByVal changedRange As Range) As Boolean
    Dim vendorCount As Long
    Dim i As Long

    If wsInfo Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function

    vendorCount = GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        If Not Intersect(changedRange, wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorValueColumnByIndex(i))) Is Nothing Then
            ChangedRangeIncludesVendorWorkTypeRow = True
            Exit Function
        End If
    Next i
End Function

Private Sub AddUniqueLongToCollection(ByVal target As Collection, ByVal value As Long)
    Dim existing As Variant
    For Each existing In target
        If CLng(existing) = value Then Exit Sub
    Next existing
    target.Add value
End Sub

Private Function IsLongInCollection(ByVal target As Collection, ByVal value As Long) As Boolean
    If target Is Nothing Then Exit Function

    Dim existing As Variant
    For Each existing In target
        If CLng(existing) = value Then
            IsLongInCollection = True
            Exit Function
        End If
    Next existing
End Function

' 外注比率(29行目)のみが変わった業者列を集める。工事種別/業者名が変わった列(excludeColumns)は
' 全展開側(RefreshVendorUnitPriceForValueColumn)が比率表示も含めて処理するため、ここでは除外する。
Private Function CollectOutsourceRatioOnlyChangedValueColumns(ByVal wsInfo As Worksheet, _
                                                              ByVal changedRange As Range, _
                                                              ByVal excludeColumns As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    If wsInfo Is Nothing Then
        Set CollectOutsourceRatioOnlyChangedValueColumns = result
        Exit Function
    End If
    If changedRange Is Nothing Then
        Set CollectOutsourceRatioOnlyChangedValueColumns = result
        Exit Function
    End If

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        Dim valueCol As Long
        valueCol = VendorValueColumnByIndex(i)

        If Not IsLongInCollection(excludeColumns, valueCol) Then
            Dim monitorCell As Range
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
            End If
        End If
    Next i

    Set CollectOutsourceRatioOnlyChangedValueColumns = result
End Function

Private Function GetPreferredWeldingRatioColumnFromChange(ByVal wsInfo As Worksheet, _
                                                          ByVal changedRange As Range) As Long
    Dim vendorCount As Long
    Dim i As Long
    Dim ratioCell As Range

    GetPreferredWeldingRatioColumnFromChange = 0
    If wsInfo Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function

    vendorCount = GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        Set ratioCell = wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, VendorValueColumnByIndex(i))
        If Not Intersect(changedRange, ratioCell) Is Nothing Then
            GetPreferredWeldingRatioColumnFromChange = ratioCell.Column
            Exit Function
        End If
    Next i
End Function

Public Sub ApplyConstructionUnitPriceImportedRowDecorations(ByVal wsUnitPrice As Worksheet, _
                                                            ByVal firstRow As Long, _
                                                            ByVal lastRow As Long)
    If wsUnitPrice Is Nothing Then Exit Sub
    If lastRow < firstRow Then Exit Sub
    If firstRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(wsUnitPrice.Parent)
    If wsInfo Is Nothing Then Exit Sub

    Dim bColRange As Range
    Set bColRange = wsUnitPrice.Range( _
        wsUnitPrice.Cells(firstRow, VENDOR_UNIT_PRICE_LAST_ROW_COL), _
        wsUnitPrice.Cells(lastRow, VENDOR_UNIT_PRICE_LAST_ROW_COL))
    ApplyVendorUnitPriceNewRowFill wsUnitPrice, wsInfo, bColRange
    ApplyVendorUnitPriceBaseRowBorders wsUnitPrice, wsInfo, bColRange
    ApplyVendorUnitPriceSourceRowsForRange wsUnitPrice, wsInfo, firstRow, lastRow
End Sub

Public Sub HandleConstructionUnitPriceSheetChange(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal changedRange As Range)
    If wsUnitPrice Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(wsUnitPrice.Parent)
    If wsInfo Is Nothing Then Exit Sub

    EnsureApplicationCalculationAutomatic

    Dim changedB As Range
    Dim changedE As Range
    Dim changedF As Range
    Set changedB = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_LAST_ROW_COL))
    Set changedE = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_UNIT_COL))
    Set changedF = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_WIDTH_COL))

    If Not changedB Is Nothing Then
        ApplyVendorUnitPriceNewRowFill wsUnitPrice, wsInfo, changedB
        ApplyVendorUnitPriceBaseRowBorders wsUnitPrice, wsInfo, changedB
    End If
    If Not changedE Is Nothing Then
        EnsureVendorUnitPriceNewRowFillForSourceRows wsUnitPrice, wsInfo, changedE
        HandleVendorUnitPriceSourceChanges wsUnitPrice, wsInfo, changedE, True
    End If
    If Not changedF Is Nothing Then
        EnsureVendorUnitPriceNewRowFillForSourceRows wsUnitPrice, wsInfo, changedF
        HandleVendorUnitPriceSourceChanges wsUnitPrice, wsInfo, changedF, False
    End If
    If Not changedB Is Nothing Then
        ResetClearedVendorUnitPriceRows wsUnitPrice, wsInfo, changedB
    End If

    mod_Construction_Order_Import.RefreshConstructionReferencePricesForUnitPriceChange _
        wsUnitPrice, changedRange

    On Error Resume Next
    wsUnitPrice.Calculate
    On Error GoTo 0
End Sub

Public Sub RefreshAllVendorUnitPricesForBasicInfo(Optional ByVal wsInfo As Worksheet, Optional ByVal deferCalculation As Boolean = False)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    EnsureApplicationCalculationAutomatic

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)
    SyncVendorUnitPriceBlocksAfterCountChange wsInfo, vendorCount, 0, deferCalculation
End Sub

' F9(施工会社数)変更時: 増減した列だけ工事単価シートへ反映し、全社・全シート再展開を避ける。
Private Sub SyncVendorUnitPriceBlocksAfterCountChange(ByVal wsInfo As Worksheet, _
                                                      ByVal vendorCount As Long, _
                                                      ByVal previousCount As Long, _
                                                      ByVal deferCalculation As Boolean)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Dim vendorUnitPriceNameMap As Object
    Dim wsUnitPrice As Worksheet
    Dim i As Long
    Dim blockIndex As Long
    Dim valueColumn As Long
    Dim dayCol As Long
    Dim nightCol As Long

    Set targetBook = wsInfo.Parent

    If previousCount <= 0 Or vendorCount = previousCount Then
        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)

        For Each wsUnitPrice In targetBook.worksheets
            If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
                RefreshVendorUnitPriceBlocksOnSheet wsUnitPrice, wsInfo, vendorCount, vendorUnitPriceNameMap
            End If
        Next wsUnitPrice

        If Not deferCalculation Then
            On Error Resume Next
            Application.Calculate
            On Error GoTo 0
        End If
        Exit Sub
    End If

    If vendorCount > previousCount Then
        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)
        For Each wsUnitPrice In targetBook.worksheets
            If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
                For i = previousCount + 1 To vendorCount
                    valueColumn = VendorValueColumnByIndex(i)
                    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
                    nightCol = dayCol + 1
                    If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                        ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
                    Else
                        ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
                    End If
                Next i
            End If
        Next wsUnitPrice
        Exit Sub
    End If

    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            Dim clearLastIndex As Long
            clearLastIndex = previousCount
            If clearLastIndex > MAX_VENDOR_BLOCK_COUNT Then clearLastIndex = MAX_VENDOR_BLOCK_COUNT
            For blockIndex = vendorCount + 1 To clearLastIndex
                dayCol = VendorUnitPriceDayColumnByValueColumn(VendorValueColumnByIndex(blockIndex))
                nightCol = dayCol + 1
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            Next blockIndex
        End If
    Next wsUnitPrice
End Sub

Private Sub RefreshVendorUnitPriceForValueColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    Dim vendorUnitPriceNameMap As Object
    Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
            ElseIf IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
                   HasVendorName(wsInfo, valueColumn) Then
                ' 11行目のみ入力済み。29行目入力待ちの間は既存列を消さない
            Else
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            End If
        End If
    Next wsUnitPrice
End Sub

' 外注比率(29行目)のみが変わった場合の軽量更新。
' 単価行の数式(day/night列)は基本情報シートの比率セルを直接参照(絶対参照)しているため、
' 比率の値が変わっても数式自体を書き直す必要はなく、Excelの再計算で自動的に反映される。
' そのため、ブロックが構築済みの単価シートに対しては比率表示欄(ラベル・値の2セル)だけを
' 更新し、見出し結合・罫線・フォント・数式の全再構築(ApplyVendorUnitPriceBlockToSheet)は行わない。
' 業者名は入力済みだが比率が今回初めて入力された(ブロック未構築)単価シートに対しては、
' 従来通り全展開(ApplyVendorUnitPriceBlockToSheet)にフォールバックする。
Private Sub RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    Dim vendorUnitPriceNameMap As Object
    Dim nameMapLoaded As Boolean
    nameMapLoaded = False

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                If IsVendorUnitPriceBlockAlreadyBuilt(wsUnitPrice, dayCol, nightCol) Then
                    ApplyVendorUnitPriceOutsourceRatioRow wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol
                Else
                    If Not nameMapLoaded Then
                        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)
                        nameMapLoaded = True
                    End If
                    ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
                End If
            ElseIf IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
                   HasVendorName(wsInfo, valueColumn) Then
                ' 11行目のみ入力済み。29行目入力待ちの間は既存列を消さない
            Else
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            End If
        End If
    Next wsUnitPrice
End Sub

' 単価シート側に指定列の業者ブロック(見出し結合セル)が既に構築済みかどうかを判定する。
' 未構築(初回)の場合は全展開が必要なため、呼び出し元でフォールバックする。
Private Function IsVendorUnitPriceBlockAlreadyBuilt(ByVal wsUnitPrice As Worksheet, _
                                                    ByVal dayCol As Long, _
                                                    ByVal nightCol As Long) As Boolean
    On Error GoTo ExitHandler

    Dim headerCell As Range
    Set headerCell = wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol)

    IsVendorUnitPriceBlockAlreadyBuilt = _
        headerCell.MergeCells And _
        (Len(Trim$(CStr(headerCell.MergeArea.Cells(1, 1).value))) > 0)
    Exit Function

ExitHandler:
    IsVendorUnitPriceBlockAlreadyBuilt = False
End Function

Private Sub RefreshVendorUnitPriceBlocksOnSheet(ByVal wsUnitPrice As Worksheet, _
                                                 ByVal wsInfo As Worksheet, _
                                                 ByVal vendorCount As Long, _
                                                 ByVal vendorUnitPriceNameMap As Object)
    Dim i As Long
    For i = 1 To vendorCount
        Dim valueColumn As Long
        Dim dayCol As Long
        Dim nightCol As Long
        valueColumn = VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1

        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
        Else
            ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
        End If
    Next i

    For i = vendorCount + 1 To MAX_VENDOR_BLOCK_COUNT
        valueColumn = VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1
        ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
    Next i
End Sub

Private Function ShouldApplyVendorUnitPriceBlock(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    ShouldApplyVendorUnitPriceBlock = _
        IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
        HasVendorName(wsInfo, valueColumn) And _
        HasVendorOutsourceRatio(wsInfo, valueColumn)
End Function

Private Function HasVendorName(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    HasVendorName = (Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))) > 0)
End Function

Private Sub ApplyVendorUnitPriceBlockToSheet(ByVal wsUnitPrice As Worksheet, _
                                              ByVal wsInfo As Worksheet, _
                                              ByVal valueColumn As Long, _
                                              ByVal vendorUnitPriceNameMap As Object)
    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    ApplyVendorUnitPriceColumnWidths wsUnitPrice, dayCol, nightCol
    ApplyVendorUnitPriceOutsourceRatioRow wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol
    ApplyVendorUnitPriceMergedHeader wsUnitPrice, dayCol, nightCol, BuildVendorUnitPriceHeaderText(wsInfo)
    ApplyVendorUnitPriceMergedVendorName wsUnitPrice, dayCol, nightCol, _
        ResolveVendorUnitPriceName(vendorUnitPriceNameMap, _
                                   CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))

    With wsUnitPrice
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).value = VendorUnitPriceDayLabelText()
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).value = VendorUnitPriceNightLabelText()
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).HorizontalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).VerticalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).HorizontalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).VerticalAlignment = xlCenter
    End With

    ApplyVendorUnitPriceDataRows wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol
    ApplyVendorUnitPriceFont wsUnitPrice, dayCol, nightCol
    ApplyVendorUnitPriceBorders wsUnitPrice, dayCol, nightCol
End Sub

Public Function BuildVendorUnitPriceNameMap(ByVal wsInfo As Worksheet) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim BranchName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").value))
    If BranchName = "" Then
        Set BuildVendorUnitPriceNameMap = result
        Exit Function
    End If

    Dim vendorRows As Collection
    Set vendorRows = LoadVendorRows(BranchName)
    If Not HasVendorRows(vendorRows) Then
        Set BuildVendorUnitPriceNameMap = result
        Exit Function
    End If

    Dim rowData As Variant
    For Each rowData In vendorRows
        Dim vendorNameKey As String
        Dim unitPriceVendorName As String
        vendorNameKey = CommonNormalizeText(CStr(rowData(VENDOR_ROW_NAME_INDEX)))
        unitPriceVendorName = CommonNzText(rowData(VENDOR_ROW_UNIT_PRICE_NAME_INDEX))

        If vendorNameKey <> "" And unitPriceVendorName <> "" Then
            If Not result.Exists(vendorNameKey) Then
                result.Add vendorNameKey, unitPriceVendorName
            End If
        End If
    Next rowData

    Set BuildVendorUnitPriceNameMap = result
End Function

Private Function ResolveVendorUnitPriceName(ByVal vendorUnitPriceNameMap As Object, _
                                            ByVal basicInfoVendorName As String) As String
    Dim vendorNameKey As String
    vendorNameKey = CommonNormalizeText(basicInfoVendorName)

    If Not vendorUnitPriceNameMap Is Nothing Then
        If vendorUnitPriceNameMap.Exists(vendorNameKey) Then
            ResolveVendorUnitPriceName = CStr(vendorUnitPriceNameMap(vendorNameKey))
            Exit Function
        End If
    End If

    ResolveVendorUnitPriceName = basicInfoVendorName
    If vendorNameKey <> "" Then
        mod_DebugLog.Log "[VendorMaster] Unit price vendor name not found: " & basicInfoVendorName
    End If
End Function

Private Sub ClearVendorUnitPriceBlockOnSheet(ByVal wsUnitPrice As Worksheet, _
                                             ByVal dayCol As Long, _
                                             ByVal nightCol As Long)
    If wsUnitPrice Is Nothing Then Exit Sub

    ClearVendorUnitPriceOutsourceRatioRow wsUnitPrice, dayCol, nightCol

    Dim lastRow As Long
    lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then lastRow = VENDOR_UNIT_PRICE_DATA_START_ROW + 200

    Dim clearRange As Range
    Set clearRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                       wsUnitPrice.Cells(lastRow, nightCol))

    SafeUnmergeRange clearRange
    clearRange.ClearContents
    clearRange.Interior.ColorIndex = xlColorIndexNone
    clearRange.ShrinkToFit = False
    clearRange.Borders.LineStyle = xlNone
End Sub

Private Sub ApplyVendorUnitPriceJrHeader(ByVal wsUnitPrice As Worksheet, ByVal wsInfo As Worksheet)
    Dim headerRange As Range
    Set headerRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_JR_HEADER_ROW, VENDOR_UNIT_PRICE_JR_HEADER_COL_START), _
                                        wsUnitPrice.Cells(VENDOR_UNIT_PRICE_JR_HEADER_ROW, VENDOR_UNIT_PRICE_JR_HEADER_COL_END))

    SafeUnmergeRange headerRange
    headerRange.Merge
    With headerRange
        .value = BuildVendorUnitPriceJrHeaderText(wsInfo)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
        .Font.Name = VendorUnitPriceFontNameText()
        On Error Resume Next
        .Font.NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With

    With headerRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Private Function BuildVendorUnitPriceJrHeaderText(ByVal wsInfo As Worksheet) As String
    BuildVendorUnitPriceJrHeaderText = Trim$(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).value)) & _
                                       VendorJrUnitPriceLabelText() & _
                                       Trim$(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL).value))
End Function

Private Sub ApplyVendorUnitPriceColumnWidths(ByVal wsUnitPrice As Worksheet, _
                                               ByVal dayCol As Long, _
                                               ByVal nightCol As Long)
    wsUnitPrice.Columns(dayCol).ColumnWidth = wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_UNIT_COL).ColumnWidth
    wsUnitPrice.Columns(nightCol).ColumnWidth = wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_WIDTH_COL).ColumnWidth
End Sub

Private Sub ApplyVendorUnitPriceOutsourceRatioRow(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal wsInfo As Worksheet, _
                                                  ByVal valueColumn As Long, _
                                                  ByVal dayCol As Long, _
                                                  ByVal nightCol As Long)
    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol)
        .Formula = ""
        .value = VendorUnitPriceOutsourceRatioLabelText()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol)
        .Formula = ""
        .value = GetVendorOutsourceRatioPercentValue(wsInfo, valueColumn)
        .NumberFormat = VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_NUMBER_FORMAT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ApplyVendorUnitPriceOutsourceRatioRowFont wsUnitPrice, dayCol, nightCol
End Sub

Private Sub ApplyVendorUnitPriceOutsourceRatioRowFont(ByVal wsUnitPrice As Worksheet, _
                                                      ByVal dayCol As Long, _
                                                      ByVal nightCol As Long)
    Dim fontRange As Range
    Set fontRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol), _
                                      wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol))

    With fontRange.Font
        .Name = VendorUnitPriceFontNameText()
        .Size = VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_FONT_SIZE
        On Error Resume Next
        .NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Private Sub ClearVendorUnitPriceOutsourceRatioRow(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal dayCol As Long, _
                                                  ByVal nightCol As Long)
    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol)
        .ClearContents
        .HorizontalAlignment = xlGeneral
    End With

    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol)
        .ClearContents
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
    End With
End Sub

Private Function GetVendorOutsourceRatioPercentValue(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Variant
    Dim ratioValue As Variant
    ratioValue = GetVendorOutsourceRatioNumericValue(wsInfo, valueColumn)
    If Not IsNumeric(ratioValue) Then Exit Function

    Dim normalizedValue As Double
    normalizedValue = CDbl(ratioValue)
    If normalizedValue > 1# And normalizedValue <= 100# Then normalizedValue = normalizedValue / 100#
    GetVendorOutsourceRatioPercentValue = normalizedValue
End Function

Private Function GetVendorOutsourceRatioNumericValue(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Variant
    Dim ratioCell As Range
    Set ratioCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn)

    Dim cellValue As Variant
    cellValue = ratioCell.Value2
    If IsNumeric(cellValue) Then
        GetVendorOutsourceRatioNumericValue = CDbl(cellValue)
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(ratioCell.value))
    If Len(textValue) = 0 Then Exit Function

    If Right$(textValue, 1) = ChrW$(&HFF05) Then
        textValue = Trim$(Left$(textValue, Len(textValue) - 1))
        If IsNumeric(textValue) Then GetVendorOutsourceRatioNumericValue = CDbl(textValue) / 100#
        Exit Function
    End If

    If IsNumeric(textValue) Then GetVendorOutsourceRatioNumericValue = CDbl(textValue)
End Function

Private Sub ApplyVendorUnitPriceMergedHeader(ByVal wsUnitPrice As Worksheet, _
                                               ByVal dayCol As Long, _
                                               ByVal nightCol As Long, _
                                               ByVal headerText As String)
    Dim headerRange As Range
    Set headerRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                        wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, nightCol))

    SafeUnmergeRange headerRange
    headerRange.Merge
    With headerRange
        .value = headerText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = False
        .WrapText = False
    End With
End Sub

Private Sub ApplyVendorUnitPriceMergedVendorName(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal dayCol As Long, _
                                                   ByVal nightCol As Long, _
                                                   ByVal vendorName As String)
    Dim nameRange As Range
    Set nameRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_NAME_ROW, dayCol), _
                                      wsUnitPrice.Cells(VENDOR_UNIT_PRICE_NAME_ROW, nightCol))

    SafeUnmergeRange nameRange
    nameRange.Merge
    With nameRange
        .value = vendorName
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
    End With
End Sub

Private Sub ApplyVendorUnitPriceDataRows(ByVal wsUnitPrice As Worksheet, _
                                           ByVal wsInfo As Worksheet, _
                                           ByVal valueColumn As Long, _
                                           ByVal dayCol As Long, _
                                           ByVal nightCol As Long)
    Dim lastRow As Long
    lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    If lastRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    Dim ratioAddress As String
    ratioAddress = GetVendorOutsourceRatioAddress(wsInfo, valueColumn)

    Dim dayFormulaR1C1 As String
    Dim nightFormulaR1C1 As String
    dayFormulaR1C1 = BuildVendorUnitPriceFormulaR1C1(True, dayCol, ratioAddress)
    nightFormulaR1C1 = BuildVendorUnitPriceFormulaR1C1(False, nightCol, ratioAddress)

    Dim wasteKeyword As String
    wasteKeyword = VendorWasteDisposalKeywordText()

    ApplyVendorUnitPriceDataColumn wsUnitPrice, wsInfo, valueColumn, dayCol, VENDOR_UNIT_PRICE_REF_UNIT_COL, _
        dayFormulaR1C1, wasteKeyword, lastRow, True
    ApplyVendorUnitPriceDataColumn wsUnitPrice, wsInfo, valueColumn, nightCol, VENDOR_UNIT_PRICE_REF_WIDTH_COL, _
        nightFormulaR1C1, wasteKeyword, lastRow, False
End Sub

Private Sub ApplyVendorUnitPriceDataColumn(ByVal wsUnitPrice As Worksheet, _
                                           ByVal wsInfo As Worksheet, _
                                           ByVal valueColumn As Long, _
                                           ByVal targetCol As Long, _
                                           ByVal sourceCol As Long, _
                                           ByVal formulaR1C1 As String, _
                                           ByVal wasteKeyword As String, _
                                           ByVal lastRow As Long, _
                                           ByVal isDayColumn As Boolean)
    Dim rowIndex As Long
    Dim formulaSegStart As Long
    formulaSegStart = 0

    For rowIndex = VENDOR_UNIT_PRICE_DATA_START_ROW To lastRow + 1
        Dim needsFormula As Boolean
        needsFormula = False
        If rowIndex <= lastRow Then
            needsFormula = VendorUnitPriceRowNeedsFormulaForSource( _
                wsUnitPrice, rowIndex, sourceCol, wasteKeyword)
        End If

        If needsFormula Then
            If formulaSegStart = 0 Then formulaSegStart = rowIndex
        Else
            If formulaSegStart > 0 Then
                ApplyVendorUnitPriceFormulaSegment wsUnitPrice, wsInfo, valueColumn, formulaSegStart, rowIndex - 1, _
                    targetCol, formulaR1C1, isDayColumn
                formulaSegStart = 0
            End If
            If rowIndex <= lastRow Then
                ApplyVendorUnitPriceGreyFill wsUnitPrice.Cells(rowIndex, targetCol)
            End If
        End If
    Next rowIndex
End Sub

Private Function VendorUnitPriceRowNeedsFormulaForSource(ByVal wsUnitPrice As Worksheet, _
                                                         ByVal rowIndex As Long, _
                                                         ByVal sourceCol As Long, _
                                                         ByVal wasteKeyword As String) As Boolean
    If IsVendorUnitPriceSourceBlank(wsUnitPrice, rowIndex, sourceCol) Then Exit Function

    Dim workTypeName As String
    workTypeName = CommonNormalizeText(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_WORK_TYPE_COL).value))
    If workTypeName <> "" Then
        If InStr(1, workTypeName, wasteKeyword, vbTextCompare) > 0 Then Exit Function
    End If

    VendorUnitPriceRowNeedsFormulaForSource = True
End Function

Private Sub ApplyVendorUnitPriceFormulaSegment(ByVal wsUnitPrice As Worksheet, _
                                               ByVal wsInfo As Worksheet, _
                                               ByVal valueColumn As Long, _
                                               ByVal segStart As Long, _
                                               ByVal segEnd As Long, _
                                               ByVal targetCol As Long, _
                                               ByVal formulaR1C1 As String, _
                                               ByVal isDayColumn As Boolean)
    If segStart > segEnd Then Exit Sub

    On Error GoTo RowFallback
    With wsUnitPrice.Range(wsUnitPrice.Cells(segStart, targetCol), wsUnitPrice.Cells(segEnd, targetCol))
        .FormulaR1C1 = formulaR1C1
        .NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
        .Interior.ColorIndex = xlColorIndexNone
        .ShrinkToFit = False
    End With
    Exit Sub

RowFallback:
    Err.Clear
    Dim ratioAddress As String
    Dim rowIndex As Long
    ratioAddress = GetVendorOutsourceRatioAddress(wsInfo, valueColumn)
    For rowIndex = segStart To segEnd
        ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, targetCol), isDayColumn, _
                                 wsUnitPrice, rowIndex, ratioAddress
    Next rowIndex
End Sub

Private Sub ApplyVendorUnitPriceBaseRowBorders(ByVal wsUnitPrice As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByVal changedCells As Range)
    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) > 0 Then
                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, 1), _
                                            wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_WIDTH_COL)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            End If
        End If
    Next changedCell

    RefreshVendorUnitPriceBordersForSheet wsUnitPrice, wsInfo
End Sub

Private Sub ApplyVendorUnitPriceNewRowFill(ByVal wsUnitPrice As Worksheet, _
                                            ByVal wsInfo As Worksheet, _
                                            ByVal changedCells As Range)
    Dim lastFillCol As Long
    lastFillCol = GetVendorUnitPriceInitialFillLastColumn(wsInfo)

    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(changedCell.value))) > 0 Then
                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                            wsUnitPrice.Cells(changedCell.Row, lastFillCol)).Interior
                    .Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_B)
                End With
            End If
        End If
    Next changedCell
End Sub

Private Function GetVendorUnitPriceInitialFillLastColumn(ByVal wsInfo As Worksheet) As Long
    Dim lastFillCol As Long
    lastFillCol = VENDOR_UNIT_PRICE_INITIAL_FILL_LAST_COL

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim nightCol As Long
            nightCol = VendorUnitPriceDayColumnByValueColumn(valueColumn) + 1
            If nightCol > lastFillCol Then lastFillCol = nightCol
        End If
    Next vendorIndex

    GetVendorUnitPriceInitialFillLastColumn = lastFillCol
End Function

Private Sub ResetClearedVendorUnitPriceRows(ByVal wsUnitPrice As Worksheet, _
                                             ByVal wsInfo As Worksheet, _
                                             ByVal changedCells As Range)
    Dim lastResetCol As Long
    lastResetCol = GetVendorUnitPriceInitialFillLastColumn(wsInfo)

    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(changedCell.value))) = 0 Then
                Dim resetRange As Range
                Set resetRange = wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, 1), _
                                                   wsUnitPrice.Cells(changedCell.Row, lastResetCol))
                resetRange.Borders.LineStyle = xlNone

                wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                  wsUnitPrice.Cells(changedCell.Row, lastResetCol)).Interior.ColorIndex = xlColorIndexNone

                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_FIRST_DAY_COL), _
                                       wsUnitPrice.Cells(changedCell.Row, lastResetCol))
                    .ClearContents
                    .NumberFormat = "General"
                    .ShrinkToFit = False
                End With
            End If
        End If
    Next changedCell
End Sub

Private Sub RefreshVendorUnitPriceBordersForSheet(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet)
    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim dayCol As Long
            dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
            ApplyVendorUnitPriceBorders wsUnitPrice, dayCol, dayCol + 1
        End If
    Next vendorIndex
End Sub

Private Sub HandleVendorUnitPriceSourceChanges(ByVal wsUnitPrice As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByVal changedCells As Range, _
                                                ByVal isDayColumn As Boolean)
    Dim sourceCell As Range
    For Each sourceCell In changedCells.Cells
        If sourceCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) = 0 Then
                GoTo ContinueNextSourceCell
            End If

            If HasNumericVendorUnitPriceSource(sourceCell) Then
                sourceCell.Interior.ColorIndex = xlColorIndexNone
                ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, sourceCell.Row, isDayColumn
            ElseIf IsVendorUnitPriceSourceCellBlank(sourceCell) Then
                ApplyVendorUnitPriceSourceGreyFill sourceCell
                ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, sourceCell.Row, isDayColumn
            End If
        End If
ContinueNextSourceCell:
    Next sourceCell
End Sub

' 独自工種などプログラム追記行は B 列の Worksheet_Change を経由しないため、
' E/F 入力時に手入力追記と同じ新規行初期化を補完する。
Private Sub EnsureVendorUnitPriceNewRowFillForSourceRows(ByVal wsUnitPrice As Worksheet, _
                                                         ByVal wsInfo As Worksheet, _
                                                         ByVal sourceChangedCells As Range)
    If wsUnitPrice Is Nothing Then Exit Sub
    If wsInfo Is Nothing Then Exit Sub
    If sourceChangedCells Is Nothing Then Exit Sub

    Dim bColRange As Range
    Dim sourceCell As Range
    For Each sourceCell In sourceChangedCells.Cells
        If sourceCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) > 0 Then
                If bColRange Is Nothing Then
                    Set bColRange = wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL)
                Else
                    Set bColRange = Union(bColRange, wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL))
                End If
            End If
        End If
    Next sourceCell

    If Not bColRange Is Nothing Then
        ApplyVendorUnitPriceNewRowFill wsUnitPrice, wsInfo, bColRange
    End If
End Sub

Private Sub ApplyVendorUnitPriceSourceRowsForRange(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet, _
                                                   ByVal firstRow As Long, _
                                                   ByVal lastRow As Long)
    Dim rowIndex As Long
    For rowIndex = firstRow To lastRow
        If Len(Trim$(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) = 0 Then GoTo ContinueRow

        ApplyVendorUnitPriceSourceRowIfNeeded wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_UNIT_COL, True
        ApplyVendorUnitPriceSourceRowIfNeeded wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_WIDTH_COL, False
ContinueRow:
    Next rowIndex
End Sub

Private Sub ApplyVendorUnitPriceSourceRowIfNeeded(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal wsInfo As Worksheet, _
                                                  ByVal rowIndex As Long, _
                                                  ByVal sourceCol As Long, _
                                                  ByVal isDayColumn As Boolean)
    Dim sourceCell As Range
    Set sourceCell = wsUnitPrice.Cells(rowIndex, sourceCol)

    If HasNumericVendorUnitPriceSource(sourceCell) Then
        sourceCell.Interior.ColorIndex = xlColorIndexNone
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    ElseIf IsVendorUnitPriceSourceCellBlank(sourceCell) Then
        ApplyVendorUnitPriceSourceGreyFill sourceCell
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    End If
End Sub

Private Function HasNumericVendorUnitPriceSource(ByVal sourceCell As Range) As Boolean
    If sourceCell Is Nothing Then Exit Function
    If IsError(sourceCell.Value2) Then Exit Function
    If Len(Trim$(CStr(sourceCell.Value2))) = 0 Then Exit Function
    HasNumericVendorUnitPriceSource = IsNumeric(sourceCell.Value2)
End Function

Private Function IsVendorUnitPriceSourceCellBlank(ByVal sourceCell As Range) As Boolean
    If sourceCell Is Nothing Then
        IsVendorUnitPriceSourceCellBlank = True
        Exit Function
    End If
    If IsError(sourceCell.Value2) Then Exit Function
    IsVendorUnitPriceSourceCellBlank = (Len(Trim$(CStr(sourceCell.Value2))) = 0)
End Function

Private Sub ApplyVendorUnitPriceSourceGreyFill(ByVal sourceCell As Range)
    sourceCell.Interior.Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                                    VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                                    VENDOR_UNIT_PRICE_FILL_COLOR_B)
End Sub

Private Sub ApplyVendorUnitPriceCellsForSourceRow(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet, _
                                                   ByVal rowIndex As Long, _
                                                   ByVal isDayColumn As Boolean)
    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim targetCol As Long
            targetCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
            If Not isDayColumn Then targetCol = targetCol + 1

            ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, targetCol), _
                                     isDayColumn, wsUnitPrice, rowIndex, _
                                     GetVendorOutsourceRatioAddress(wsInfo, valueColumn)
        End If
    Next vendorIndex
End Sub

Private Sub ApplyVendorUnitPriceCell(ByVal targetCell As Range, _
                                       ByVal isDayColumn As Boolean, _
                                       ByVal wsUnitPrice As Worksheet, _
                                       ByVal rowIndex As Long, _
                                       ByVal ratioAddress As String)
    Dim sourceCol As Long
    Dim workTypeName As String
    If isDayColumn Then
        sourceCol = VENDOR_UNIT_PRICE_REF_UNIT_COL
    Else
        sourceCol = VENDOR_UNIT_PRICE_REF_WIDTH_COL
    End If

    workTypeName = CommonNormalizeText(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_WORK_TYPE_COL).value))

    With targetCell
        .ShrinkToFit = False
        .Interior.ColorIndex = xlColorIndexNone

        If IsVendorUnitPriceSourceBlank(wsUnitPrice, rowIndex, sourceCol) Then
            ApplyVendorUnitPriceGreyFill targetCell
            Exit Sub
        End If

        If InStr(1, workTypeName, VendorWasteDisposalKeywordText(), vbTextCompare) > 0 Then
            ApplyVendorUnitPriceGreyFill targetCell
            Exit Sub
        End If

        .Formula = BuildVendorUnitPriceFormula(wsUnitPrice, rowIndex, sourceCol, ratioAddress)
        .NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
    End With
End Sub

Private Sub ApplyVendorUnitPriceGreyFill(ByVal targetCell As Range)
    With targetCell
        .ClearContents
        .NumberFormat = "General"
        .Interior.Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_B)
    End With
End Sub

Private Sub ApplyVendorUnitPriceFont(ByVal wsUnitPrice As Worksheet, _
                                     ByVal dayCol As Long, _
                                     ByVal nightCol As Long)
    Dim lastRow As Long
    lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then lastRow = VENDOR_UNIT_PRICE_DATA_START_ROW

    Dim fontRange As Range
    Set fontRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                      wsUnitPrice.Cells(lastRow, nightCol))

    With fontRange.Font
        .Name = VendorUnitPriceFontNameText()
        On Error Resume Next
        .NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Private Sub ApplyVendorUnitPriceBorders(ByVal wsUnitPrice As Worksheet, _
                                        ByVal dayCol As Long, _
                                        ByVal nightCol As Long)
    Dim lastRow As Long
    lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then Exit Sub

    Dim borderRange As Range
    Set borderRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                        wsUnitPrice.Cells(lastRow, nightCol))

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

Private Sub EnsureApplicationCalculationAutomatic()
    ' 計算モードはイベント/呼び出し側で一括管理する。
    ' 一括書込み中に自動計算へ強制すると、セル書込みごとに再計算が走り低速化するため、
    ' ここでは計算モードを変更しない(各処理終端の .Calculate と呼び出し側の復元で再計算)。
End Sub

Private Function IsVendorUnitPriceSourceBlank(ByVal wsUnitPrice As Worksheet, _
                                              ByVal rowIndex As Long, _
                                              ByVal sourceCol As Long) As Boolean
    IsVendorUnitPriceSourceBlank = _
        (Len(Trim$(CStr(wsUnitPrice.Cells(rowIndex, sourceCol).value))) = 0)
End Function

Private Function BuildVendorUnitPriceFormula(ByVal wsUnitPrice As Worksheet, _
                                             ByVal rowIndex As Long, _
                                             ByVal sourceCol As Long, _
                                             ByVal ratioAddress As String) As String
    Dim unitCellRef As String
    unitCellRef = wsUnitPrice.Cells(rowIndex, sourceCol).Address(False, False)

    BuildVendorUnitPriceFormula = "=ROUND(" & unitCellRef & "*(" & ratioAddress & ")," & _
                                  "-INT(LOG10(" & unitCellRef & "*(" & ratioAddress & ")))+2)"
End Function

Private Function BuildVendorUnitPriceFormulaR1C1(ByVal isDayColumn As Boolean, _
                                                 ByVal targetCol As Long, _
                                                 ByVal ratioAddress As String) As String
    Dim sourceCol As Long
    If isDayColumn Then
        sourceCol = VENDOR_UNIT_PRICE_REF_UNIT_COL
    Else
        sourceCol = VENDOR_UNIT_PRICE_REF_WIDTH_COL
    End If

    Dim sourceOffset As Long
    sourceOffset = sourceCol - targetCol

    BuildVendorUnitPriceFormulaR1C1 = "=ROUND(RC[" & sourceOffset & "]*(" & ratioAddress & ")," & _
        "-INT(LOG10(RC[" & sourceOffset & "]*(" & ratioAddress & ")))+2)"
End Function

Private Function BuildVendorUnitPriceHeaderText(ByVal wsInfo As Worksheet) As String
    BuildVendorUnitPriceHeaderText = Trim$(CStr(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value)) & _
                                     CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).value)) & _
                                     VendorUnitPriceOutsourceLabelText()
End Function

Private Function GetVendorOutsourceRatioAddress(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As String
    GetVendorOutsourceRatioAddress = "'" & Replace$(wsInfo.Name, "'", "''") & "'!" & _
                                     wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn).Address(True, True)
End Function

Private Function GetVendorUnitPriceLastDataRow(ByVal wsUnitPrice As Worksheet) As Long
    GetVendorUnitPriceLastDataRow = wsUnitPrice.Cells(wsUnitPrice.rows.Count, VENDOR_UNIT_PRICE_LAST_ROW_COL).End(xlUp).Row
End Function

Private Function VendorUnitPriceDayColumnByValueColumn(ByVal valueColumn As Long) As Long
    Dim vendorIndex As Long
    vendorIndex = GetVendorIndexFromValueColumn(valueColumn)
    If vendorIndex < 1 Then vendorIndex = 1
    VendorUnitPriceDayColumnByValueColumn = VENDOR_UNIT_PRICE_FIRST_DAY_COL + ((vendorIndex - 1) * 2)
End Function

Private Function GetVendorIndexFromValueColumn(ByVal valueColumn As Long) As Long
    Dim i As Long
    For i = 1 To MAX_VENDOR_BLOCK_COUNT
        If VendorValueColumnByIndex(i) = valueColumn Then
            GetVendorIndexFromValueColumn = i
            Exit Function
        End If
    Next i
End Function

Private Function IsRailConstructionVendorBlock(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    IsRailConstructionVendorBlock = _
        (StrComp(CommonNormalizeText(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueColumn).value)), _
                 VendorRailConstructionText(), vbTextCompare) = 0)
End Function

Private Function HasVendorOutsourceRatio(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    HasVendorOutsourceRatio = (Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn).value))) > 0)
End Function

Private Sub SafeUnmergeRange(ByVal targetRange As Range)
    If targetRange Is Nothing Then Exit Sub

    On Error Resume Next
    Dim cell As Range
    For Each cell In targetRange.Cells
        If cell.MergeCells Then cell.MergeArea.UnMerge
    Next cell
    On Error GoTo 0
End Sub

Private Function VendorRailConstructionText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
    End If
    VendorRailConstructionText = cached
End Function

Private Function VendorUnitPriceOutsourceLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H5358) & ChrW$(&H4FA1)
    End If
    VendorUnitPriceOutsourceLabelText = cached
End Function

Private Function VendorUnitPriceOutsourceRatioLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&HFF1D)
    End If
    VendorUnitPriceOutsourceRatioLabelText = cached
End Function

Private Function VendorJrUnitPriceLabelText() As String
    VendorJrUnitPriceLabelText = "JR"
End Function

Private Function VendorUnitPriceFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    VendorUnitPriceFontNameText = cached
End Function

Private Function VendorUnitPriceDayLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H663C) & ChrW$(&H9593)
    VendorUnitPriceDayLabelText = cached
End Function

Private Function VendorUnitPriceNightLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H591C) & ChrW$(&H9593)
    VendorUnitPriceNightLabelText = cached
End Function

Private Function VendorWasteDisposalKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H7523) & ChrW$(&H5EC3) & ChrW$(&H51E6) & ChrW$(&H7406)
    VendorWasteDisposalKeywordText = cached
End Function

Private Function GetVendorTargetCell(ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range) As Range
    Dim resolvedCell As Range

    If Not targetCell Is Nothing Then
        Set resolvedCell = ResolveVendorNameChangeCell(wsInfo, targetCell)
        If Not resolvedCell Is Nothing Then
            Set GetVendorTargetCell = resolvedCell
            mVendorTargetAddress = GetVendorTargetCell.Address(False, False)
            Exit Function
        End If
    End If

    If mVendorTargetAddress <> "" Then
        On Error Resume Next
        Set GetVendorTargetCell = wsInfo.Range(mVendorTargetAddress)
        On Error GoTo 0
        If Not GetVendorTargetCell Is Nothing Then
            Set resolvedCell = ResolveVendorNameChangeCell(wsInfo, GetVendorTargetCell)
            If resolvedCell Is Nothing Then Set GetVendorTargetCell = Nothing
        End If
    End If

    If GetVendorTargetCell Is Nothing Then
        If Not ActiveCell Is Nothing Then
            If ActiveCell.Worksheet Is wsInfo Then
                Set resolvedCell = ResolveVendorNameChangeCell(wsInfo, ActiveCell)
                If Not resolvedCell Is Nothing Then Set GetVendorTargetCell = resolvedCell
            End If
        End If
    End If

    If GetVendorTargetCell Is Nothing Then Set GetVendorTargetCell = wsInfo.Range(BASIC_INFO_VENDOR_NAME_CELL)
    mVendorTargetAddress = GetVendorTargetCell.Address(False, False)
End Function

Private Function VendorWritableValueCell(ByVal wsInfo As Worksheet, ByVal rowIndex As Long, ByVal valueColumn As Long) As Range
    Dim targetCell As Range
    Set targetCell = wsInfo.Cells(rowIndex, valueColumn)
    If targetCell.MergeCells Then
        Dim mergeArea As Range
        Set mergeArea = targetCell.MergeArea
        If valueColumn >= mergeArea.Column And _
           valueColumn <= mergeArea.Column + mergeArea.Columns.Count - 1 Then
            Set VendorWritableValueCell = mergeArea.Cells(1, 1)
            Exit Function
        End If
    End If
    Set VendorWritableValueCell = targetCell
End Function

Private Function CountExistingVendorBlocks(ByVal wsInfo As Worksheet) As Long
    Dim i As Long
    For i = 1 To MAX_VENDOR_BLOCK_COUNT
        If InStr(1, CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i)).value), _
                 VendorInfoHeaderPrefixText(), vbTextCompare) = 0 Then
            Exit For
        End If
        CountExistingVendorBlocks = i
    Next i
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

Private Sub ApplyVendorRow10ValueCellFormat(ByVal valueCell As Range)
    If valueCell Is Nothing Then Exit Sub

    Dim vendorIndex As Long
    vendorIndex = GetVendorIndexFromValueColumn(valueCell.Column)
    If vendorIndex <= 0 Then Exit Sub

    mod_VendorInfoColors.ApplyVendorInfoRow10Color valueCell.Worksheet, vendorIndex
End Sub

Private Function VendorBlockNeedsPresentationRestore(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If vendorIndex < 2 Then Exit Function

    Dim srcLabelCol As Long
    Dim dstLabelCol As Long
    Dim dstValueCol As Long
    srcLabelCol = VendorLabelColumnByIndex(1)
    dstLabelCol = VendorLabelColumnByIndex(vendorIndex)
    dstValueCol = VendorValueColumnByIndex(vendorIndex)

    If Len(Trim$(CStr(wsInfo.Cells(12, srcLabelCol).value))) > 0 Then
        If Len(Trim$(CStr(wsInfo.Cells(12, dstLabelCol).value))) = 0 Then
            VendorBlockNeedsPresentationRestore = True
            Exit Function
        End If
    End If

    If wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, dstValueCol).Borders(xlEdgeLeft).LineStyle = xlNone Then
        VendorBlockNeedsPresentationRestore = True
    End If
End Function

Private Sub EnsureVendorBlockFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    CopyVendorBlockFromTemplate wsInfo, destVendorIndex
    RestoreVendorBlockPresentationFromTemplate wsInfo, destVendorIndex
End Sub

Private Sub CopyVendorBlockFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim sourceRange As Range
    Dim destRange As Range
    Set sourceRange = VendorBlockRangeByIndex(wsInfo, 1)
    Set destRange = VendorBlockRangeByIndex(wsInfo, destVendorIndex)

    SafeUnmergeRange destRange
    sourceRange.Copy Destination:=destRange
    Application.CutCopyMode = False
    CopyVendorBlockMergeAreasFromTemplate wsInfo, destVendorIndex
    CopyVendorBlockFormatsFromTemplate wsInfo, destVendorIndex

    wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(destVendorIndex)).value = VendorInfoHeaderText(destVendorIndex)

    ClearVendorInfoBlock VendorNameCellByIndex(wsInfo, destVendorIndex)

    wsInfo.Cells(BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW, VendorValueColumnByIndex(destVendorIndex)).ClearContents
    wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorValueColumnByIndex(destVendorIndex)).ClearContents
End Sub

Private Sub CopyVendorBlockMergeAreasFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim srcRange As Range
    Set srcRange = VendorBlockRangeByIndex(wsInfo, 1)
    Dim colOffset As Long
    colOffset = VendorLabelColumnByIndex(destVendorIndex) - VendorLabelColumnByIndex(1)

    On Error Resume Next
    Dim cell As Range
    For Each cell In srcRange.Cells
        If cell.MergeCells Then
            Dim mergeArea As Range
            Set mergeArea = cell.MergeArea
            If cell.Row = mergeArea.Row And cell.Column = mergeArea.Column Then
                Dim destMerge As Range
                Set destMerge = wsInfo.Range(wsInfo.Cells(mergeArea.Row, mergeArea.Column + colOffset), _
                                             wsInfo.Cells(mergeArea.Row + mergeArea.Rows.Count - 1, _
                                                          mergeArea.Column + colOffset + mergeArea.Columns.Count - 1))
                SafeUnmergeRange destMerge
                destMerge.Merge
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

Private Function VendorBlockRangeByIndex(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Range
    Set VendorBlockRangeByIndex = wsInfo.Range( _
        wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(vendorIndex)), _
        wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorSpacerColumnByIndex(vendorIndex)))
End Function

Private Sub RestoreVendorBlockPresentationFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim sourceRange As Range
    Dim destRange As Range
    Set sourceRange = VendorBlockRangeByIndex(wsInfo, 1)
    Set destRange = VendorBlockRangeByIndex(wsInfo, destVendorIndex)

    CopyRangeBorders sourceRange, destRange
    RestoreVendorBlockLabelTextsFromTemplate wsInfo, destVendorIndex

    Dim srcLabelCol As Long
    Dim dstLabelCol As Long
    Dim srcSpacerCol As Long
    Dim dstSpacerCol As Long
    srcLabelCol = VendorLabelColumnByIndex(1)
    dstLabelCol = VendorLabelColumnByIndex(destVendorIndex)
    srcSpacerCol = VendorSpacerColumnByIndex(1)
    dstSpacerCol = VendorSpacerColumnByIndex(destVendorIndex)

    CopyRangeFormats wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, srcLabelCol), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, srcLabelCol)), _
                     wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, dstLabelCol), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, dstLabelCol))
    CopyRangeFormats wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, srcSpacerCol), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, srcSpacerCol)), _
                     wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, dstSpacerCol), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, dstSpacerCol))
End Sub

Private Sub RestoreVendorBlockLabelTextsFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    Dim srcLabelCol As Long
    Dim dstLabelCol As Long
    srcLabelCol = VendorLabelColumnByIndex(1)
    dstLabelCol = VendorLabelColumnByIndex(destVendorIndex)

    Dim rowIndex As Long
    For rowIndex = BASIC_INFO_VENDOR_BLOCK_TOP_ROW + 1 To BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW
        If rowIndex <> 30 Then
            wsInfo.Cells(rowIndex, dstLabelCol).value = wsInfo.Cells(rowIndex, srcLabelCol).value
        End If
    Next rowIndex

    For rowIndex = BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW To BASIC_INFO_VENDOR_TOTAL_ROW
        wsInfo.Cells(rowIndex, dstLabelCol).value = wsInfo.Cells(rowIndex, srcLabelCol).value
    Next rowIndex
End Sub

Private Sub CopyRangeFormats(ByVal sourceRange As Range, ByVal destRange As Range)
    If sourceRange Is Nothing Then Exit Sub
    If destRange Is Nothing Then Exit Sub
    If sourceRange.Rows.Count <> destRange.Rows.Count Then Exit Sub
    If sourceRange.Columns.Count <> destRange.Columns.Count Then Exit Sub

    sourceRange.Copy
    destRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub RestoreVendorBlockValueColumnRightBorders(ByVal wsInfo As Worksheet, ByVal vendorCount As Long)
    If wsInfo Is Nothing Then Exit Sub
    If vendorCount < 1 Then Exit Sub

    Dim templateValueCol As Long
    templateValueCol = VendorValueColumnByIndex(1)

    Dim rowIndex As Long
    Dim vendorIndex As Long

    For rowIndex = BASIC_INFO_VENDOR_BLOCK_TOP_ROW To BASIC_INFO_VENDOR_TOTAL_ROW
        Dim templateCell As Range
        Set templateCell = wsInfo.Cells(rowIndex, templateValueCol)

        For vendorIndex = 1 To vendorCount
            CopyCellBorderEdge templateCell, _
                wsInfo.Cells(rowIndex, VendorValueColumnByIndex(vendorIndex)), xlEdgeRight
        Next vendorIndex
    Next rowIndex
End Sub

Private Sub CopyCellBorderEdge(ByVal sourceCell As Range, ByVal destCell As Range, ByVal edgeId As Long)
    On Error Resume Next
    With destCell.Borders(edgeId)
        .LineStyle = sourceCell.Borders(edgeId).LineStyle
        .Weight = sourceCell.Borders(edgeId).Weight
        .Color = sourceCell.Borders(edgeId).Color
    End With
    On Error GoTo 0
End Sub

Private Sub CopyRangeBorders(ByVal sourceRange As Range, ByVal destRange As Range)
    If sourceRange Is Nothing Then Exit Sub
    If destRange Is Nothing Then Exit Sub
    If sourceRange.Rows.Count <> destRange.Rows.Count Then Exit Sub
    If sourceRange.Columns.Count <> destRange.Columns.Count Then Exit Sub

    Dim rowOffset As Long
    Dim colOffset As Long
    Dim edgeIndex As Long
    Dim borderEdgeIds As Variant
    borderEdgeIds = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                          xlInsideVertical, xlInsideHorizontal)

    For rowOffset = 1 To sourceRange.Rows.Count
        For colOffset = 1 To sourceRange.Columns.Count
            Dim sourceCell As Range
            Dim destCell As Range
            Set sourceCell = sourceRange.Cells(rowOffset, colOffset)
            Set destCell = destRange.Cells(rowOffset, colOffset)

            For edgeIndex = LBound(borderEdgeIds) To UBound(borderEdgeIds)
                On Error Resume Next
                With destCell.Borders(borderEdgeIds(edgeIndex))
                    .LineStyle = sourceCell.Borders(borderEdgeIds(edgeIndex)).LineStyle
                    .Weight = sourceCell.Borders(borderEdgeIds(edgeIndex)).Weight
                    .Color = sourceCell.Borders(borderEdgeIds(edgeIndex)).Color
                End With
                On Error GoTo 0
            Next edgeIndex
        Next colOffset
    Next rowOffset
End Sub

Private Sub CopyVendorBlockFormatsFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim rowCount As Long
    rowCount = BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW - BASIC_INFO_VENDOR_BLOCK_TOP_ROW + 1

    Dim sourceBlock As Range
    Dim destBlock As Range
    Set sourceBlock = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, BASIC_INFO_VENDOR_BLOCK_LABEL_COL).Resize(rowCount, 2)
    Set destBlock = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(destVendorIndex)).Resize(rowCount, 2)

    sourceBlock.Copy
    destBlock.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub CopyVendorBlockTotalRowsFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim sourceRange As Range
    Set sourceRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW, BASIC_INFO_VENDOR_BLOCK_LABEL_COL), _
                                   wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, BASIC_INFO_VENDOR_BLOCK_VALUE_COL))

    Dim destRange As Range
    Set destRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW, VendorLabelColumnByIndex(destVendorIndex)), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorValueColumnByIndex(destVendorIndex)))

    SafeUnmergeRange destRange
    destRange.ClearContents
    sourceRange.Copy Destination:=destRange
    Application.CutCopyMode = False
End Sub

Private Sub ClearVendorWorkTypeWhenCompanyEmpty(ByVal wsInfo As Worksheet, Optional ByVal vendorCount As Long = 0)
    If wsInfo Is Nothing Then Exit Sub
    If vendorCount <= 0 Then vendorCount = GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueCol As Long
        valueCol = VendorValueColumnByIndex(vendorIndex)
        If Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).value))) = 0 Then
            VendorWritableValueCell(wsInfo, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol).ClearContents
        End If
    Next vendorIndex
End Sub

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
        VendorWritableValueCell(targetCell.Worksheet, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).ClearContents
        .Range(.Cells(11, targetCell.Column), .Cells(16, targetCell.Column)).ClearContents
        .Range(.Cells(18, targetCell.Column), .Cells(23, targetCell.Column)).ClearContents
        .Range(.Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column), _
               .Cells(BASIC_INFO_VENDOR_TOTAL_ROW, targetCell.Column)).ClearContents
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
        ClearVendorBlockColumns wsInfo, i
    Next i
End Sub

Private Sub ClearVendorBlockColumns(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If vendorIndex < 1 Or vendorIndex > MAX_VENDOR_BLOCK_COUNT Then Exit Sub

    Dim labelCol As Long
    labelCol = VendorLabelColumnByIndex(vendorIndex)

    Dim clearRange As Range
    Set clearRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, labelCol), _
                                  wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorSpacerColumnByIndex(vendorIndex)))

    SafeUnmergeRange clearRange
    clearRange.ClearContents
    clearRange.Interior.Color = RGB(6, 17, 29)
    clearRange.Borders.LineStyle = xlNone
End Sub

Private Function VendorInfoHeaderText(ByVal vendorIndex As Long) As String
    VendorInfoHeaderText = "  " & ChrW$(&H25B8) & "  " & VendorInfoHeaderPrefixText() & CStr(vendorIndex)
End Function

Private Function VendorInfoHeaderPrefixText() As String
    VendorInfoHeaderPrefixText = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H60C5) & ChrW$(&H5831) & "-"
End Function

Private Function LoadVendorRows(ByVal BranchName As String) As Collection
    Dim cacheKey As String
    cacheKey = CommonNormalizeText(BranchName)

    If mVendorRowsCache Is Nothing Then
        Set mVendorRowsCache = CreateObject("Scripting.Dictionary")
        mVendorRowsCache.CompareMode = vbTextCompare
    End If
    If cacheKey <> "" Then
        If mVendorRowsCache.Exists(cacheKey) Then
            Set LoadVendorRows = mVendorRowsCache(cacheKey)
            Exit Function
        End If
    End If

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
    If cacheKey <> "" And Not LoadVendorRows Is Nothing Then
        If mVendorRowsCache.Exists(cacheKey) Then
            Set mVendorRowsCache(cacheKey) = LoadVendorRows
        Else
            mVendorRowsCache.Add cacheKey, LoadVendorRows
        End If
    End If
End Function

' 業者マスタ(外部ファイル)の読込結果を支店名キーで保持し、1操作内の重複ADO接続を防ぐ。
' 基本情報シートのActivate時にクリアして最新化する。
Public Sub ClearVendorRowsCache()
    Set mVendorRowsCache = Nothing
    Set mVendorNameIndexCache = Nothing
    mod_Construction_Order_Import.ClearVendorAliasMapCache
End Sub

Private Function GetVendorNameIndex(ByVal BranchName As String, ByVal vendorRows As Collection) As Object
    Dim cacheKey As String
    cacheKey = CommonNormalizeText(BranchName)

    If mVendorNameIndexCache Is Nothing Then
        Set mVendorNameIndexCache = CreateObject("Scripting.Dictionary")
        mVendorNameIndexCache.CompareMode = vbTextCompare
    End If
    If cacheKey <> "" Then
        If mVendorNameIndexCache.Exists(cacheKey) Then
            Set GetVendorNameIndex = mVendorNameIndexCache(cacheKey)
            Exit Function
        End If
    End If

    Dim nameIndex As Object
    Set nameIndex = CreateObject("Scripting.Dictionary")
    nameIndex.CompareMode = vbTextCompare

    Dim rowData As Variant
    Dim normalizedName As String
    Dim rowIndex As Long
    rowIndex = 0
    For Each rowData In vendorRows
        rowIndex = rowIndex + 1
        normalizedName = CommonNormalizeText(CStr(rowData(0)))
        If normalizedName <> "" Then
            If Not nameIndex.Exists(normalizedName) Then
                nameIndex.Add normalizedName, rowIndex
            End If
        End If
    Next rowData

    If cacheKey <> "" Then
        mVendorNameIndexCache.Add cacheKey, nameIndex
    End If
    Set GetVendorNameIndex = nameIndex
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
    Dim columnOValue As String
    columnOValue = GetVendorMasterColumnOValue(recordset, sheetName, _
                                               CommonNzText(CommonGetAdoFieldValue(recordset, 1)))
    mod_DebugLog.Log "[VendorMaster] BuildVendorRow Count=" & recordset.Fields.Count & _
                     " O=[" & columnOValue & "] vendor=[" & CommonNzText(CommonGetAdoFieldValue(recordset, 1)) & "]"

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
                                        CommonNormalizeText(sheetName), _
                                        columnOValue)
End Function

Private Function GetVendorMasterColumnOValue(ByVal recordset As Object, _
                                             ByVal sheetName As String, _
                                             ByVal vendorName As String) As String
    Dim valueText As String

    On Error Resume Next
    valueText = CommonNzText(recordset.Fields(VENDOR_MASTER_ADO_COLUMN_O_NAME).value)
    If valueText = "" Then valueText = CommonNzText(CommonGetAdoFieldValue(recordset, VENDOR_MASTER_EXCEL_COLUMN_O - 1))
    If valueText = "" Then valueText = CommonNzText(recordset.Fields(VENDOR_MASTER_EXCEL_COLUMN_O - 1).value)
    If valueText = "" Then valueText = CommonNzText(recordset.Fields(VENDOR_MASTER_EXCEL_COLUMN_O).value)
    On Error GoTo 0

    If valueText <> "" Then
        GetVendorMasterColumnOValue = valueText
        Exit Function
    End If

    GetVendorMasterColumnOValue = GetVendorMasterColumnOValueFromExcel(sheetName, vendorName)
End Function

Private Function GetVendorMasterColumnOValueFromExcel(ByVal sheetName As String, ByVal vendorName As String) As String
    Dim sourceFilePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim normalizedVendorName As String
    Dim wasAlreadyOpen As Boolean
    Dim screenUpdating As Boolean

    normalizedVendorName = CommonNormalizeText(vendorName)
    If normalizedVendorName = "" Then Exit Function

    Set wb = Nothing
    sourceFilePath = GetVendorMasterFilePath()
    If sourceFilePath = "" Then Exit Function

    wasAlreadyOpen = False
    screenUpdating = Application.screenUpdating
    On Error GoTo Cleanup

    Dim openWorkbook As Workbook
    For Each openWorkbook In Application.Workbooks
        If StrComp(openWorkbook.FullName, sourceFilePath, vbTextCompare) = 0 Then
            Set wb = openWorkbook
            wasAlreadyOpen = True
            Exit For
        End If
    Next openWorkbook

    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open( _
            fileName:=sourceFilePath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            Notify:=False, _
            AddToMru:=False)
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.worksheets(sheetName)
    On Error GoTo Cleanup
    If ws Is Nothing Then GoTo Cleanup

    For rowIndex = VENDOR_SOURCE_START_ROW To VENDOR_SOURCE_END_ROW
        If StrComp(CommonNormalizeText(CStr(ws.Cells(rowIndex, 1).value)), normalizedVendorName, vbTextCompare) = 0 Then
            GetVendorMasterColumnOValueFromExcel = CommonNzText(ws.Cells(rowIndex, VENDOR_MASTER_EXCEL_COLUMN_O).value)
            Exit For
        End If
    Next rowIndex

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not wasAlreadyOpen Then wb.Close SaveChanges:=False
    End If
    Application.screenUpdating = screenUpdating
    On Error GoTo 0
End Function

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
    GetAdoVendorRangeName = "[" & Replace$(sheetName, "]", "]]") & "$A" & VENDOR_SOURCE_START_ROW & ":O" & VENDOR_SOURCE_END_ROW & "]"
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
        cached = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF)
    End If
    VendorMasterFolderNameText = cached
End Function

Private Function VendorMasterFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & _
                 "(" & ChrW$(&H5168) & ChrW$(&H793E) & ChrW$(&H7248) & ").xlsx"
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
