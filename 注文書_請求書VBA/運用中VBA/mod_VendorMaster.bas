Option Explicit

'==========================================================================
'  業者マスタ参照・業者情報選択モジュール
'  改修履歴: CHANGELOG.md 参照
'==========================================================================

Private Const VENDOR_LIST_COL As String = "AD"
Private Const VENDOR_LIST_START_ROW As Long = 2
Private Const BASIC_INFO_VENDOR_NAME_CELL As String = "F11"
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_LABEL_COL As Long = 5
Private Const BASIC_INFO_VENDOR_BLOCK_VALUE_COL As Long = 6
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW As Long = 31
Private Const BASIC_INFO_VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_LABEL_COL_WIDTH As Double = 26.38
Private Const BASIC_INFO_VENDOR_VALUE_COL_WIDTH As Double = 42.5
Private Const BASIC_INFO_VENDOR_SPACER_COL_WIDTH As Double = 0.92
Private Const BASIC_INFO_VENDOR_PERCENT_ROW As Long = 25
Private Const BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW As Long = 29
Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C22"
Private Const VENDOR_UNIT_PRICE_JR_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_START As Long = 5
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_END As Long = 6
Private Const VENDOR_UNIT_PRICE_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_NAME_ROW As Long = 5
Private Const VENDOR_UNIT_PRICE_LABEL_ROW As Long = 6
Private Const VENDOR_UNIT_PRICE_DATA_START_ROW As Long = 7
Private Const VENDOR_UNIT_PRICE_FIRST_DAY_COL As Long = 7
Private Const VENDOR_UNIT_PRICE_REF_UNIT_COL As Long = 5
Private Const VENDOR_UNIT_PRICE_REF_WIDTH_COL As Long = 6
Private Const VENDOR_UNIT_PRICE_WORK_TYPE_COL As Long = 3
Private Const VENDOR_UNIT_PRICE_LAST_ROW_COL As Long = 2
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_R As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_G As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_B As Long = 128
Private Const VENDOR_UNIT_PRICE_NUMBER_FORMAT As String = "#,##0"

Private Const MAX_VENDOR_BLOCK_COUNT As Long = 20
Private Const VENDOR_SOURCE_START_ROW As Long = 2
Private Const VENDOR_SOURCE_END_ROW As Long = 36
Private Const VENDOR_MASTER_ADO_COLUMN_O_NAME As String = "F15"
Private Const VENDOR_MASTER_EXCEL_COLUMN_O As Long = 15
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
        RefreshVendorUnitPriceForValueColumn wsInfo, targetCell.Column
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
'  ComboBox 用の業者候補を配列から設定
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
            If CommonNormalizeText(CStr(arr)) <> "" Then .AddItem CStr(arr)
        End If
        .LinkedCell = ""
        .ListRows = Application.Max(1, Application.Min(12, .ListCount))
        .MatchRequired = False
        .value = currentValue
    End With
End Sub

'--------------------------------------------------------------------------
'  業者名の入力規則リストを非表示列へ作成
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
    mod_DebugLog.Log "[VendorMaster] ApplyVendorRow F10=[" & CStr(rowData(14)) & "] vendor=[" & CStr(rowData(0)) & "]"
    With targetCell.Worksheet
        ApplyVendorRow10ValueCellFormat .Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column)
        .Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).value = rowData(14)
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
        Set destBlock = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i)).Resize(sourceBlock.rows.Count, sourceBlock.Columns.Count)
        sourceBlock.Copy Destination:=destBlock
        destBlock.Cells(1, 1).value = VendorInfoHeaderText(i)
        ClearVendorInfoBlock VendorNameCellByIndex(wsInfo, i)
        CopyVendorBlockFormatsFromTemplate wsInfo, i
    Next i

    ClearUnusedVendorBlocks wsInfo, vendorCount + 1

    Dim formatIndex As Long
    For formatIndex = 1 To vendorCount
        ApplyVendorRow10ValueCellFormat wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorValueColumnByIndex(formatIndex))
    Next formatIndex

    Application.CutCopyMode = False
    RefreshVendorListForBasicInfo wsInfo
    RefreshAllVendorUnitPricesForBasicInfo wsInfo
    Exit Sub

ExitHandler:
    Application.CutCopyMode = False
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
                               wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol))

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
    If Intersect(changedRange, GetVendorUnitPriceMonitorRange(wsInfo)) Is Nothing Then Exit Sub

    RefreshAllVendorUnitPricesForBasicInfo wsInfo
End Sub

Public Sub RefreshAllVendorUnitPricesForBasicInfo(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    EnsureApplicationCalculationAutomatic

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            RefreshVendorUnitPriceBlocksOnSheet wsUnitPrice, wsInfo, vendorCount
        End If
    Next wsUnitPrice

    On Error Resume Next
    targetBook.Calculate
    On Error GoTo 0
End Sub

Private Sub RefreshVendorUnitPriceForValueColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then
            If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn
            Else
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            End If
        End If
    Next wsUnitPrice
End Sub

Private Sub RefreshVendorUnitPriceBlocksOnSheet(ByVal wsUnitPrice As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByVal vendorCount As Long)
    Dim i As Long
    For i = 1 To vendorCount
        Dim valueColumn As Long
        Dim dayCol As Long
        Dim nightCol As Long
        valueColumn = VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1

        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn
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
        HasVendorOutsourceRatio(wsInfo, valueColumn)
End Function

Private Sub ApplyVendorUnitPriceBlockToSheet(ByVal wsUnitPrice As Worksheet, _
                                             ByVal wsInfo As Worksheet, _
                                             ByVal valueColumn As Long)
    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    ApplyVendorUnitPriceColumnWidths wsUnitPrice, dayCol, nightCol
    ApplyVendorUnitPriceMergedHeader wsUnitPrice, dayCol, nightCol, BuildVendorUnitPriceHeaderText(wsInfo)
    ApplyVendorUnitPriceMergedVendorName wsUnitPrice, dayCol, nightCol, _
        CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value)

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

Private Sub ClearVendorUnitPriceBlockOnSheet(ByVal wsUnitPrice As Worksheet, _
                                             ByVal dayCol As Long, _
                                             ByVal nightCol As Long)
    If wsUnitPrice Is Nothing Then Exit Sub

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

    Dim rowIndex As Long
    For rowIndex = VENDOR_UNIT_PRICE_DATA_START_ROW To lastRow
        ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, dayCol), True, _
                                 wsUnitPrice, rowIndex, ratioAddress
        ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, nightCol), False, _
                                 wsUnitPrice, rowIndex, ratioAddress
    Next rowIndex
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
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
End Sub

Private Sub EnsureApplicationCalculationAutomatic()
    On Error Resume Next
    If Application.Calculation <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
    End If
    On Error GoTo 0
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
    On Error Resume Next
    If targetRange.MergeCells Then targetRange.UnMerge
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

Private Sub ApplyVendorRow10ValueCellFormat(ByVal valueCell As Range)
    If valueCell Is Nothing Then Exit Sub

    With valueCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
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
        .Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).ClearContents
        .Range(.Cells(11, targetCell.Column), .Cells(16, targetCell.Column)).ClearContents
        .Range(.Cells(18, targetCell.Column), .Cells(23, targetCell.Column)).ClearContents
        .Range(.Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column), _
               .Cells(BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW, targetCell.Column)).ClearContents
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
