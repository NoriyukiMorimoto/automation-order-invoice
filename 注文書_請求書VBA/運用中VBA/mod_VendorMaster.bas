Option Explicit

Private Const VENDOR_LIST_COL As String = "AJ"
Private Const VENDOR_LIST_START_ROW As Long = 2
Private Const BASIC_INFO_VENDOR_NAME_CELL As String = "F11"
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_TOTAL_ROW As Long = 33
Private Const BASIC_INFO_VENDOR_PERCENT_ROW As Long = 25
Private Const VENDOR_SOURCE_START_ROW As Long = 2
Private Const VENDOR_SOURCE_END_ROW As Long = 500
Private Const VENDOR_MASTER_ADO_COLUMN_O_NAME As String = "F15"
Private Const VENDOR_MASTER_EXCEL_COLUMN_O As Long = 15
Private Const VENDOR_COMBO_NAME As String = "ComboBoxVendor"

Private mVendorPromptTime As Date
Private mVendorTargetAddress As String
Private mVendorRowsCache As Object
Private mVendorNameIndexCache As Object

Private Function BuildVendorRowFromAdoRecord(ByVal recordset As Object, ByVal sheetName As String) As Variant
    Dim columnOValue As String
    columnOValue = GetVendorMasterColumnOValue(recordset)
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

' O列(工事区分)は ADO のみで取得する。Excel ブックを開くフォールバックは行わない。

Private Function GetAdoVendorRangeName(ByVal sheetName As String) As String
    GetAdoVendorRangeName = "[" & Replace$(sheetName, "]", "]]") & "$A" & VENDOR_SOURCE_START_ROW & ":O" & VENDOR_SOURCE_END_ROW & "]"
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

Private Function GetVendorMasterColumnOValue(ByVal recordset As Object) As String
    Dim valueText As String

    On Error Resume Next
    valueText = CommonNzText(recordset.Fields(VENDOR_MASTER_ADO_COLUMN_O_NAME).value)
    If valueText = "" Then valueText = CommonNzText(CommonGetAdoFieldValue(recordset, VENDOR_MASTER_EXCEL_COLUMN_O))
    On Error GoTo 0

    GetVendorMasterColumnOValue = valueText
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

Public Function GetVendorNameIndex(ByVal BranchName As String, ByVal vendorRows As Collection) As Object
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

Public Function GetVendorTargetCell(ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range) As Range
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

Public Function HasVendorRows(ByVal rows As Collection) As Boolean
    On Error Resume Next
    HasVendorRows = Not rows Is Nothing
    If HasVendorRows Then HasVendorRows = (rows.Count > 0)
    On Error GoTo 0
End Function

Public Function LoadAllVendorRows() As Collection
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

Public Function LoadVendorRows(ByVal BranchName As String) As Collection
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

Private Function OrderInvoiceDocumentFolderText() As String
    OrderInvoiceDocumentFolderText = mod_common.CommonOrderInvoiceDocumentFolderText()
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

Private Function VendorMasterFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & _
                 "(" & ChrW$(&H5168) & ChrW$(&H793E) & ChrW$(&H7248) & ").xlsx"
    End If
    VendorMasterFileNameText = cached
End Function

Private Function VendorMasterFolderNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF)
    End If
    VendorMasterFolderNameText = cached
End Function

Public Function VendorWritableValueCell(ByVal wsInfo As Worksheet, ByVal rowIndex As Long, ByVal valueColumn As Long) As Range
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

Private Sub ApplyVendorRowToBasicInfo(ByVal targetCell As Range, ByVal rowData As Variant)
    mod_DebugLog.Log "[VendorMaster] ApplyVendorRow F10=[" & CStr(rowData(14)) & "] vendor=[" & CStr(rowData(0)) & "]"
    Dim wsInfo As Worksheet
    Set wsInfo = targetCell.Worksheet
    With wsInfo
        Dim workTypeCell As Range
        Set workTypeCell = VendorWritableValueCell(wsInfo, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column)
        mod_VendorBlockLayout.ApplyVendorRow10ValueCellFormat .Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column)
        workTypeCell.value = rowData(14)
        VendorWritableValueCell(wsInfo, 11, targetCell.Column).value = rowData(0)
        VendorWritableValueCell(wsInfo, 12, targetCell.Column).value = rowData(3)
        VendorWritableValueCell(wsInfo, 13, targetCell.Column).value = rowData(11)
        VendorWritableValueCell(wsInfo, 14, targetCell.Column).value = rowData(2)
        VendorWritableValueCell(wsInfo, 15, targetCell.Column).value = rowData(4)
        VendorWritableValueCell(wsInfo, 16, targetCell.Column).value = rowData(1)
        VendorWritableValueCell(wsInfo, 18, targetCell.Column).value = rowData(5)
        VendorWritableValueCell(wsInfo, 19, targetCell.Column).value = rowData(6)
        VendorWritableValueCell(wsInfo, 20, targetCell.Column).value = rowData(10)
        VendorWritableValueCell(wsInfo, 21, targetCell.Column).value = rowData(7)
        VendorWritableValueCell(wsInfo, 22, targetCell.Column).value = rowData(8)
        VendorWritableValueCell(wsInfo, 23, targetCell.Column).value = rowData(9)
        With VendorWritableValueCell(wsInfo, BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column)
            .NumberFormatLocal = "@"
            .value = "100" & ChrW$(&HFF05)
        End With
    End With

    mod_VendorUnitPrice.RefreshVendorUnitPriceForValueColumn targetCell.Worksheet, targetCell.Column
End Sub

Private Sub ClearAllVendorInfoBlocks(ByVal wsInfo As Worksheet)
    Dim vendorCells As Range
    Set vendorCells = GetVendorNameRange(wsInfo)

    Dim vendorCell As Range
    For Each vendorCell In vendorCells.Cells
        ClearVendorInfoBlock vendorCell
    Next vendorCell
End Sub

Public Sub ClearVendorInfoBlock(ByVal targetCell As Range)
    With targetCell.Worksheet
        VendorWritableValueCell(targetCell.Worksheet, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, targetCell.Column).ClearContents
        .Range(.Cells(11, targetCell.Column), .Cells(16, targetCell.Column)).ClearContents
        .Range(.Cells(18, targetCell.Column), .Cells(23, targetCell.Column)).ClearContents
        .Range(.Cells(BASIC_INFO_VENDOR_PERCENT_ROW, targetCell.Column), _
               .Cells(BASIC_INFO_VENDOR_TOTAL_ROW, targetCell.Column)).ClearContents
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

Public Function GetVendorIndexFromValueColumnPublic(ByVal valueColumn As Long) As Long
    GetVendorIndexFromValueColumnPublic = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(valueColumn)
End Function

' F9 件数内の施工会社名セル（F11/I11/L11...）を target から特定する。
' Union レンジとの Intersect が不安定な場合でも列位置で判定する。

Public Function GetVendorNameRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim i As Long
    Dim result As Range
    For i = 1 To vendorCount
        If result Is Nothing Then
            Set result = mod_VendorBlockLayout.VendorNameCellByIndex(wsInfo, i)
        Else
            Set result = Union(result, mod_VendorBlockLayout.VendorNameCellByIndex(wsInfo, i))
        End If
    Next i

    Set GetVendorNameRange = result
End Function

Public Function ResolveVendorNameChangeCell(ByVal wsInfo As Worksheet, ByVal target As Range) As Range
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim hitCell As Range
    For Each hitCell In target.Cells
        If hitCell.Row = BASIC_INFO_VENDOR_NAME_ROW Then
            Dim vendorIndex As Long
            vendorIndex = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(hitCell.Column)
            If vendorIndex >= 1 And vendorIndex <= vendorCount Then
                Set ResolveVendorNameChangeCell = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, hitCell.Column)
                Exit Function
            End If
        End If
    Next hitCell
End Function

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

    mod_VendorUnitPrice.RefreshVendorUnitPriceForValueColumn wsInfo, targetCell.Column
    NotifyVendorBasicInfoBlockChanged wsInfo, targetCell.Column
    mod_DebugLog.Log "[VendorMaster] ApplyVendorSelection: no match in vendor master. Branch=[" & _
                      BranchName & "] Vendor=[" & vendorName & "] Col=" & targetCell.Column & _
                      " RowCount=" & vendorRows.Count & " -> row12+ left blank"
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

Public Sub ClearVendorRowsCache()
    Set mVendorRowsCache = Nothing
    Set mVendorNameIndexCache = Nothing
    mod_Construction_Order_Import.ClearVendorAliasMapCache
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
        mod_VendorUnitPrice.RefreshVendorUnitPriceForValueColumn wsInfo, targetCell.Column
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

Public Sub HideVendorComboBox(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    DeleteVendorComboBox wsInfo
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
            ' 会社名のみ変更(工事区分は不変)でも、当該列の溶接単価/軌道会社の溶接手元単価が
            ' 未展開のまま残ることがあるため、対象列だけ単価展開を保証する。
            ' (溶接数式は手元比率=整理番号キー・外注比率=絶対参照で業者非依存。既展開なら結果は等価)
            Dim weldingNameChangeCols As Collection
            Set weldingNameChangeCols = New Collection
            weldingNameChangeCols.Add valueColumn
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfoColumns wsInfo, weldingNameChangeCols, valueColumn
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
    vendorIndex = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(valueColumn)
    If vendorIndex > 0 Then
        mod_BasicInfoGuide.RefreshSingleVendorRowGuidePublic wsInfo, vendorIndex
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals vendorIndex
    End If
    Exit Sub

ErrorHandler:
    Err.Clear
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

Public Sub ShowAllVendorSelection(Optional ByVal wsInfo As Worksheet, Optional ByVal targetCell As Range)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Set targetCell = GetVendorTargetCell(wsInfo, targetCell)
    If targetCell Is Nothing Then Exit Sub
    mVendorTargetAddress = targetCell.Address(False, False)

    DeleteVendorComboBox wsInfo
    AllVenderSelection.Show vbModal
End Sub

Public Sub ApplyConstructionUnitPriceImportedRowDecorations(ByVal wsUnitPrice As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Call mod_VendorUnitPrice.ApplyConstructionUnitPriceImportedRowDecorations(wsUnitPrice, firstRow, lastRow)
End Sub

Public Sub ApplyImportedUnitPriceJrHeadersForBasicInfo(Optional ByVal wsInfo As Worksheet)
    Call mod_VendorUnitPrice.ApplyImportedUnitPriceJrHeadersForBasicInfo(wsInfo)
End Sub

Public Function BuildVendorUnitPriceNameMap(ByVal wsInfo As Worksheet) As Object
    Set BuildVendorUnitPriceNameMap = mod_VendorUnitPrice.BuildVendorUnitPriceNameMap(wsInfo)
End Function

Public Sub CleanupLegacyVendorListDebrisInColumnAD(Optional ByVal wsInfo As Worksheet)
    Call mod_VendorBlockLayout.CleanupLegacyVendorListDebrisInColumnAD(wsInfo)
End Sub

Public Sub EnsureVendorCountInputValidation(Optional ByVal wsInfo As Worksheet)
    Call mod_VendorBlockLayout.EnsureVendorCountInputValidation(wsInfo)
End Sub

Public Function GetVendorUnitPriceMonitorRange(ByVal wsInfo As Worksheet) As Range
    Set GetVendorUnitPriceMonitorRange = mod_VendorUnitPrice.GetVendorUnitPriceMonitorRange(wsInfo)
End Function

Public Sub HandleConstructionUnitPriceSheetChange(ByVal wsUnitPrice As Worksheet, ByVal changedRange As Range)
    Call mod_VendorUnitPrice.HandleConstructionUnitPriceSheetChange(wsUnitPrice, changedRange)
End Sub

Public Sub HandleVendorUnitPriceMonitorChange(ByVal wsInfo As Worksheet, ByVal changedRange As Range)
    Call mod_VendorUnitPrice.HandleVendorUnitPriceMonitorChange(wsInfo, changedRange)
End Sub

Public Sub InitVendorBlockCountFromSheet(Optional ByVal wsInfo As Worksheet)
    Call mod_VendorBlockLayout.InitVendorBlockCountFromSheet(wsInfo)
End Sub

Public Function IsSyncVendorBlocksInProgress() As Boolean
    IsSyncVendorBlocksInProgress = mod_VendorBlockLayout.IsSyncVendorBlocksInProgressImpl()
End Function

Public Sub RefreshAllConstructionUnitPriceSheetDataDecorations(Optional ByVal wsInfo As Worksheet)
    Call mod_VendorUnitPrice.RefreshAllConstructionUnitPriceSheetDataDecorations(wsInfo)
End Sub

Public Sub RefreshAllVendorUnitPricesForBasicInfo(Optional ByVal wsInfo As Worksheet, Optional ByVal deferCalculation As Boolean = False)
    Call mod_VendorUnitPrice.RefreshAllVendorUnitPricesForBasicInfo(wsInfo, deferCalculation)
End Sub

Public Sub RefreshConstructionUnitPriceSheetDataDecorations(ByVal wsUnitPrice As Worksheet, ByVal wsInfo As Worksheet, Optional ByVal skipVendorColumnRefresh As Boolean = False)
    Call mod_VendorUnitPrice.RefreshConstructionUnitPriceSheetDataDecorations(wsUnitPrice, wsInfo, skipVendorColumnRefresh)
End Sub

Public Sub ResetVendorBlockSyncState()
    mod_VendorBlockLayout.ResetVendorBlockSyncStateImpl
End Sub

Public Sub SyncVendorBlocksFromCount(ByVal wsInfo As Worksheet)
    Call mod_VendorBlockLayout.SyncVendorBlocksFromCount(wsInfo)
End Sub
