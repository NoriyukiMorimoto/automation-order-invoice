Option Explicit

' Vendor info colors: fixed palette for vendor blocks 1-10 (?????).
' Used for basic-info row 10 headers and future template sheet tab colors.

Private Const VENDOR_INFO_COLOR_COUNT As Long = 10
Private Const VENDOR_INFO_ROW As Long = 10
Private Const VENDOR_BLOCK_LABEL_COL As Long = 5
Private Const VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const MAX_VENDOR_BLOCK_COUNT As Long = 10
Private Const VENDOR_SPACER_FILL_COLOR_R As Long = 6
Private Const VENDOR_SPACER_FILL_COLOR_G As Long = 17
Private Const VENDOR_SPACER_FILL_COLOR_B As Long = 29

Public Function GetVendorInfoColorBackground(ByVal vendorIndex As Long) As Long
    GetVendorInfoColorBackground = VendorInfoColorBackgroundValue(NormalizeVendorColorIndex(vendorIndex))
End Function

Public Function GetVendorInfoColorForeground(ByVal vendorIndex As Long) As Long
    GetVendorInfoColorForeground = VendorInfoColorForegroundValue(NormalizeVendorColorIndex(vendorIndex))
End Function

Public Sub ApplyOutputSheetVendorCellColor(ByVal ws As Worksheet, _
                                           ByVal rowIndex As Long, _
                                           ByVal vendorColumn As Long, _
                                           Optional ByVal workTypeKeyword As String = "")
    If ws Is Nothing Then Exit Sub
    If rowIndex < 2 Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Cells(rowIndex, vendorColumn)

    Dim vendorName As String
    vendorName = Trim$(CommonNzText(targetCell.value))
    If vendorName = "" Then
        RestoreOutputSheetVendorCellDefault targetCell
        Exit Sub
    End If

    Dim vendorIndex As Long
    vendorIndex = mod_Construction_Order_Import.ResolveBasicInfoVendorInfoIndex( _
        vendorName, workTypeKeyword)
    If vendorIndex <= 0 Then
        RestoreOutputSheetVendorCellDefault targetCell
        Exit Sub
    End If

    With targetCell
        .Interior.Color = GetVendorInfoColorBackground(vendorIndex)
        .Font.Color = GetVendorInfoColorForeground(vendorIndex)
    End With
End Sub

Public Sub ApplyVendorInfoRow10Colors(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        ApplyVendorInfoRow10Color wsInfo, vendorIndex
    Next vendorIndex
End Sub

Public Sub ApplyVendorInfoRow10Color(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If vendorIndex < 1 Then Exit Sub

    Dim labelCol As Long
    Dim valueCol As Long
    Dim spacerCol As Long
    labelCol = VendorLabelColumnByIndex(vendorIndex)
    valueCol = labelCol + 1
    spacerCol = labelCol + 2

    Dim colorRange As Range
    Set colorRange = wsInfo.Range(wsInfo.Cells(VENDOR_INFO_ROW, labelCol), _
                                  wsInfo.Cells(VENDOR_INFO_ROW, valueCol))

    Dim normalizedIndex As Long
    normalizedIndex = NormalizeVendorColorIndex(vendorIndex)

    With colorRange
        .Interior.Color = VendorInfoColorBackgroundValue(normalizedIndex)
        .Font.Color = VendorInfoColorForegroundValue(normalizedIndex)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    RestoreVendorSpacerRow10Cell wsInfo, spacerCol
End Sub

Private Sub RestoreVendorSpacerRow10Cell(ByVal wsInfo As Worksheet, ByVal spacerCol As Long)
    With wsInfo.Cells(VENDOR_INFO_ROW, spacerCol)
        .Interior.Color = RGB(VENDOR_SPACER_FILL_COLOR_R, _
                              VENDOR_SPACER_FILL_COLOR_G, _
                              VENDOR_SPACER_FILL_COLOR_B)
    End With
End Sub

Private Sub RestoreOutputSheetVendorCellDefault(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub

    With targetCell
        .Interior.Pattern = xlNone
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlAutomatic
    End With
End Sub

Private Function GetBasicInfoVendorBlockCount(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CStr(wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > MAX_VENDOR_BLOCK_COUNT Then countValue = MAX_VENDOR_BLOCK_COUNT
    GetBasicInfoVendorBlockCount = countValue
End Function

Private Function VendorLabelColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorLabelColumnByIndex = VENDOR_BLOCK_LABEL_COL + ((vendorIndex - 1) * VENDOR_BLOCK_STEP_COLS)
End Function

Private Function NormalizeVendorColorIndex(ByVal vendorIndex As Long) As Long
    Dim normalizedIndex As Long
    normalizedIndex = ((vendorIndex - 1) Mod VENDOR_INFO_COLOR_COUNT) + 1
    NormalizeVendorColorIndex = normalizedIndex
End Function

Private Function VendorInfoColorBackgroundValue(ByVal colorIndex As Long) As Long
    Select Case colorIndex
        Case 1
            VendorInfoColorBackgroundValue = RGB(70, 130, 180)
        Case 2
            VendorInfoColorBackgroundValue = RGB(255, 160, 180)
        Case 3
            VendorInfoColorBackgroundValue = RGB(0, 128, 128)
        Case 4
            VendorInfoColorBackgroundValue = RGB(186, 140, 255)
        Case 5
            VendorInfoColorBackgroundValue = RGB(178, 34, 34)
        Case 6
            VendorInfoColorBackgroundValue = RGB(255, 222, 173)
        Case 7
            VendorInfoColorBackgroundValue = RGB(72, 61, 139)
        Case 8
            VendorInfoColorBackgroundValue = RGB(143, 188, 143)
        Case 9
            VendorInfoColorBackgroundValue = RGB(148, 0, 211)
        Case 10
            VendorInfoColorBackgroundValue = RGB(210, 180, 140)
        Case Else
            VendorInfoColorBackgroundValue = RGB(70, 130, 180)
    End Select
End Function

Private Function VendorInfoColorForegroundValue(ByVal colorIndex As Long) As Long
    Select Case colorIndex
        Case 2, 4, 6, 8, 10
            VendorInfoColorForegroundValue = RGB(0, 0, 0)
        Case Else
            VendorInfoColorForegroundValue = RGB(255, 255, 255)
    End Select
End Function
