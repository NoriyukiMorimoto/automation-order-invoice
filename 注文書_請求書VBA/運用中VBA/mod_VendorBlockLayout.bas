Option Explicit

Private Const VENDOR_LIST_START_ROW As Long = 2
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_LABEL_COL As Long = 5
Private Const BASIC_INFO_VENDOR_BLOCK_VALUE_COL As Long = 6
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_BLOCK_BOTTOM_ROW As Long = 31
Private Const BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW As Long = 32
Private Const BASIC_INFO_VENDOR_TOTAL_ROW As Long = 33
Private Const BASIC_INFO_OTHER_INPUT_TOP_ROW As Long = 37
Private Const BASIC_INFO_OTHER_INPUT_BOTTOM_ROW As Long = 42
Private Const BASIC_INFO_VENDOR_BLOCK_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_LABEL_COL_WIDTH As Double = 26.38
Private Const BASIC_INFO_VENDOR_VALUE_COL_WIDTH As Double = 42.5
Private Const BASIC_INFO_VENDOR_SPACER_COL_WIDTH As Double = 0.92
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const MAX_VENDOR_BLOCK_COUNT As Long = 10
Private Const OTHER_INPUT_BASE_ROW_HEIGHT As Double = 24#

Private mLastVendorBlockCount As Long
Private mSyncVendorBlocksInProgress As Boolean

Public Function CountExistingVendorBlocks(ByVal wsInfo As Worksheet) As Long
    Dim i As Long
    For i = 1 To MAX_VENDOR_BLOCK_COUNT
        If InStr(1, CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(i)).value), _
                 VendorInfoHeaderPrefixText(), vbTextCompare) = 0 Then
            Exit For
        End If
        CountExistingVendorBlocks = i
    Next i
End Function

Public Function GetVendorBlockCount(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CStr(wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > MAX_VENDOR_BLOCK_COUNT Then countValue = MAX_VENDOR_BLOCK_COUNT
    GetVendorBlockCount = countValue
End Function

' F9(�{�H��А�)�ɑz��O�̒l(65�Ȃ�)�����͂�����SyncVendorBlocksFromCount��
' ��ʃu���b�N�̈ꊇ���������݂Ē����ԉ����Ȃ��ɂȂ邽�߁A1�`MAX_VENDOR_BLOCK_COUNT��
' �����̂݋�������͋K�������炩���ߐݒ肵�Ă����BWorksheet_Activate����Ăяo���z��B

Public Function VendorBlockNeedsPresentationRestore(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Boolean
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

Public Function VendorBlockRangeByIndex(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Range
    Set VendorBlockRangeByIndex = wsInfo.Range( _
        wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorLabelColumnByIndex(vendorIndex)), _
        wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorSpacerColumnByIndex(vendorIndex)))
End Function

Public Function VendorCountValidationErrorText() As String
    VendorCountValidationErrorText = "1" & ChrW$(&H301C) & CStr(MAX_VENDOR_BLOCK_COUNT) & _
                                     ChrW$(&H306E) & ChrW$(&H6574) & ChrW$(&H6570) & ChrW$(&H3092) & _
                                     ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & _
                                     ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & "."
End Function

' ��VENDOR_LIST_COL("AD")�Ɏ��c���ꂽ�ƎҖ��h���b�v�_�E���ꗗ�̎c�[����������ꎞ�Ή��B
' AD���9�Ж�(�Ǝҏ��-9)�̒l��ƌ��p����Ă������߁A�ꗗ���������܂�Ă���
' VENDOR_LIST_START_ROW�`�Ǝ҃u���b�N���[(�_����z���v�s)�܂ł�ΏۂɃN���A���A
' 9�Жڃu���b�N�𖢓��͏��(�ΐ��K�C�h)�֐��������B��x���s����Ώ\���ŁA
' ����͏�L��VENDOR_LIST_COL�ύX(AJ��)�ɂ��AD��Ɉꗗ���������܂�邱�Ƃ͂Ȃ��B

Public Function VendorCountValidationInputText() As String
    VendorCountValidationInputText = "1" & ChrW$(&H301C) & CStr(MAX_VENDOR_BLOCK_COUNT) & _
                                     ChrW$(&H306E) & ChrW$(&H6574) & ChrW$(&H6570) & ChrW$(&H3092) & _
                                     ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & _
                                     ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & "."
End Function

Public Function VendorCountValidationTitleText() As String
    VendorCountValidationTitleText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H6570)
End Function

Public Function VendorInfoHeaderPrefixText() As String
    VendorInfoHeaderPrefixText = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H60C5) & ChrW$(&H5831) & "-"
End Function

Public Function VendorInfoHeaderText(ByVal vendorIndex As Long) As String
    VendorInfoHeaderText = "  " & ChrW$(&H25B8) & "  " & VendorInfoHeaderPrefixText() & CStr(vendorIndex)
End Function

Public Function VendorLabelColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorLabelColumnByIndex = BASIC_INFO_VENDOR_BLOCK_LABEL_COL + ((vendorIndex - 1) * BASIC_INFO_VENDOR_BLOCK_STEP_COLS)
End Function

Public Function VendorNameCellByIndex(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Range
    Set VendorNameCellByIndex = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, VendorValueColumnByIndex(vendorIndex))
End Function

Public Function VendorSpacerColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorSpacerColumnByIndex = VendorLabelColumnByIndex(vendorIndex) + 2
End Function

Public Function VendorValueColumnByIndex(ByVal vendorIndex As Long) As Long
    VendorValueColumnByIndex = VendorLabelColumnByIndex(vendorIndex) + 1
End Function

Public Sub ApplyVendorBlockColumnWidths(ByVal wsInfo As Worksheet, ByVal vendorCount As Long)
    Dim i As Long
    For i = 1 To vendorCount
        wsInfo.Columns(VendorLabelColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_LABEL_COL_WIDTH
        wsInfo.Columns(VendorValueColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_VALUE_COL_WIDTH
        wsInfo.Columns(VendorSpacerColumnByIndex(i)).ColumnWidth = BASIC_INFO_VENDOR_SPACER_COL_WIDTH
    Next i
End Sub

Public Sub ApplyVendorRow10ValueCellFormat(ByVal valueCell As Range)
    If valueCell Is Nothing Then Exit Sub

    Dim vendorIndex As Long
    vendorIndex = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(valueCell.Column)
    If vendorIndex <= 0 Then Exit Sub

    mod_VendorInfoColors.ApplyVendorInfoRow10Color valueCell.Worksheet, vendorIndex
End Sub

Public Sub ClearUnusedVendorBlocks(ByVal wsInfo As Worksheet, ByVal firstUnusedIndex As Long)
    Dim i As Long
    For i = firstUnusedIndex To MAX_VENDOR_BLOCK_COUNT
        ClearVendorBlockColumns wsInfo, i
    Next i
End Sub

Public Sub ClearVendorBlockColumns(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
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

    ' ���̑����͎���(37-42�s)�����g�p�u���b�N���珜������
    ClearOtherInputBlockColumns wsInfo, vendorIndex
End Sub

Public Sub ClearOtherInputBlockColumns(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If vendorIndex < 1 Or vendorIndex > MAX_VENDOR_BLOCK_COUNT Then Exit Sub

    Dim labelCol As Long
    labelCol = VendorLabelColumnByIndex(vendorIndex)

    Dim clearRange As Range
    Set clearRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_OTHER_INPUT_TOP_ROW, labelCol), _
                                  wsInfo.Cells(BASIC_INFO_OTHER_INPUT_BOTTOM_ROW, VendorSpacerColumnByIndex(vendorIndex)))

    SafeUnmergeRange clearRange
    clearRange.ClearContents
    clearRange.Interior.Color = RGB(6, 17, 29)
    clearRange.Borders.LineStyle = xlNone

    ' 未使用ブロックの左隣(直前ブロックの右端)に残る罫線が、未入力の38-42行に線として
    ' 見えてしまうため除去する(例: 2社時の K38-K42 の左罫線 = 2社目スペーサ列の右罫線)
    If labelCol > VendorLabelColumnByIndex(1) Then
        wsInfo.Range(wsInfo.Cells(38, labelCol - 1), _
                     wsInfo.Cells(BASIC_INFO_OTHER_INPUT_BOTTOM_ROW, labelCol - 1)).Borders(xlEdgeRight).LineStyle = xlNone
    End If
End Sub

Public Sub ClearVendorWorkTypeWhenCompanyEmpty(ByVal wsInfo As Worksheet, Optional ByVal vendorCount As Long = 0)
    If wsInfo Is Nothing Then Exit Sub
    If vendorCount <= 0 Then vendorCount = GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueCol As Long
        valueCol = VendorValueColumnByIndex(vendorIndex)
        If Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).value))) = 0 Then
            mod_VendorMaster.VendorWritableValueCell(wsInfo, BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol).ClearContents
        End If
    Next vendorIndex
End Sub

Public Sub CopyCellBorderEdge(ByVal sourceCell As Range, ByVal destCell As Range, ByVal edgeId As Long)
    On Error Resume Next
    With destCell.Borders(edgeId)
        .LineStyle = sourceCell.Borders(edgeId).LineStyle
        .Weight = sourceCell.Borders(edgeId).Weight
        .Color = sourceCell.Borders(edgeId).Color
    End With
    On Error GoTo 0
End Sub

Public Sub CopyRangeBorders(ByVal sourceRange As Range, ByVal destRange As Range)
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

Public Sub CopyRangeFormats(ByVal sourceRange As Range, ByVal destRange As Range)
    If sourceRange Is Nothing Then Exit Sub
    If destRange Is Nothing Then Exit Sub
    If sourceRange.Rows.Count <> destRange.Rows.Count Then Exit Sub
    If sourceRange.Columns.Count <> destRange.Columns.Count Then Exit Sub

    sourceRange.Copy
    destRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Public Sub CopyVendorBlockFormatsFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
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

Public Sub CopyVendorBlockFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
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

    mod_VendorMaster.ClearVendorInfoBlock VendorNameCellByIndex(wsInfo, destVendorIndex)

    wsInfo.Cells(BASIC_INFO_VENDOR_PURCHASE_TOTAL_ROW, VendorValueColumnByIndex(destVendorIndex)).ClearContents
    wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, VendorValueColumnByIndex(destVendorIndex)).ClearContents
End Sub

Public Sub CopyVendorBlockMergeAreasFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
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

Public Sub EnsureVendorBlockFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    CopyVendorBlockFromTemplate wsInfo, destVendorIndex
    RestoreVendorBlockPresentationFromTemplate wsInfo, destVendorIndex
    ' ���̑����͂����쐬�̂Ƃ������l�����ɂ���(�������͕͂ێ�)
    EnsureOtherInputBlockFromTemplate wsInfo, destVendorIndex, OtherInputBlockNeedsRestore(wsInfo, destVendorIndex)
End Sub

' ���̑����͎���(37-42�s): 1�Жڂ𐗌`�Ƀ��x���E����(�r��/�h�F/�t�H���g/�܂�Ԃ���)���R�s�[����
Public Function OtherInputBlockNeedsRestore(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If vendorIndex < 2 Then Exit Function
    OtherInputBlockNeedsRestore = _
        (Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_OTHER_INPUT_TOP_ROW, VendorLabelColumnByIndex(vendorIndex)).value))) = 0)
End Function

Public Sub EnsureOtherInputBlockFromTemplate(ByVal wsInfo As Worksheet, _
                                              ByVal destVendorIndex As Long, _
                                              Optional ByVal clearValues As Boolean = True)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim srcLabelCol As Long
    Dim dstLabelCol As Long
    Dim srcValueCol As Long
    Dim dstValueCol As Long
    Dim srcSpacerCol As Long
    Dim dstSpacerCol As Long
    srcLabelCol = VendorLabelColumnByIndex(1)
    dstLabelCol = VendorLabelColumnByIndex(destVendorIndex)
    srcValueCol = VendorValueColumnByIndex(1)
    dstValueCol = VendorValueColumnByIndex(destVendorIndex)
    srcSpacerCol = VendorSpacerColumnByIndex(1)
    dstSpacerCol = VendorSpacerColumnByIndex(destVendorIndex)

    Dim sourceRange As Range
    Dim destRange As Range
    ' ���x����`�X�y�[�T��(3��)��ΏۂɁA�r���E�h�F�E�t�H���g�E�܂�Ԃ������܂ޏ����𕡐�����
    Set sourceRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_OTHER_INPUT_TOP_ROW, srcLabelCol), _
                                   wsInfo.Cells(BASIC_INFO_OTHER_INPUT_BOTTOM_ROW, srcSpacerCol))
    Set destRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_OTHER_INPUT_TOP_ROW, dstLabelCol), _
                                 wsInfo.Cells(BASIC_INFO_OTHER_INPUT_BOTTOM_ROW, dstSpacerCol))

    ' �������͒l��ޔ�(�����ēK�p���ɕێ�����)
    Dim savedValues(BASIC_INFO_OTHER_INPUT_TOP_ROW To BASIC_INFO_OTHER_INPUT_BOTTOM_ROW) As Variant
    Dim rowIndex As Long
    If Not clearValues Then
        For rowIndex = BASIC_INFO_OTHER_INPUT_TOP_ROW To BASIC_INFO_OTHER_INPUT_BOTTOM_ROW
            savedValues(rowIndex) = wsInfo.Cells(rowIndex, dstValueCol).value
        Next rowIndex
    End If

    SafeUnmergeRange destRange

    On Error Resume Next
    sourceRange.Copy
    destRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    CopyRangeBorders sourceRange, destRange
    CopyOtherInputMergeAreasFromTemplate wsInfo, destVendorIndex

    For rowIndex = BASIC_INFO_OTHER_INPUT_TOP_ROW To BASIC_INFO_OTHER_INPUT_BOTTOM_ROW
        wsInfo.Cells(rowIndex, dstLabelCol).value = wsInfo.Cells(rowIndex, srcLabelCol).value

        If clearValues Then
            ' �l��͉�Ђ��Ƃɓ��͂��邽�ߐ��`�̒l�̓R�s�[���Ȃ�(������PasteSpecial�ňێ�)
            wsInfo.Cells(rowIndex, dstValueCol).ClearContents
        Else
            wsInfo.Cells(rowIndex, dstValueCol).value = savedValues(rowIndex)
        End If
    Next rowIndex
End Sub

' ���̑����͎����u���b�N�̌����Z����1�Жڂ��畡������
Public Sub CopyOtherInputMergeAreasFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If destVendorIndex < 2 Then Exit Sub

    Dim srcLabelCol As Long
    Dim srcSpacerCol As Long
    srcLabelCol = VendorLabelColumnByIndex(1)
    srcSpacerCol = VendorSpacerColumnByIndex(1)

    Dim srcRange As Range
    Set srcRange = wsInfo.Range(wsInfo.Cells(BASIC_INFO_OTHER_INPUT_TOP_ROW, srcLabelCol), _
                                wsInfo.Cells(BASIC_INFO_OTHER_INPUT_BOTTOM_ROW, srcSpacerCol))

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
                Set destMerge = wsInfo.Range( _
                    wsInfo.Cells(mergeArea.Row, mergeArea.Column + colOffset), _
                    wsInfo.Cells(mergeArea.Row + mergeArea.Rows.Count - 1, _
                                 mergeArea.Column + colOffset + mergeArea.Columns.Count - 1))
                SafeUnmergeRange destMerge
                destMerge.Merge
            End If
        End If
    Next cell
    On Error GoTo 0
End Sub

Public Sub RefreshOtherInputBlocks(ByVal wsInfo As Worksheet, Optional ByVal vendorCount As Long = 0)
    If wsInfo Is Nothing Then Exit Sub
    If vendorCount <= 0 Then vendorCount = GetVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 2 To vendorCount
        ' �����u���b�N�������𐗌`�֑�����B�l�͖��쐬���̂݃N���A�B
        EnsureOtherInputBlockFromTemplate wsInfo, i, OtherInputBlockNeedsRestore(wsInfo, i)
    Next i

    For i = vendorCount + 1 To MAX_VENDOR_BLOCK_COUNT
        ClearOtherInputBlockColumns wsInfo, i
    Next i
End Sub

Public Sub RestoreVendorBlockLabelTextsFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
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

Public Sub RestoreVendorBlockPresentationFromTemplate(ByVal wsInfo As Worksheet, ByVal destVendorIndex As Long)
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

Public Sub RestoreVendorBlockValueColumnRightBorders(ByVal wsInfo As Worksheet, ByVal vendorCount As Long)
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

Public Sub SafeUnmergeRange(ByVal targetRange As Range)
    If targetRange Is Nothing Then Exit Sub

    On Error Resume Next
    Dim cell As Range
    For Each cell In targetRange.Cells
        If cell.MergeCells Then cell.MergeArea.UnMerge
    Next cell
    On Error GoTo 0
End Sub

Public Function IsSyncVendorBlocksInProgressImpl() As Boolean
    IsSyncVendorBlocksInProgressImpl = mSyncVendorBlocksInProgress
End Function

Public Sub CleanupLegacyVendorListDebrisInColumnAD(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Const LEGACY_LIST_COL As String = "AD"

    On Error Resume Next
    Dim clearRange As Range
    Set clearRange = wsInfo.Range(LEGACY_LIST_COL & VENDOR_LIST_START_ROW & ":" & _
                                  LEGACY_LIST_COL & CStr(BASIC_INFO_VENDOR_TOTAL_ROW))
    SafeUnmergeRange clearRange
    clearRange.ClearContents
    On Error GoTo 0

    ' 9�Жڃu���b�N(�l��AD=30���)�̌����ڂ𖢓��͏�Ԃ֐�������
    Dim vendorIndex9 As Long
    vendorIndex9 = mod_VendorUnitPrice.GetVendorIndexFromValueColumn(30)
    If vendorIndex9 >= 1 Then
        mod_BasicInfoGuide.RefreshSingleVendorRowGuidePublic wsInfo, vendorIndex9
    End If
End Sub

Public Sub EnsureVendorCountInputValidation(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error Resume Next
    With wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="1", Formula2:=CStr(MAX_VENDOR_BLOCK_COUNT)
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = VendorCountValidationTitleText()
        .InputMessage = VendorCountValidationInputText()
        .ErrorTitle = VendorCountValidationTitleText()
        .ErrorMessage = VendorCountValidationErrorText()
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Public Sub InitVendorBlockCountFromSheet(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    ResetVendorBlockSyncState

    Dim vendorCount As Long
    vendorCount = GetVendorBlockCount(wsInfo)

    Dim existingBlockCount As Long
    existingBlockCount = CountExistingVendorBlocks(wsInfo)

    ' F9�����ɑ΂��Ǝҏ��u���b�N���s�����Ă���Ƃ������t����������B
    ' ���� SyncVendorBlocksFromCount ���ĂԂƗn�ڒP���W�J���� Activate ���ł܂�B
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

Public Sub ResetVendorBlockSyncStateImpl()
    mSyncVendorBlocksInProgress = False
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

    ' ���̑����͎���(37-42�s)���{�H��А��������p��/��������
    RefreshOtherInputBlocks wsInfo, vendorCount

    ClearVendorWorkTypeWhenCompanyEmpty wsInfo, vendorCount

    Dim formatIndex As Long
    For formatIndex = 1 To vendorCount
        ApplyVendorRow10ValueCellFormat wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, VendorValueColumnByIndex(formatIndex))
    Next formatIndex

    Application.CutCopyMode = False
    mod_VendorMaster.RefreshVendorListForBasicInfo wsInfo
    mod_VendorUnitPrice.SyncVendorUnitPriceBlocksAfterCountChange wsInfo, vendorCount, previousCount, True

    If vendorCount > previousCount Then
        ' �ǉ��u���b�N�͋�̂��ߗn�ڒP���̍ēW�J�͕s�v
    ElseIf vendorCount < previousCount Then
        mod_WeldingUnitPrice.RefreshWeldingAfterVendorCountDecrease wsInfo, vendorCount, _
            oldWeldingBlockCount, oldRailBlockCount, True
    ElseIf vendorBlocksEnsured Then
        mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, 0, True
    End If
    mLastVendorBlockCount = vendorCount

    ' ��А����ς�����u���b�N��32/33�s��(�_����z���v)���e���v���[�g�R���̋󗓂̂܂�
    ' �c�邽�߁A�����Ŗ����I�ɍ��v���Čv�Z�E�����݂���(�V�K�ǉ��u���b�N��0�~���������܂�)�B
    If vendorCount <> previousCount Then
        On Error Resume Next
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
        On Error GoTo 0
    End If

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
        mod_OrderTpl_Shared.OrderTplRepairAllGeneratedPlaceholderFormulas True
        On Error Resume Next
        Application.Calculate
        On Error GoTo ExitHandler
    End If

    ' ����I������ ExitHandler �֗��Ƃ��A�������t���O�̉����Ə�ԕ�����K���s���B
    ' (�������� Exit Sub �Ńt���O���������ꂸ�A����̉�Ж��I�𓙂̕ύX�C�x���g��
    '  Sheet1 ���̃K�[�h�Ŗ�������A�V�[�g�ăA�N�e�B�u�܂œ]�L����Ȃ��s���������)

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
