Option Explicit

' ?��?��?��?��?��?��?��e?��?��?��v?��?��?��[?��g?��e?��V?��[?��g?��ւ̊�{?��?��?��w?��b?��_?��[?��]?��L?��B
' ?��?��?��?��?��?��?��ɉ�?��?��?��A?��?��{?��?��?��V?��[?��g?��̓]?��L?��?��?��Z?��?��?��ύX?��?��?��ɂ�?��ē]?��L?��?��?��?��?��(?��?��?��C?��u?��?��?��f)?��B
' ?��]?��L?��ΏۃV?��[?��g?��̒ǉ�?��?�� ApplyVendorSheetHeaders ?��̃f?��B?��X?��p?��b?��`?��֎�?��?��?��?��?��?��?��?��?��?��?��ށB
' ?��?��?��C?��?��?��?��: CHANGELOG.md ?��Q?��?��

' ?��?��{?��?��?��̃w?��b?��_?��[?��]?��L?��?��?��Z?��?��(?��S?��Ћ�?��ʕ�)?��B?��u?��?��?��b?��N?��?��(16/27?��s?��?��)?��͓�?��I?��ɑg?��ݗ�?��Ă�
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' ?��?��җp?��V?��[?��g?��]?��L?��ŎQ?��Ƃ�?��?��{?��H?��?��Ѓu?��?��?��b?��N?��s(?��?��{?��?��?��)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35
' ?��?��{?��?��?�� ?��{?��H?��?��Ѓu?��?��?��b?��N: ?��?��\?��Җ�(12?��s?��?��)?��E?��Z?��?��(14?��s?��?��)
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
' �`�F�b�N�{�b�N�X��TopLeftCell.Row�̓e���v���[�g�̃A���J�[�ʒu�Ɉˑ����A
' Excel�̋��E�ۂ߂ɂ��35�`40 �܂��� 36�`41 �̂ǂ��炩�ɂȂ肤��B
' �ǂ���ł����p�F������悤�ɔ͈͂�35�`41�Ɋg���Ĕ��肷��B
Private Const ATTACHMENT3_CHECK_ROW_MIN As Long = 35
Private Const ATTACHMENT3_CHECK_ROW_MAX As Long = 41
Private Const ATTACHMENT3_COL_F As Long = 6
Private Const ATTACHMENT3_COL_H As Long = 8
Private Const ATTACHMENT3_COL_J As Long = 10
Private Const ATTACHMENT3_COL_L As Long = 12
Public Const ATTACHMENT3_SANPAI_ROW_MIN As Long = 49
Public Const ATTACHMENT3_SANPAI_ROW_MAX As Long = 52
' �{�ݖ��Z���͌������Z�� D49:I49(D��=4)�A���ݒn�� J49:O49(J��=10)�B
' �_�u���N���b�N���茋���̍�����(D��)�Ɉ�v������K�v�����邽�߁A�{�ݖ���=4�B
Public Const ATTACHMENT3_SANPAI_NAME_COL As Long = 4
Public Const ATTACHMENT3_SANPAI_VALUE_COL As Long = 10
Private Const SANPAI_FACILITY_MASTER_START_ROW As Long = 59

' ?��w?��?��u?��?��?��b?��N?��̎{?��H?��?��ЂɑΉ�?��?��?��?��e?��?��?��v?��?��?��[?��g5?��V?��[?��g?��փw?��b?��_?��[?��?��]?��L?��?��?��?��(?��f?��B?��X?��p?��b?��`?��?��)
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

' ?��S?��m?��?��?��Ђ̃e?��?��?��v?��?��?��[?��g?��V?��[?��g?��փw?��b?��_?��[?��?��?��ē]?��L?��?��?��?��
Public Sub RefreshAllVendorSheetHeaders(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
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
                    ApplyConditionSheetHeader wsInfo, ThisWorkbook.Worksheets(condSheetName), vendorIndex
                    SetupConditionCheckboxExclusivity ThisWorkbook.Worksheets(condSheetName)
                End If
            End If
        End If
    Next vendorIndex
    Exit Sub

Quiet:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorSheetHeaders error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Sheet1(?��?��{?��?��?��)?��?��Worksheet_Change?��?��?��?��Ă΂�?��?��?��?��?��B
' ?��w?��b?��_?��[?��]?��L?��?��?��Z?��?��(B6/C6/C2/C9/C10/C15:C16/F6?��A?��e?��u?��?��?��b?��N?��?��16/27?��s?��?��)?��̕ύX?��?��?��e?��ЃV?��[?��g?��֔�?��f?��?��?��?��
Public Sub HandleBasicInfoHeaderSourceChange(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim sourceRange As Range
    Set sourceRange = BuildHeaderSourceRange(wsInfo)
    If sourceRange Is Nothing Then Exit Sub
    If Intersect(target, sourceRange) Is Nothing Then Exit Sub

    RefreshAllVendorSheetHeaders wsInfo
    Exit Sub

Quiet:
    Err.Clear
End Sub

' ?��w?��b?��_?��[?��]?��L?��?��?��Z?��?��?��̊Ď�?��͈͂�Ԃ�?��?��?��J?��?��?��b?��p?��[(Sheet1?��̕ύX?��Q?��[?��g?��\?��z?��p)
Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' ?��w?��b?��_?��[?��]?��L?��?��?��Z?��?��?��̊Ď�?��͈�(?��?��?��ʃZ?��?�� + ?��e?��u?��?��?��b?��N?��̋Ǝ҃R?��[?��h16?��s?��?��/?��?��?��?��?��ԍ�27?��s?��?��)
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

' ?��V?��[?��g?��?��?��o?��?��(?��^?��u)?��F: ?��H?��?��?��敪(?��?��{?��?��?��10?��s?��?��)?��Z?��?��?��̓h?��?��Ԃ�?��F?��?��K?��p?��?��?��?��
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


' ?��?��?��?��׃w?��b?��_?��[?��?��?��֊�{?��?��?��V?��[?��g?��̓�?��e?��?��]?��L?��?��?��?��
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: ?��?��?��X?��R?��[?��h(?��o?��?��?��?��?��?��_?��P?��?��?��K?��p?��?��?��?��̒P?��?��?��K?��p?��?��?��?��V?��[?��g B/C?��?��ƍ� ?��?�� G?��?��)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: ?��?��?��?��?��ԍ�(?��{?��H?��?��Ѓu?��?��?��b?��N27?��s?��?��)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: ?��H?��?��?��ԍ�(?��?��{?��?��?��C9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' D6: ?��H?��?��?��?��?��?��(?��?��{?��?��?��C10?��A?��?��?��?��?��Z?��?��?��̂�?��ߒ�?��?��?��?��?��?��)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: ?��H?��?�� ?��?��/?��?��(?��?��{?��?��?��C15/C16?��A?��a?��?��\?��?��)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: ?��?��?��?��(?��?��{?��?��?��C2?��A?��a?��?��\?��?��)
    ' P2:Q2 created date from C2. O2 keeps template label.
    WriteHeaderDate wsBreakdown.Range("P2"), wsInfo.Range("C2").value
    wsBreakdown.Range("P2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE

    ' O3: ?��?��?��?��?��?��(?��?��{?��?��?��F6)
    ' P3:Q3 branch chief from F6. O3 keeps template label.
    WriteHeaderValue wsBreakdown.Range("P3"), wsInfo.Range("F6").value, True

    ' P5: ?��O?��?��?��?��Ж�(?��{?��H?��?��Ѓu?��?��?��b?��N11?��s?��?��)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: ?��Ǝ҃R?��[?��h(?��{?��H?��?��Ѓu?��?��?��b?��N16?��s?��?��)
    WriteHeaderValue wsBreakdown.Range("P6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' ?��?��?��?��?��y?��[?��W?��?��?��?��p?��Ƀw?��b?��_?��[?��s(7:10?��s?��?��)?��?��?��^?��C?��g?��?��?��s?��ɐݒ肷?��?��
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

    ' �N���b�N���ꂽ�`�F�b�N�{�b�N�X�͕K���A�N�e�B�u�V�[�g��ɂ���B
    ' �����̏����V(���)�Ԃœ����`�F�b�N�{�b�N�X�����݂��邽�߁A�S�V�[�g����
    ' �擪��v�Ŏ��V�[�g�ɌН������̂�h���A�A�N�e�B�u�V�[�g�Ɍ��肷��B
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

' ?��?��җp/?��?��?��?��?��?��?��?��/?��x?��X?��T ?��?��?��?��: S1:U1 ?��?��?��?��?��ԍ�(27) / Q2:V2 ?��?��?��?��C2(?��?��?��?��) /
'   ?��s20-34(E20:?��H?��?��?��?��C10 E22:?��s?��?��?��{?��?��C13 G24:?��H?��?��?��?��C15(?��?��?��?��) G26:?��H?��?��?��?��C16(?��?��?��?��)?��A
'   Q22:?��ō�(35) Q23:?��Ŕ�(33) Q24:?��?��?��?��?��(34)?��AC30/H30/J30/M34/R34/F32/F33 ?��?�� Reapply?��n?��o?��R)?��?��
'   ApplyContractorStyleCommonFields ?��ŁA3?��V?��[?��g?��Ƃ�?��?��?��ꃍ�W?��b?��N?��?��K?��p?��?��?��?��B
' ?��?��җp?��V?��[?��g?��ŗL: E9:?��Ǝ҃R?��[?��h(16) A13:?��?��Ж�(11) M10:?��?��?��?��?��ҏZ?��?��
'   M11:?��?��S?��H?��Ɗ�?��?��?��?��?��+?��?��o?��?��?��?�� M12:?��?��E?��?��?��?��(?��o?��?��?��?��?��?��?��?��?��X?��g?��Q?��?��)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    ' E9: ?��Ǝ҃R?��[?��h(?��{?��H?��?��Ѓu?��?��?��b?��N16?��s), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: ?��?��Ж�(?��{?��H?��?��Ѓu?��?��?��b?��N11?��s), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' M10/M11/M12: ?��o?��?��?��?��?��?��?��?��?��X?��g?��Q?��?��(?��?��?��?��?��?�� ?��Z?��?��/?��?��?��?��/?��?��E?��?��?��?��)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    ' ?��s38-42(?��?��?��?��?��?��?��?��/?��J?��?��/?��x?��?��?��ޗ�/?��ݗ^?��i/?��?��?��T?��C?��N?��?��)?��?��?��?��{?��?��?���
    ' ?��?��җp?��E?��?��?��?��?��?��?��?��?��E?��x?��X?��T?��֍ē]?��L(Reapply?��?��?��őS?��Ă̐�?��?��?��σV?��[?��g?��֓K?��p)
    mod_BasicInfoExclusiveChoice.ReapplyExclusiveChoices wsInfo, vendorIndex
    mod_BasicInfoSupplyLoan.ReapplySupplyLoan wsInfo, vendorIndex

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?��?��җp/?��?��?��?��?��?��?��?��/?��x?��X?��T?��ŋ�?��ʂ̓]?��L(S1:U1 ?��?��?��?��?��ԍ� / Q2:V2 ?��?��?��?��(?��?��?��?��) /
' ?��s20-34: E20 ?��H?��?��?��?��?��EE22 ?��s?��?��?��{?��?��?��EG24/G26 ?��H?��?��?��?��?��?��(?��?��?��?��)?��EQ22-24 ?��?��?��?��׎Q?��?��)?��B
' ?��A3?��V?��[?��g?��Ƃ�?��Z?��?��?��\?��?��?��?��?��?��?��S?��?��v?��?��?��Ă�?��邽?��߁A?��?��?��ꃍ�W?��b?��N?��?��?��?��?��ꂼ?��?��K?��p?��?��?��?��B
Private Sub ApplyContractorStyleCommonFields(ByVal wsInfo As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal valueColumn As Long)
    ' S1:U1: ?��?��?��?��?��ԍ�(?��{?��H?��?��Ѓu?��?��?��b?��N27?��s), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2:V2: ?��?��?��?��(?��?��{?��?��?��C2, ?��?��?��?��), ?��?��?��?��
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E20: ?��H?��?��?��?��(?��?��{?��?��?��C10), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: ?��s?��?��?��{?��?��(?��?��{?��?��?��C13), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: ?��H?��?�� ?��?��(?��?��{?��?��?��C15, ?��?��?��?��), ?��?��?��?��
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: ?��H?��?�� ?��?��(?��?��{?��?��?��C16, ?��?��?��?��), ?��?��?��?��
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22/Q23/Q24: ?��?��?��?��O?��?��?��[?��v?��̓�?��?��׃V?��[?��g ?��?��?��v/?��v/?��?��?��?��?�� ?��s(Q?��?��)?��?��?��Q?��Ƃ�?��鐶?��?��?��?��?��?��
    ' (?��?��?��?��ׂ̒l?��?��?��X?��V?��?��?��?��?��Ǝ�?��?��?��ōX?��V?��?��?��?��?��)
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

' wsTarget(?��?��?��?��?��σe?��?��?��v?��?��?��[?��g?��V?��[?��g)?��Ɠ�?��?��O?��?��?��[?��v?��̎󒍎җp(?��?��?��?��)?��V?��[?��g?��?��?��?��?��?��?��?��?��?��
Private Function ResolveContractorSheetFromTarget(ByVal wsTarget As Worksheet) As Worksheet
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then Set ResolveContractorSheetFromTarget = ThisWorkbook.Worksheets(nm)
End Function

' ?��?��җp?��V?��[?��g?��̌�?��?��?��Z?��?��?��?��?��?��l?��?��?��擾?��?��?��?��(?��?��?��݂�?��Ȃ�/?��?��Ȃ�?���)
Private Function MirroredContractorText(ByVal wsContractor As Worksheet, ByVal address As String) As String
    If wsContractor Is Nothing Then Exit Function
    MirroredContractorText = CommonNzText(wsContractor.Range(address).MergeArea.Cells(1, 1).value)
End Function

' ?��?��җp M10:V ?��s?��̌�?��?��?��?��?��e?��?��?��v?��?��?��[?��g(M:U)?��?��?��?��?��?��S?��Ɋg?��?��?��?��?��?��
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

' M10:?��Z?��?�� / M11:?��?��S?��H?��Ɗ�?��?��?��?��?��+?��S?��p?��?��+?��?��o?��?��?��?�� / M12:(?��?��?��?��?��?��)?��?��E?��?��?��?�� ?��?��]?��L
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
    ' M10:U ?��?�� M10:V ?��֌�?��?��?��?��?��?��?��?��?��A?��k?��?��?��?��?��đS?��̕\?��?��?��ɐݒ�(M/13?��?�� ?��` V/22?��?��)
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

    ' M10: ?��Z?��?��
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: ?��?��S?��H?��Ɗ�?��?��?��?��?�� + ?��S?��p?��?�� + ?��?��o?��?��?��?��
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: ?��o?��?��?��?��=?��?��o?��?��?��?��?��Ȃ�u?��?��E ?��?��?��?��?��v?��A?��s?��?��v?��Ȃ�u?��o?��?��?��?�� ?��?��E ?��?��?��?��?��v
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValueRight wsTarget.Range("M12"), m12
End Sub

' ?��E?��l?��̒l?��]?��L(BizUD?��S?��V?��b?��N?��K?��p)
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

' ?��?��?��l?��̒l?��]?��L(BizUD?��S?��V?��b?��N?��K?��p)
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

' ?��?��?��?��?��?��t?��̓]?��L(yyyy?��Nm?��?��d?��?��?��E?��?��?��?��)
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

' ?��?��?��?��?��?��?��?��?��V?��[?��g?��ւ̓]?��L?��B?��?��?��ʕ�(S1:U1/Q2:V2/?��s20-34)?��?�� ApplyContractorStyleCommonFields?��B
'   G8:K9 ?��Ǝ҃R?��[?��h(?��?��җpE9?��E?��?��?��?��) / B12:K12 ?��Z?��?��(?��?��{?��?��?��14?��s?��ځE?��?��?��l?��?��) /
'   B14:K14 ?��?��Ж�(?��?��{?��?��?��11?��s?��ځE?��?��?��l?��?��) / C15:I16 ?��?��\?��Җ�(?��?��{?��?��?��12?��s?��ځE?��E?��l?��?��) /
'   M9:V9 ?��Z?��?��(?��?��җpM10?��E?��?��?��l?��?��) / M10:V10 ?��?��S?��H?��Ɗ�?��?��?��?��?��+?��?��o?��?��?��?��(?��?��җpM11?��E?��?��?��l?��?��) /
'   M11:V11 ?��?��E?��?��?��?��(?��?��җpM12)+?��u?��@?��@?��a?��v(?��E?��E?��l?��?��)
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

    ' G8:K9: ?��Ǝ҃R?��[?��h(?��?��җpE9), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("G8"), MirroredContractorText(wsContractor, "E9"), True

    ' B12:K12: ?��Z?��?��(?��?��{?��?��?�� ?��{?��H?��?��Ѓu?��?��?��b?��N14?��s?��?��), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("B12"), wsInfo.Cells(CONTRACTOR_ADDRESS_ROW, valueColumn).value

    ' B14:K14: ?��?��Ж�(?��?��{?��?��?�� ?��{?��H?��?��Ѓu?��?��?��b?��N11?��s?��?��), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("B14"), wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' C15:I16: ?��?��\?��Җ�(?��?��{?��?��?�� ?��{?��H?��?��Ѓu?��?��?��b?��N12?��s?��?��), ?��E?��l?��?��
    WriteHeaderValueRight wsTarget.Range("C15"), wsInfo.Cells(CONTRACTOR_REPRESENTATIVE_ROW, valueColumn).value

    ' M9:V9: ?��Z?��?��(?��?��җpM10), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M10")

    ' M10:V10: ?��?��S?��H?��Ɗ�?��?��?��?��?��+?��?��o?��?��?��?��(?��?��җpM11), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M11")

    ' M11:V11: ?��?��E?��?��?��?��(?��?��җpM12) + ?��u?��@?��@?��a?��v, ?��E?��l?��?��
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

' ?��x?��X?��T?��V?��[?��g?��ւ̓]?��L?��B?��?��?��ʕ�(S1:U1/Q2:V2/?��s20-34)?��?�� ApplyContractorStyleCommonFields?��B
'   E9:I10 ?��Ǝ҃R?��[?��h(?��?��җpE9?��E?��?��?��?��) / A13:I15 ?��?��Ж�(?��?��җpA13?��E?��?��?��?��) /
'   M8:V8 ?��Z?��?��(?��?��җpM10?��E?��?��?��l?��?��) / M9:V9 ?��?��S?��H?��Ɗ�?��?��?��?��?��+?��?��o?��?��?��?��(?��?��җpM11?��E?��?��?��l?��?��) /
'   M10:V10 ?��?��E?��?��?��?��(?��?��җpM12?��E?��E?��l?��?��)
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

    ' E9:I10: ?��Ǝ҃R?��[?��h(?��?��җpE9), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("E9"), MirroredContractorText(wsContractor, "E9"), True

    ' A13:I15: ?��?��Ж�(?��?��җpA13), ?��?��?��?��
    WriteHeaderValue wsTarget.Range("A13"), MirroredContractorText(wsContractor, "A13"), True

    ' M8:V8: ?��Z?��?��(?��?��җpM10), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("M8"), MirroredContractorText(wsContractor, "M10")

    ' M9:V9: ?��?��S?��H?��Ɗ�?��?��?��?��?��+?��?��o?��?��?��?��(?��?��җpM11), ?��?��?��l?��?��
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M11")

    ' M10:V10: ?��?��E?��?��?��?��(?��?��җpM12), ?��E?��l?��?��
    WriteHeaderValueRight wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M12")

    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?��ʎ�?��V?��V?��[?��g?��ւ̓]?��L(?��]?��L?��d?��l?��?��?��m?��肵?��?��?��炱?��?��?��֎�?��?��?��?��?��?��)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    On Error GoTo ErrorHandler
    ' O1: �{�H��Ж�(��{��� �{�H��З�11�s��)���E�l�߁EBIZ UD�S�V�b�N�œ]�L
    Dim vendorCol As Long
    vendorCol = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    Dim vendorName As String
    vendorName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, vendorCol).Address)
    WriteAttachment3VendorName wsTarget.Range("O1"), vendorName

    ' 36�`41�s�� F/H��EJ/L��̔r���`�F�b�N�{�b�N�X����
    SetupAttachment3CheckboxExclusivity wsTarget

    ' E49:I49�`E52:I52 �_�u���N���b�N�ɂ��Y�Ɣp���������{�ݑI��
    SetupAttachment3SanpaiFacilityDoubleClickHint wsTarget

    ' J54:M54 �Y�p�sJR���z���v(�H���V�[�g�ɕR�t�����{�H��Ђ݂̂ɓ���)
    RefreshAttachment3SanpaiJrTotal wsTarget, wsInfo, vendorIndex
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyAttachment3Header error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' O1(�{�H��Ж�)�̓]�L(�E�l�߁EBIZ UD�S�V�b�N)
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

' ===== �ʎ��V 36�`41�s�� F/H��EJ/L�� �r���`�F�b�N�{�b�N�X =====
' F��=�L / H��=�� �̃y�A�ƁAJ��EL��̃y�A�����ꂼ��r��������B
' ����� H��(��)���`�F�b�N����Ă���s�� J��EL��̂ǂ�����`�F�b�N�ł��Ȃ��B
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

' �`�F�b�N�{�b�N�X�N���b�N���̔r������{��(�e�`�F�b�N�{�b�N�X�� OnAction ����Ă΂��)
Public Sub Attachment3CheckboxClick()
    On Error GoTo Done

    Dim callerName As String
    callerName = Application.Caller

    ' �N���b�N���ꂽ�`�F�b�N�{�b�N�X�͕K���A�N�e�B�u�V�[�g��ɂ���B
    ' �����̕ʎ��V(���)�Ɠ����`�F�b�N�{�b�N�X�����݂��邽�߁A�S�V�[�g����
    ' �擪��v�Ŏ��V�[�g�ɌН������̂�h���A�A�N�e�B�u�V�[�g�Ɍ��肷��B
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
                ' H(��)��I�������s�� J��EL��̂ǂ�����`�F�b�N�s�Ƃ���
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

' �����̑I����e�ɖ���(F/H�����`�F�b�N�AH+J/L�̕��p�AJ/L�����`�F�b�N��)������ꍇ�̂ݕ␳����B
' �V�K�Ɋ���l��t�^���邱�Ƃ͂��Ȃ��B
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

' ===== �ʎ��V E49:I49�`E52:I52 �Y�Ɣp���������{�ݑI�� =====
' �_�u���N���b�N�Ώ۔͈͂͐������_�ŏ����E�ی쓙�̒ǉ��ݒ�͕s�v�̂��ߌ���͋�����B
' (�Ώ۔���E�I�������� ThisWorkbook ������ mod_OrderTpl_Header �̊֐��𒼐ڌĂяo��)
Private Sub SetupAttachment3SanpaiFacilityDoubleClickHint(ByVal ws As Worksheet)
    ' �\��(�����̏����t�^���ɔ������t�b�N)
End Sub

' Attachment3(�ʎ��V)�V�[�g���ǂ����𔻒肷��
Public Function IsAttachment3Sheet(ByVal ws As Worksheet) As Boolean
    IsAttachment3Sheet = False
    If ws Is Nothing Then Exit Function
    Dim baseName As String, aliasText As String
    If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then
        IsAttachment3Sheet = (StrComp(baseName, mod_OrderTpl_Shared.OrderTplBaseNameAttachment3Text(), vbTextCompare) = 0)
    End If
End Function

' E49:I49�`E52:I52 (�Y�Ɣp���������{�ݑI���Z��)���ǂ����𔻒肷��
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

' frmSubconSelector.frm ���ė��p���ĎY�Ɣp���������{�݂�I�����A
' E��(�{�ݖ�)�EJ��(�Ή��l�A�Ǝ҃}�X�^B��)�֏������ށB
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
        MsgBox "�Y�Ɣp���������{�݂̃}�X�^���擾�ł��܂���ł����B", vbExclamation
        Exit Sub
    End If
    If names.Count = 0 Then
        MsgBox "�Y�Ɣp���������{�݂̃}�X�^�Ƀf�[�^������܂���B", vbExclamation
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

' �Ǝ҃}�X�^(�S�Д�).xlsx �̎x�X���V�[�g���� A59�ȉ�(�{�ݖ�)�EB59�ȉ�(�Ή��l)��ǂݍ��ށB
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

' "�Y�Ɣp���������{�ݑI��"
Private Function Attachment3SanpaiFacilityCaptionText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H7523, &H696D, &H5EC3, &H68C4, &H7269, _
                                      &H51E6, &H7406, &H65BD, &H8A2D, &H9078, &H629E)
    End If
    Attachment3SanpaiFacilityCaptionText = cached
End Function

' ===== �ʎ��V J54:M54 �Y�p�sJR���z���v =====
' �{�s�w����(�H��)/�{�s�ʒm��(�H��)�V�[�g�̎Y�p�s(�{�H��Ђ������I���ł��Ȃ������s)��
' JR���z�񍇌v���A����؂萮���� J54:M54(�����Z��)�֓]�L����B
Private Sub RefreshAttachment3SanpaiJrTotal(ByVal ws As Worksheet, _
                                            ByVal wsInfo As Worksheet, _
                                            ByVal vendorIndex As Long)
    If ws Is Nothing Then Exit Sub
    On Error GoTo Done

    Dim jrTotalArea As Range
    Set jrTotalArea = ws.Range("J54").MergeArea      ' J54:M54(�����Z��)
    Dim writeCell As Range
    Set writeCell = jrTotalArea.Cells(1, 1)

    ' �{�H�w����(�H��)/�{�H�ʒm��(�H��)��A���ɕR�t���Ă��Ȃ��{�H��Ђ̕ʎ�III�ɂ�
    ' �Y�pJR���z���v����͂��Ȃ�(�����Z���S�̂��N���A)�B
    ' ���̌����Z���̈ꕔ(Cells(1,1))�ɑ΂��� ClearContents �� 1004
    '   �u���̑���͌����������Z���ɂ͍s���܂���v�ɂȂ邽�߁AMergeArea �S�̂��N���A����B
    If Not IsAttachment3VendorLinkedToWorks(wsInfo, vendorIndex) Then
        jrTotalArea.ClearContents
        Exit Sub
    End If

    Dim total As Double
    total = mod_Construction_BasicTotals.SumSanpaiJrAmount()

    writeCell.value = Int(total)   ' ���؂�(�����؂�̂�)�œ��͂���
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

' ?��?��?��?��?��Z?��?��?��Ή�?��̒l?��]?��L(?��t?��H?��?��?��g?��K?��p?��A?��K?��v?��ɉ�?��?��?��ď㉺?��?��?��E?��?��?��?��?��?��?��?��)
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

' ?��?��?��?��?��?��Ƃ�?��Ă̓]?��L(?��?��?��X?��R?��[?��h?��?��?��A?��?��?��t?��?��ϊ�?��?��h?��?��)
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

' ?��?��?��t?��̓]?��L(?��a?��?��\?��?��?��`?��?��?��A?��?��?��?��?��Z?��?��?��̂�?��ߒ�?��?��?��?��?��?��)
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

' ?��?��җp?��V?��[?��gwsTarget?��Ɠ�?��?��O?��?��?��[?��v?��̓�?��?��׃V?��[?��g?��?��?��?��?��?��?��?��?��?��?��?��
Private Function ResolveBreakdownSheetNameFromTarget(ByVal wsTarget As Worksheet) As String
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then ResolveBreakdownSheetNameFromTarget = nm
End Function

' ?��?��?��?��׃V?��[?��g?��?��A?��?���??��?��x?��?��?��s?��?��T?��?��?��A?��?��?��?��Q?��?��?��Ԃ�?��?��?��?��?��?��?��?��(?��?��?���?��?��Ȃ�?��?��΋�)
Private Function BuildBreakdownQFormula(ByVal sheetName As String, ByVal labelText As String) As String
    Dim q As String
    q = "'" & Replace$(sheetName, "'", "''") & "'"
    BuildBreakdownQFormula = "=IFERROR(INDEX(" & q & "!Q:Q,MATCH(""" & labelText & """," & q & "!A:A,0)),"""")"
End Function

' ?��?��?��?��(?��E?��l)?��̓]?��L(BizUD?��S?��V?��b?��N?��E?��?��?��?��؂�)
Private Sub WriteHeaderFormulaRight(ByVal target As Range, ByVal formulaText As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    writeCell.Formula = formulaText
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' "?��?��?��v"
Private Function ContractorGrandTotalLabelText() As String
    ContractorGrandTotalLabelText = ChrW$(&H5408) & ChrW$(&H8A08)
End Function

' "?��v"
Private Function ContractorNetTotalLabelText() As String
    ContractorNetTotalLabelText = ChrW$(&H8A08)
End Function

' "?��?��?��?��?��"
Private Function ContractorTaxLabelText() As String
    ContractorTaxLabelText = ChrW$(&H6D88) & ChrW$(&H8CBB) & ChrW$(&H7A0E)
End Function

' "?��@?��@?��a"(?��S?��p?��?��2?��?�� + ?��a?��B?��?��?��?��?��?��?��?�� M11:V11 ?��̖�?��?��?��֕t?��т�?��?��)
Private Function ContractorHonorificSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H6BBF)
    End If
    ContractorHonorificSuffixText = cached
End Function
