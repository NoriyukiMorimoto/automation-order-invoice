Option Explicit

' �������e���v���[�g�e�V�[�g�ւ̊�{���w�b�_�[�]�L�B
' �������ɉ����A��{���V�[�g�̓]�L���Z���ύX���ɂ��ē]�L�����(���C�u���f)�B
' �]�L�ΏۃV�[�g�̒ǉ��� ApplyVendorSheetHeaders �̃f�B�X�p�b�`�֎������������ށB
' ���C����: CHANGELOG.md �Q��

' ��{���̃w�b�_�[�]�L���Z��(�S�Ћ��ʕ�)�B�u���b�N��(16/27�s��)�͓��I�ɑg�ݗ��Ă�
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' �󒍎җp�V�[�g�]�L�ŎQ�Ƃ���{�H��Ѓu���b�N�s(��{���)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35
' ��{��� �{�H��Ѓu���b�N: ��\�Җ�(12�s��)�E�Z��(14�s��)
Private Const CONTRACTOR_REPRESENTATIVE_ROW As Long = 12
Private Const CONTRACTOR_ADDRESS_ROW As Long = 14

' �w��u���b�N�̎{�H��ЂɑΉ�����e���v���[�g5�V�[�g�փw�b�_�[��]�L����(�f�B�X�p�b�`��)
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
    End If
End Sub

' �S�m���Ђ̃e���v���[�g�V�[�g�փw�b�_�[���ē]�L����
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
                End If
            End If
        End If
    Next vendorIndex
    Exit Sub

Quiet:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorSheetHeaders error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Sheet1(��{���)��Worksheet_Change����Ă΂������B
' �w�b�_�[�]�L���Z��(B6/C6/C2/C9/C10/C15:C16/F6�A�e�u���b�N��16/27�s��)�̕ύX���e�ЃV�[�g�֔��f����
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

' �w�b�_�[�]�L���Z���̊Ď��͈͂�Ԃ����J���b�p�[(Sheet1�̕ύX�Q�[�g�\�z�p)
Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' �w�b�_�[�]�L���Z���̊Ď��͈�(���ʃZ�� + �e�u���b�N�̋Ǝ҃R�[�h16�s��/�����ԍ�27�s��)
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

' �V�[�g���o��(�^�u)�F: �H���敪(��{���10�s��)�Z���̓h��Ԃ��F��K�p����
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


' ���󖾍׃w�b�_�[���֊�{���V�[�g�̓��e��]�L����
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: ���X�R�[�h(�o������_�P���K�p����̒P���K�p����V�[�g B/C��ƍ� �� G��)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: �����ԍ�(�{�H��Ѓu���b�N27�s��)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: �H���ԍ�(��{���C9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' D6: �H������(��{���C10�A�����Z���̂��ߒ�������)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: �H�� ��/��(��{���C15/C16�A�a��\��)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: �쐬��(��{���C2�A�a��\��)
    WriteHeaderDate wsBreakdown.Range("P2"), wsInfo.Range("C2").value
    wsBreakdown.Range("P2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE
    wsBreakdown.Range("O2").MergeArea.Cells(1, 1).ClearContents

    ' O3: ������(��{���F6)
    WriteHeaderValue wsBreakdown.Range("P3"), wsInfo.Range("F6").value, True
    wsBreakdown.Range("O3").MergeArea.Cells(1, 1).ClearContents

    ' P5: �O����Ж�(�{�H��Ѓu���b�N11�s��)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: �Ǝ҃R�[�h(�{�H��Ѓu���b�N16�s��)
    WriteHeaderValue wsBreakdown.Range("P6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' �����y�[�W����p�Ƀw�b�_�[�s(7:10�s��)���^�C�g���s�ɐݒ肷��
    On Error Resume Next
    wsBreakdown.PageSetup.PrintTitleRows = ORDER_TPL_PRINT_TITLE_ROWS
    On Error GoTo ErrorHandler

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader done: " & wsBreakdown.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' 軌道工事(条件書(軌道工事)のみH25:U25転記の判定用)
Private Function ConditionRailWorkTypeText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H8ECC, &H9053, &H5DE5, &H4E8B)
    End If
    ConditionRailWorkTypeText = cached
End Function

' 条件書シートへ基本情報/当該施工会社の内容を転記する。
' 4条件書共通の10項目 + 条件書(軌道工事)のみ H25:U25 <- 当該業者41行(左詰め)。
Public Sub ApplyConditionSheetHeader(ByVal wsInfo As Worksheet, _
                                     ByVal wsCondition As Worksheet, _
                                     ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsCondition Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' S1:X1 作成日(基本情報C2、グレゴリオ暦 yyyy年m月d日、中央)
    WriteHeaderDateGregorian wsCondition.Range("S1"), wsInfo.Range("C2").value
    wsCondition.Range("S1").MergeArea.Cells(1, 1).HorizontalAlignment = xlCenter

    ' P3:S3 支店(B6) / T3:X3 出張所(C6) / T4:X4 所長名(F6)  下詰め
    WriteHeaderValueBottom wsCondition.Range("P3"), wsInfo.Range("B6").value
    WriteHeaderValueBottom wsCondition.Range("T3"), wsInfo.Range("C6").value
    WriteHeaderValueBottom wsCondition.Range("T4"), wsInfo.Range("F6").value

    ' B5:C6 施工会社名(当該業者11行), 下詰め
    WriteHeaderValueBottom wsCondition.Range("B5"), _
        wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' F9:I9 工事番号(C9) / L9:X9 工事件名(C10)  中央
    WriteHeaderValue wsCondition.Range("F9"), wsInfo.Range("C9").value, True
    WriteHeaderValue wsCondition.Range("L9"), wsInfo.Range("C10").value, True

    ' F11:L11 工期自(C15) / R11:X11 工期至(C16)  グレゴリオ暦 yyyy年m月d日
    WriteHeaderDateGregorian wsCondition.Range("F11"), wsInfo.Range("C15").value
    WriteHeaderDateGregorian wsCondition.Range("R11"), wsInfo.Range("C16").value

    ' D16:X16 施工場所(C13), 中央
    WriteHeaderValue wsCondition.Range("D16"), wsInfo.Range("C13").value, True

    ' 条件書(軌道工事)のみ: H25:U25 <- 当該業者41行(左詰め)
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

' ===== 条件書チェックボックスの排他グループ化 =====
' 各行D/X(10行/18-33行/38行)と E34/E35 のペアで、常にどちらか一方だけONになる。
' 片方をONにするともう片方をOFF、片方をOFFにするともう片方をONにする。

Private Const CONDITION_CHECKBOX_D_COL As Long = 4
Private Const CONDITION_CHECKBOX_X_COL As Long = 24
Private Const CONDITION_CHECKBOX_E_COL As Long = 5

' 排他ペアの対象行かどうか(D/X: 10,18-33,38 / E縦ペア: 34,35)
Private Function IsConditionDxExclusiveRow(ByVal rowIndex As Long) As Boolean
    IsConditionDxExclusiveRow = (rowIndex = 10) Or (rowIndex >= 18 And rowIndex <= 33) Or (rowIndex = 38)
End Function

Private Function IsConditionEVerticalPairRow(ByVal rowIndex As Long) As Boolean
    IsConditionEVerticalPairRow = (rowIndex = 34) Or (rowIndex = 35)
End Function

Private Function IsConditionDxCheckboxColumn(ByVal colIndex As Long) As Boolean
    IsConditionDxCheckboxColumn = (colIndex = CONDITION_CHECKBOX_D_COL Or colIndex = CONDITION_CHECKBOX_X_COL)
End Function

Private Function IsConditionExclusiveCheckbox(ByVal cb As Object) As Boolean
    Dim rowIndex As Long
    Dim colIndex As Long
    rowIndex = cb.TopLeftCell.Row
    colIndex = cb.TopLeftCell.Column

    If IsConditionEVerticalPairRow(rowIndex) Then
        IsConditionExclusiveCheckbox = (colIndex = CONDITION_CHECKBOX_E_COL)
        Exit Function
    End If

    If IsConditionDxExclusiveRow(rowIndex) Then
        IsConditionExclusiveCheckbox = IsConditionDxCheckboxColumn(colIndex)
    ElseIf rowIndex = 39 And IsConditionDxCheckboxColumn(colIndex) Then
        ' 38行目のコントロールが39行にアンカーされる場合がある
        IsConditionExclusiveCheckbox = True
    End If
End Function

Private Function ConditionDxPairTargetColumn(ByVal clickedCol As Long) As Long
    If clickedCol <= CONDITION_CHECKBOX_D_COL Then
        ConditionDxPairTargetColumn = CONDITION_CHECKBOX_X_COL
    ElseIf clickedCol >= CONDITION_CHECKBOX_X_COL Then
        ConditionDxPairTargetColumn = CONDITION_CHECKBOX_D_COL
    End If
End Function

' 排他ペアの対象行かどうか(D/X: 10,18-33,38 / E縦ペア: 34,35)
Private Function IsConditionExclusiveRow(ByVal r As Long) As Boolean
    IsConditionExclusiveRow = IsConditionDxExclusiveRow(r) Or IsConditionEVerticalPairRow(r)
End Function

' 条件書シートの対象チェックボックスに排他クリックのマクロを割り当てる
Public Sub SetupConditionCheckboxExclusivity(ByVal wsCondition As Worksheet)
    If wsCondition Is Nothing Then Exit Sub
    On Error Resume Next
    Dim cb As Object
    For Each cb In wsCondition.CheckBoxes
        If IsConditionExclusiveCheckbox(cb) Then
            cb.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
        End If
    Next cb
    On Error GoTo 0
End Sub

' チェックボックスクリック時に呼ばれる。ペアの相手を反対状態にする。
Public Sub ConditionCheckboxExclusiveClick()
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim clicked As Object
    Set clicked = ws.CheckBoxes(Application.Caller)
    If clicked Is Nothing Then Exit Sub

    Dim pair As Object
    Set pair = FindConditionCheckboxPair(ws, clicked)
    If pair Is Nothing Then Exit Sub

    If clicked.Value = xlOn Then
        pair.Value = xlOff
    Else
        pair.Value = xlOn
    End If

Done:
End Sub

' クリックされたチェックボックスのペアを返す(D/X=反対列・近傍Top / E34-E35=行34と35)
Private Function FindConditionCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim rowIndex As Long
    rowIndex = clicked.TopLeftCell.Row

    If IsConditionEVerticalPairRow(rowIndex) Then
        Set FindConditionCheckboxPair = FindConditionEVerticalCheckboxPair(ws, clicked)
        Exit Function
    End If

    If IsConditionDxExclusiveRow(rowIndex) Or _
       (rowIndex = 39 And IsConditionDxCheckboxColumn(clicked.TopLeftCell.Column)) Then
        Set FindConditionCheckboxPair = FindConditionDxCheckboxPair(ws, clicked)
    End If
End Function

Private Function FindConditionEVerticalCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim rowIndex As Long
    rowIndex = clicked.TopLeftCell.Row

    Dim targetRow As Long
    If rowIndex = 34 Then
        targetRow = 35
    Else
        targetRow = 34
    End If

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If cb.TopLeftCell.Row = targetRow And cb.TopLeftCell.Column = CONDITION_CHECKBOX_E_COL Then
                Set FindConditionEVerticalCheckboxPair = cb
                Exit Function
            End If
        End If
    Next cb
End Function

Private Function FindConditionDxCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim clickedCol As Long
    clickedCol = clicked.TopLeftCell.Column
    Dim pairCol As Long
    pairCol = ConditionDxPairTargetColumn(clickedCol)
    If pairCol = 0 Then Exit Function

    Dim clickedRow As Long
    clickedRow = clicked.TopLeftCell.Row
    Dim rowTolerance As Long
    If clickedRow = 38 Or clickedRow = 39 Then
        rowTolerance = 1
    Else
        rowTolerance = 0
    End If

    Dim bestCb As Object
    Dim bestTopDist As Double
    bestTopDist = -1

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If cb.TopLeftCell.Column = pairCol Then
                If Abs(cb.TopLeftCell.Row - clickedRow) <= rowTolerance Then
                    Dim topDist As Double
                    topDist = Abs(cb.Top - clicked.Top)
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

' �󒍎җp/��������/�x�X�T ����: S1:U1 �����ԍ�(27) / Q2:V2 �쐬��C2(����) /
'   �s20-34(E20:�H����C10 E22:�s���{��C13 G24:�H����C15(����) G26:�H����C16(����)�A
'   Q22:�ō�(35) Q23:�Ŕ�(33) Q24:�����(34)�AC30/H30/J30/M34/R34/F32/F33 �� Reapply�n�o�R)��
'   ApplyContractorStyleCommonFields �ŁA3�V�[�g�Ƃ����ꃍ�W�b�N��K�p����B
' �󒍎җp�V�[�g�ŗL: E9:�Ǝ҃R�[�h(16) A13:��Ж�(11) M10:�����ҏZ��
'   M11:��S�H�Ɗ������+��o���� M12:��E����(�o���������X�g�Q��)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    ' E9: �Ǝ҃R�[�h(�{�H��Ѓu���b�N16�s), ����
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: ��Ж�(�{�H��Ѓu���b�N11�s), ����
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' M10/M11/M12: �o���������X�g�Q��(������ �Z��/����/��E����)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    ' �s38-42(��������/�J��/�x���ޗ�/�ݗ^�i/���T�C�N��)����{��񂩂�
    ' �󒍎җp�E���������E�x�X�T�֍ē]�L(Reapply���őS�Ă̐����σV�[�g�֓K�p)
    mod_BasicInfoExclusiveChoice.ReapplyExclusiveChoices wsInfo, vendorIndex
    mod_BasicInfoSupplyLoan.ReapplySupplyLoan wsInfo, vendorIndex

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' �󒍎җp/��������/�x�X�T�ŋ��ʂ̓]�L(S1:U1 �����ԍ� / Q2:V2 �쐬��(����) /
' �s20-34: E20 �H�����EE22 �s���{���EG24/G26 �H������(����)�EQ22-24 ���󖾍׎Q��)�B
' �A3�V�[�g�Ƃ��Z���\�������S��v���Ă��邽�߁A���ꃍ�W�b�N�����ꂼ��K�p����B
Private Sub ApplyContractorStyleCommonFields(ByVal wsInfo As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal valueColumn As Long)
    ' S1:U1: �����ԍ�(�{�H��Ѓu���b�N27�s), ����
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2:V2: �쐬��(��{���C2, ����), ����
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E20: �H����(��{���C10), ����
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: �s���{��(��{���C13), ����
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: �H�� ��(��{���C15, ����), ����
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: �H�� ��(��{���C16, ����), ����
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22/Q23/Q24: ����O���[�v�̓��󖾍׃V�[�g ���v/�v/����� �s(Q��)���Q�Ƃ��鐶������
    ' (���󖾍ׂ̒l���X�V�����Ǝ����ōX�V�����)
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

' wsTarget(�����σe���v���[�g�V�[�g)�Ɠ���O���[�v�̎󒍎җp(����)�V�[�g����������
Private Function ResolveContractorSheetFromTarget(ByVal wsTarget As Worksheet) As Worksheet
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then Set ResolveContractorSheetFromTarget = ThisWorkbook.Worksheets(nm)
End Function

' �󒍎җp�V�[�g�̌����Z������l���擾����(���݂��Ȃ�/��Ȃ�󕶎�)
Private Function MirroredContractorText(ByVal wsContractor As Worksheet, ByVal address As String) As String
    If wsContractor Is Nothing Then Exit Function
    MirroredContractorText = CommonNzText(wsContractor.Range(address).MergeArea.Cells(1, 1).value)
End Function

' �󒍎җp M10:V �s�̌������e���v���[�g(M:U)������S�Ɋg������
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

' M10:�Z�� / M11:��S�H�Ɗ������+�S�p��+��o���� / M12:(������)��E���� ��]�L
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
    ' M10:U �� M10:V �֌����������A�k�����đS�̕\���ɐݒ�(M/13�� �` V/22��)
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

    ' M10: �Z��
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: ��S�H�Ɗ������ + �S�p�� + ��o����
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: �o����=��o�����Ȃ�u��E �����v�A�s��v�Ȃ�u�o���� ��E �����v
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValueRight wsTarget.Range("M12"), m12
End Sub

' �E�l�̒l�]�L(BizUD�S�V�b�N�K�p)
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

' ���l�̒l�]�L(BizUD�S�V�b�N�K�p)
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

' ������t�̓]�L(yyyy�Nm��d���E����)
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

' ���������V�[�g�ւ̓]�L�B���ʕ�(S1:U1/Q2:V2/�s20-34)�� ApplyContractorStyleCommonFields�B
'   G8:K9 �Ǝ҃R�[�h(�󒍎җpE9�E����) / B12:K12 �Z��(��{���14�s�ځE���l��) /
'   B14:K14 ��Ж�(��{���11�s�ځE���l��) / C15:I16 ��\�Җ�(��{���12�s�ځE�E�l��) /
'   M9:V9 �Z��(�󒍎җpM10�E���l��) / M10:V10 ��S�H�Ɗ������+��o����(�󒍎җpM11�E���l��) /
'   M11:V11 ��E����(�󒍎җpM12)+�u�@�@�a�v(�E�E�l��)
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

    ' G8:K9: �Ǝ҃R�[�h(�󒍎җpE9), ����
    WriteHeaderValue wsTarget.Range("G8"), MirroredContractorText(wsContractor, "E9"), True

    ' B12:K12: �Z��(��{��� �{�H��Ѓu���b�N14�s��), ���l��
    WriteHeaderValueLeft wsTarget.Range("B12"), wsInfo.Cells(CONTRACTOR_ADDRESS_ROW, valueColumn).value

    ' B14:K14: ��Ж�(��{��� �{�H��Ѓu���b�N11�s��), ���l��
    WriteHeaderValueLeft wsTarget.Range("B14"), wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' C15:I16: ��\�Җ�(��{��� �{�H��Ѓu���b�N12�s��), �E�l��
    WriteHeaderValueRight wsTarget.Range("C15"), wsInfo.Cells(CONTRACTOR_REPRESENTATIVE_ROW, valueColumn).value

    ' M9:V9: �Z��(�󒍎җpM10), ���l��
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M10")

    ' M10:V10: ��S�H�Ɗ������+��o����(�󒍎җpM11), ���l��
    WriteHeaderValueLeft wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M11")

    ' M11:V11: ��E����(�󒍎җpM12) + �u�@�@�a�v, �E�l��
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

' �x�X�T�V�[�g�ւ̓]�L�B���ʕ�(S1:U1/Q2:V2/�s20-34)�� ApplyContractorStyleCommonFields�B
'   E9:I10 �Ǝ҃R�[�h(�󒍎җpE9�E����) / A13:I15 ��Ж�(�󒍎җpA13�E����) /
'   M8:V8 �Z��(�󒍎җpM10�E���l��) / M9:V9 ��S�H�Ɗ������+��o����(�󒍎җpM11�E���l��) /
'   M10:V10 ��E����(�󒍎җpM12�E�E�l��)
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

    ' E9:I10: �Ǝ҃R�[�h(�󒍎җpE9), ����
    WriteHeaderValue wsTarget.Range("E9"), MirroredContractorText(wsContractor, "E9"), True

    ' A13:I15: ��Ж�(�󒍎җpA13), ����
    WriteHeaderValue wsTarget.Range("A13"), MirroredContractorText(wsContractor, "A13"), True

    ' M8:V8: �Z��(�󒍎җpM10), ���l��
    WriteHeaderValueLeft wsTarget.Range("M8"), MirroredContractorText(wsContractor, "M10")

    ' M9:V9: ��S�H�Ɗ������+��o����(�󒍎җpM11), ���l��
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M11")

    ' M10:V10: ��E����(�󒍎җpM12), �E�l��
    WriteHeaderValueRight wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M12")

    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' �ʎ��V�V�[�g�ւ̓]�L(�]�L�d�l���m�肵���炱���֎�������)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    ' �]�L�d�l ���w��(�R�s�[�̂�)
End Sub

' �����Z���Ή��̒l�]�L(�t�H���g�K�p�A�K�v�ɉ����ď㉺���E��������)
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

' ������Ƃ��Ă̓]�L(���X�R�[�h���A���t��ϊ���h��)
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

' ���t�̓]�L(�a��\���`���A�����Z���̂��ߒ�������)
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

' �󒍎җp�V�[�gwsTarget�Ɠ���O���[�v�̓��󖾍׃V�[�g������������
Private Function ResolveBreakdownSheetNameFromTarget(ByVal wsTarget As Worksheet) As String
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then ResolveBreakdownSheetNameFromTarget = nm
End Function

' ���󖾍׃V�[�g��A��Ń��x���s��T���A����Q���Ԃ���������(������Ȃ���΋�)
Private Function BuildBreakdownQFormula(ByVal sheetName As String, ByVal labelText As String) As String
    Dim q As String
    q = "'" & Replace$(sheetName, "'", "''") & "'"
    BuildBreakdownQFormula = "=IFERROR(INDEX(" & q & "!Q:Q,MATCH(""" & labelText & """," & q & "!A:A,0)),"""")"
End Function

' ����(�E�l)�̓]�L(BizUD�S�V�b�N�E����؂�)
Private Sub WriteHeaderFormulaRight(ByVal target As Range, ByVal formulaText As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    writeCell.Formula = formulaText
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' "���v"
Private Function ContractorGrandTotalLabelText() As String
    ContractorGrandTotalLabelText = ChrW$(&H5408) & ChrW$(&H8A08)
End Function

' "�v"
Private Function ContractorNetTotalLabelText() As String
    ContractorNetTotalLabelText = ChrW$(&H8A08)
End Function

' "�����"
Private Function ContractorTaxLabelText() As String
    ContractorTaxLabelText = ChrW$(&H6D88) & ChrW$(&H8CBB) & ChrW$(&H7A0E)
End Function

' "�@�@�a"(�S�p��2�� + �a�B�������� M11:V11 �̖����֕t�т���)
Private Function ContractorHonorificSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H6BBF)
    End If
    ContractorHonorificSuffixText = cached
End Function
