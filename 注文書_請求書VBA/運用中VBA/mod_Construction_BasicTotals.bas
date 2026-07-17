Option Explicit

' ���C����: CHANGELOG.md �Q��
' mod_Construction_BasicTotals (split from mod_Construction_Order_Import)

Public Sub RefreshBasicInfoConstructionTotalsCore(Optional ByVal changedVendorIndex As Long = 0)
    On Error GoTo ErrorHandler

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim fullRefresh As Boolean
    fullRefresh = (changedVendorIndex <= 0)

    If Not fullRefresh Then
        If changedVendorIndex < 1 Or changedVendorIndex > vendorCount Then Exit Sub
    End If

    Dim vendorNames() As String
    Dim vendorTotals() As Double
    ReDim vendorNames(1 To BASIC_INFO_VENDOR_MAX_BLOCKS)
    ReDim vendorTotals(1 To BASIC_INFO_VENDOR_MAX_BLOCKS)

    Dim vendorStart As Long
    Dim vendorEnd As Long
    If fullRefresh Then
        vendorStart = 1
        vendorEnd = vendorCount
    Else
        vendorStart = changedVendorIndex
        vendorEnd = changedVendorIndex
    End If

    Dim i As Long
    For i = vendorStart To vendorEnd
        vendorNames(i) = GetBasicInfoCellText(wsInfo, _
            wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, BasicInfoVendorColumn(i)).Address)
    Next i

    If fullRefresh Then
        Dim vendorNameLog As String
        For i = 1 To vendorCount
            vendorNameLog = vendorNameLog & " [" & i & ":" & vendorNames(i) & "]"
        Next i
        LogCI "��{���Ǝ� F9����=" & vendorCount & vendorNameLog
    End If

    Dim branchName As String
    branchName = GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL)
    Dim vendorAliasMap As Object
    Set vendorAliasMap = GetVendorAliasMap(branchName)

    Dim worksTotal As Double
    Dim purchaseTotal As Double

    ' �����X�V(�P��Ǝ�)�ł����󖾍��u�v�v�sQ��D��̗p����(���̏�����)�B
    ' ���Q���擾���A���l�Ȃ��Ǝҕ ʏW�v���[�v(�S�{�H�V�[�g����+��}�b�v�\�z)��
    ' vendorTotals���g���Ȃ����ߊۂ��Əȗ����ĕ��ׂ���炷�B�W�v���[�v�͓ǎ���p��
    ' ���ʂ������̂ŁA�ȗ����Ă�����(�������ޒl)�͓���B
    Dim partialNetTotalQ As Variant
    Dim partialUsesNetQ As Boolean
    partialUsesNetQ = False
    If Not fullRefresh Then
        If vendorNames(changedVendorIndex) <> "" Then
            partialNetTotalQ = mod_OrderTpl_Detail.OrderTplGetBreakdownNetTotalQForVendor(wsInfo, changedVendorIndex)
            partialUsesNetQ = IsNumeric(partialNetTotalQ)
        End If
    End If

    Dim ws As Worksheet
    Dim columnMap As Object
    Dim vendorKey As String
    If fullRefresh Or Not partialUsesNetQ Then
    For Each ws In ThisWorkbook.worksheets
        If fullRefresh And IsPurchaseOutputSheet(ws) Then
            purchaseTotal = purchaseTotal + SumOutputJrAmount(ws)
        ElseIf IsConstructionOutputSheet(ws) Then
            If fullRefresh Then
                worksTotal = worksTotal + SumOutputJrAmount(ws)
            End If
            Set columnMap = BuildSheetVendorAmountColumnMap(ws, vendorAliasMap)
            For i = vendorStart To vendorEnd
                If vendorNames(i) <> "" Then
                    vendorKey = ResolveVendorCanonicalKey(vendorNames(i), vendorAliasMap)
                    If vendorKey <> "" Then
                        If Not columnMap Is Nothing Then
                            If columnMap.Exists(vendorKey) Then
                                vendorTotals(i) = vendorTotals(i) + _
                                    SumVendorAmountByColumn(ws, CLng(columnMap(vendorKey)))
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next ws
    End If

    If fullRefresh Then
        WriteBasicInfoAmount wsInfo, BASIC_INFO_WORKS_TOTAL_CELL, worksTotal
        WriteBasicInfoAmount wsInfo, BASIC_INFO_PURCHASE_TOTAL_CELL, purchaseTotal
        mod_Construction_BasicTotals.UpdateBasicInfoTaxTotalsCore wsInfo
    End If

    Dim totalCellAddress As String
    Dim totalCell As Range
    Dim writeStart As Long
    Dim writeEnd As Long
    If fullRefresh Then
        writeStart = 1
        writeEnd = BASIC_INFO_VENDOR_MAX_BLOCKS
    Else
        writeStart = changedVendorIndex
        writeEnd = changedVendorIndex
    End If

    For i = writeStart To writeEnd
        totalCellAddress = wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, _
                                        BasicInfoVendorColumn(i)).Address
        If i <= vendorCount And vendorNames(i) <> "" Then
            ' 33�s(�_����z�Ŕ�)�͕R�t�����󖾍׃V�[�g�́u�v�v�sQ��Ɉ�v������B
            ' ���󖾍ׂ�������/���v�Z���Ŏ擾�ł��Ȃ��ꍇ�͏]���̏W�v�l�Ƀt�H�[���o�b�N�B
            Dim netTotalQ As Variant
            If (Not fullRefresh) And i = changedVendorIndex Then
                netTotalQ = partialNetTotalQ    ' �O�Ɏ擾�ς�(��d�v�Z���)
            Else
                netTotalQ = mod_OrderTpl_Detail.OrderTplGetBreakdownNetTotalQForVendor(wsInfo, i)
            End If
            If IsNumeric(netTotalQ) Then
                WriteBasicInfoAmount wsInfo, totalCellAddress, CDbl(netTotalQ), True
            Else
                WriteBasicInfoAmount wsInfo, totalCellAddress, vendorTotals(i), True
            End If
        Else
            WriteBasicInfoAmount wsInfo, totalCellAddress, 0, False
        End If

        Set totalCell = wsInfo.Range(totalCellAddress)
        If totalCell.MergeCells Then Set totalCell = totalCell.mergeArea.Cells(1, 1)
        totalCell.NumberFormatLocal = BasicInfoYenNumberFormat()
    Next i

    ' �e�u���b�N��34/35�s��(����ŁE�ō��݋��z)��Ǐ]�X�V����
    RefreshVendorBlockTaxRows wsInfo
    Exit Sub

ErrorHandler:
    LogCI "��{��񍇌v���z�X�V�G���[ Err " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

'  RefreshVendorBlockTaxRows
'  �{�H��Ѓu���b�N��34/35�s�ڂ��X�V����B
'  ���x����(�l��̍�)�ɂ�B34:B35(�����(10%)�F/�ō��݋��z�F)�̃f�[�^�E�r�����R�s�[���A
'  �l��34�s��=33�s��(�_����z�Ŕ�)�~�ŗ�(B34�̃J�b�R���AC34�Ɠ����؂�̂�)�A
'  35�s��=33�s��+34�s�ځB������33�s�ڂ�͕�BF9�̉�А������ɂ��Ǐ]����B
Public Sub RefreshVendorBlockTaxRows(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim taxRate As Double
    taxRate = ResolveBasicInfoTaxRate(wsInfo)

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To BASIC_INFO_VENDOR_MAX_BLOCKS
        Dim valueColumn As Long
        Dim labelColumn As Long
        valueColumn = BasicInfoVendorColumn(i)
        labelColumn = valueColumn - 1

        If i <= vendorCount Then
            ' ���x��(B34:B35)�̃f�[�^�ƌr�����R�s�[
            wsInfo.Range("B34:B35").Copy Destination:=wsInfo.Cells(34, labelColumn)

            ' �l�Z���̏�����33�s��(�_����z)��͕�
            wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, valueColumn).Copy
            wsInfo.Cells(34, valueColumn).Resize(2, 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False

            Dim baseAmount As Double
            baseAmount = 0
            Dim baseValue As Variant
            baseValue = wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, valueColumn).value
            If IsNumeric(baseValue) And Len(Trim$(CStr(baseValue))) > 0 Then
                baseAmount = CDbl(baseValue)
            End If

            Dim taxAmount As Double
            taxAmount = RoundDownAmount(baseAmount * taxRate)

            wsInfo.Cells(34, valueColumn).value = taxAmount
            wsInfo.Cells(35, valueColumn).value = baseAmount + taxAmount
        Else
            ' ��А��������͑ΏۊO�u���b�N��34/35�s�ڂ̒l�̂ݏ�������B
            ' .Clear �͓h��Ԃ��܂ŉ������邽�ߎg��Ȃ�(���g�p�u���b�N�� #06111D ���ێ�)�B
            ClearInactiveVendorBlockTaxRows wsInfo, labelColumn, valueColumn
        End If
    Next i
    Application.CutCopyMode = False
    Exit Sub

ErrorHandler:
    Application.CutCopyMode = False
    LogCI "�u���b�N����ōs�X�V�G���[ Err " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

' �ΏۊO�{�H��Ѓu���b�N��34/35�s��: �l�̂ݏ������A���g�p�u���b�N�Ɠ����w�i�F�𕜌�����B
Private Sub ClearInactiveVendorBlockTaxRows(ByVal wsInfo As Worksheet, _
                                            ByVal labelColumn As Long, _
                                            ByVal valueColumn As Long)
    Dim clearRange As Range
    Set clearRange = wsInfo.Range(wsInfo.Cells(34, labelColumn), _
                                  wsInfo.Cells(35, valueColumn))
    clearRange.ClearContents
    clearRange.Interior.Color = RGB(6, 17, 29)
End Sub

Public Sub ClearVendorAliasMapCacheCore()
    Set mVendorAliasMapCache = Nothing
End Sub

Public Function GetVendorAliasMap(ByVal branchName As String) As Object
    Dim cacheKey As String
    cacheKey = CommonNormalizeText(branchName)

    If mVendorAliasMapCache Is Nothing Then
        Set mVendorAliasMapCache = CreateObject("Scripting.Dictionary")
        mVendorAliasMapCache.CompareMode = vbTextCompare
    End If
    If cacheKey <> "" Then
        If mVendorAliasMapCache.Exists(cacheKey) Then
            Dim cachedAliasMap As Object
            Set cachedAliasMap = mVendorAliasMapCache(cacheKey)
            If Not cachedAliasMap Is Nothing Then
                Set GetVendorAliasMap = cachedAliasMap
                Exit Function
            End If
        End If
    End If

    Dim aliasMap As Object
    Set aliasMap = BuildVendorAliasMap(branchName)
    If cacheKey <> "" Then
        If mVendorAliasMapCache.Exists(cacheKey) Then
            Set mVendorAliasMapCache(cacheKey) = aliasMap
        Else
            mVendorAliasMapCache.Add cacheKey, aliasMap
        End If
    End If
    Set GetVendorAliasMap = aliasMap
End Function

Public Function BuildSheetVendorAmountColumnMap(ByVal ws As Worksheet, _
                                                 ByVal aliasMap As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim kindColumn As Long
    kindColumn = FindHeaderColumn(ws, "�H�핪��")
    If kindColumn <= mod_Construction_OutputLayout.OutputSheetSubconPriceFirstColCore(ws) Then
        Set BuildSheetVendorAmountColumnMap = result
        Exit Function
    End If

    Dim subconFirstCol As Long
    subconFirstCol = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstColCore(ws)

    Dim c As Long
    For c = subconFirstCol To kindColumn - 1
        Dim headerText As String
        headerText = CommonNzText(ws.Cells(1, c).value)
        If Len(headerText) > Len("���z") Then
            If Right$(headerText, Len("���z")) = "���z" Then
                Dim vendorKey As String
                vendorKey = ResolveVendorCanonicalKey(Left$(headerText, Len(headerText) - Len("���z")), aliasMap)
                If vendorKey <> "" Then
                    If Not result.Exists(vendorKey) Then
                        result.Add vendorKey, c
                    End If
                End If
            End If
        End If
    Next c

    Set BuildSheetVendorAmountColumnMap = result
End Function

Public Function SumVendorAmountByColumn(ByVal ws As Worksheet, _
                                         ByVal amountColumn As Long) As Double
    Dim SeiriColumn As Long
    SeiriColumn = FindHeaderColumn(ws, "�����ԍ�")
    If SeiriColumn = 0 Or amountColumn = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws, SeiriColumn)
    If lastRow < 2 Then Exit Function

    SumVendorAmountByColumn = RoundDownAmount(SumNumericColumn(ws, amountColumn, lastRow))
End Function

Public Function BasicInfoVendorColumn(ByVal vendorIndex As Long) As Long
    BasicInfoVendorColumn = BASIC_INFO_VENDOR_FIRST_COL + _
                            ((vendorIndex - 1) * BASIC_INFO_VENDOR_STEP_COLS)
End Function

Public Function GetBasicInfoVendorBlockCount(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CommonNzText( _
        wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > BASIC_INFO_VENDOR_MAX_BLOCKS Then countValue = BASIC_INFO_VENDOR_MAX_BLOCKS
    GetBasicInfoVendorBlockCount = countValue
End Function

Public Function IsPurchaseOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    IsPurchaseOutputSheet = _
        (StrComp(ws.Name, CommonPurchaseOrderOutputSheetName(), vbTextCompare) = 0) Or _
        (StrComp(ws.Name, CommonPurchaseNoticeOutputSheetName(), vbTextCompare) = 0)
End Function

Public Function IsConstructionOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If IsPurchaseOutputSheet(ws) Then Exit Function

    IsConstructionOutputSheet = _
        ((FindHeaderColumn(ws, "�{�H�Ǝ�") > 0) Or mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws)) And _
        (FindHeaderColumn(ws, "�����ԍ�") > 0) And _
        (FindHeaderColumn(ws, "JR���z") > 0) And _
        (FindHeaderColumn(ws, "�H�핪��") > 0)
End Function

Public Function SumOutputJrAmount(ByVal ws As Worksheet) As Double
    Dim SeiriColumn As Long
    Dim amountColumn As Long
    SeiriColumn = FindHeaderColumn(ws, "�����ԍ�")
    amountColumn = FindHeaderColumn(ws, "JR���z")
    If SeiriColumn = 0 Or amountColumn = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws, SeiriColumn)
    If lastRow < 2 Then Exit Function

    SumOutputJrAmount = RoundDownAmount(SumNumericColumn(ws, amountColumn, lastRow))
End Function
' �{�s�w����(�H��)/�{�s�ʒm��(�H��)�V�[�g�̂����A�Y�p�s(�{�H�Ǝ҂��I���ł��Ȃ�����s)��
' JR���z�̂ݍ��v����(�ʎ�III J54:M54 �p)�B�n�ڃV�[�g�͑ΏۊO�B
Public Function SumSanpaiJrAmount() As Double
    Dim total As Double
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsConstructionOutputSheet(ws) Then
            If Not mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                total = total + SumSanpaiJrAmountOnSheet(ws)
            End If
        End If
    Next ws
    SumSanpaiJrAmount = RoundDownAmount(total)
End Function

Private Function SumSanpaiJrAmountOnSheet(ByVal ws As Worksheet) As Double
    Dim SeiriColumn As Long
    Dim amountColumn As Long
    SeiriColumn = FindHeaderColumn(ws, "�����ԍ�")
    amountColumn = FindHeaderColumn(ws, "JR���z")
    If SeiriColumn = 0 Or amountColumn = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws, SeiriColumn)
    If lastRow < 2 Then Exit Function

    Dim subtotal As Double
    Dim r As Long
    For r = 2 To lastRow
        If IsSanpaiRow(ws, r) Then
            Dim v As Variant
            v = ws.Cells(r, amountColumn).Value2
            If IsNumeric(v) Then subtotal = subtotal + CDbl(v)
        End If
    Next r
    SumSanpaiJrAmountOnSheet = subtotal
End Function

Public Function SumVendorAmountOnSheet(ByVal ws As Worksheet, _
                                        ByVal vendorName As String, _
                                        ByVal aliasMap As Object) As Double
    Dim vendorKey As String
    vendorKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
    If vendorKey = "" Then Exit Function

    Dim columnMap As Object
    Set columnMap = BuildSheetVendorAmountColumnMap(ws, aliasMap)
    If columnMap.Exists(vendorKey) Then
        SumVendorAmountOnSheet = SumVendorAmountByColumn(ws, CLng(columnMap(vendorKey)))
    End If
End Function

' �{�s�w����(�H��)/�{�s�ʒm��(�H��)�� A��(�{�H�Ǝ�)�ɁA�w��Ǝ҂��I������Ă��邩�B
' (�ʎ�III��JR���v���u�H���V�[�g�ɕR�t�����{�H��Ёv�Ɍ��肷�邽�߂̔���B�n�ڂ̂ݓ��� False)
Public Function IsVendorSelectedOnWorksSheet(ByVal branchName As String, _
                                             ByVal vendorName As String) As Boolean
    On Error GoTo Done
    Dim wsWorks As Worksheet
    Set wsWorks = mod_OrderTpl_Shared.OrderTplFindWorksSourceSheet()
    If wsWorks Is Nothing Then Exit Function

    Dim aliasMap As Object
    Set aliasMap = GetVendorAliasMap(branchName)

    Dim vendorKey As String
    vendorKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
    If vendorKey = "" Then Exit Function

    Dim columnMap As Object
    Set columnMap = BuildSheetVendorAmountColumnMap(wsWorks, aliasMap)
    If columnMap Is Nothing Then Exit Function
    IsVendorSelectedOnWorksSheet = columnMap.Exists(vendorKey)

Done:
End Function

Public Function SumNumericColumn(ByVal ws As Worksheet, _
                                  ByVal targetColumn As Long, _
                                  ByVal lastRow As Long) As Double
    Dim totalAmount As Double
    Dim r As Long
    For r = 2 To lastRow
        Dim cellValue As Variant
        cellValue = ws.Cells(r, targetColumn).value
        If Not IsError(cellValue) Then
            If IsNumeric(cellValue) Then totalAmount = totalAmount + CDbl(cellValue)
        End If
    Next r

    SumNumericColumn = totalAmount
End Function

Public Function RoundDownAmount(ByVal amount As Double) As Double
    RoundDownAmount = Fix(amount)
End Function

Public Function GetBasicInfoCellText(ByVal wsInfo As Worksheet, _
                                      ByVal cellAddress As String) As String
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.mergeArea.Cells(1, 1)
    GetBasicInfoCellText = Trim$(CommonNzText(targetCell.value))
End Function

Public Sub WriteBasicInfoAmount(ByVal wsInfo As Worksheet, _
                                 ByVal cellAddress As String, _
                                 ByVal amount As Double, _
                                 Optional ByVal hasValue As Boolean = True)
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.mergeArea.Cells(1, 1)

    targetCell.NumberFormatLocal = BasicInfoYenNumberFormat()
    If hasValue Then
        targetCell.value = RoundDownAmount(amount)
    Else
        targetCell.ClearContents
    End If
End Sub

'  UpdateBasicInfoTaxTotals
'  ��{���V�[�g�� C33(���v)�EC34(�����)�EC35(�ō����v) ���X�V���A
'  C31:C35 �Ɂu\�{����؂�v�̕\���`����K�p����B
'  C33 = C31 + C32
'  C34 = C33 �~ �ŗ�(B34�̕\�L����擾�B�擾�ł��Ȃ��ꍇ��10%) �������_�ȉ��؂�̂�
'  C35 = C33 + C34
Public Sub UpdateBasicInfoTaxTotalsCore(Optional ByVal wsInfo As Worksheet)
    On Error GoTo ErrorHandler

    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim subtotal As Double
    subtotal = GetBasicInfoCellAmount(wsInfo, BASIC_INFO_WORKS_TOTAL_CELL) + _
               GetBasicInfoCellAmount(wsInfo, BASIC_INFO_PURCHASE_TOTAL_CELL)

    Dim taxAmount As Double
    taxAmount = Fix(subtotal * ResolveBasicInfoTaxRate(wsInfo))

    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_SUBTOTAL_CELL, subtotal
    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_TAX_CELL, taxAmount
    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_GRAND_TOTAL_CELL, subtotal + taxAmount

    ApplyBasicInfoYenTotalFormat wsInfo

    LogCI "�ō����v�X�V: ���v=" & subtotal & " / �����=" & taxAmount & _
          " / �ō����v=" & (subtotal + taxAmount)
    Exit Sub

ErrorHandler:
    LogCI "��{���ō����v�X�V�G���[ Err " & Err.Number & ": " & Err.Description
End Sub

'  GetBasicInfoCellAmount
'  �w��Z���̐��l���擾����(�����Z���Ή��B�󗓁E�񐔒l�E�G���[�l��0)�B
Public Function GetBasicInfoCellAmount(ByVal wsInfo As Worksheet, _
                                        ByVal cellAddress As String) As Double
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.mergeArea.Cells(1, 1)

    Dim cellValue As Variant
    cellValue = targetCell.value
    If Not IsError(cellValue) Then
        If IsNumeric(cellValue) Then GetBasicInfoCellAmount = CDbl(cellValue)
    End If
End Function

'  WriteBasicInfoPlainValue
'  �w��Z���֒l�݂̂���������(�����Z���Ή��B�\���`���͕ύX���Ȃ�)�B
Public Sub WriteBasicInfoPlainValue(ByVal wsInfo As Worksheet, _
                                     ByVal cellAddress As String, _
                                     ByVal amount As Double)
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.mergeArea.Cells(1, 1)
    targetCell.value = amount
End Sub

'  ResolveBasicInfoTaxRate
'  B34 �̃��x��(��:�u�����(10%)�v)����ŗ��𒊏o����B
'  �u%�v���O�̐��l��ŗ��Ƃ��ĉ��߂��A�擾�ł��Ȃ��ꍇ�͊����10%��Ԃ��B
Public Function ResolveBasicInfoTaxRate(ByVal wsInfo As Worksheet) As Double
    ResolveBasicInfoTaxRate = BASIC_INFO_TAX_RATE_DEFAULT

    Dim labelText As String
    On Error Resume Next
    labelText = StrConv(CommonNzText(wsInfo.Range(BASIC_INFO_TAX_LABEL_CELL).value), vbNarrow)
    On Error GoTo 0
    If labelText = "" Then Exit Function

    Dim numText As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(labelText)
        ch = Mid$(labelText, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            numText = numText & ch
        ElseIf ch = "%" Then
            If numText <> "" And IsNumeric(numText) Then
                ResolveBasicInfoTaxRate = CDbl(numText) / 100
            End If
            Exit Function
        Else
            numText = ""
        End If
    Next i
End Function

'  ApplyBasicInfoYenTotalFormat
'  C31:C35 �Ɂu\�{����؂�v(�����͐Ԏ�)�̕\���`����K�p����B
Public Sub ApplyBasicInfoYenTotalFormat(ByVal wsInfo As Worksheet)
    wsInfo.Range(BASIC_INFO_YEN_TOTAL_RANGE).NumberFormatLocal = BasicInfoYenNumberFormat()
End Sub

'  BasicInfoYenNumberFormat
'  �u\�{����؂�v(�����͐Ԏ�)�̕\���`���������Ԃ��B
'  \�L����CP932�ł̕�������������邽�� ChrW$ �Ő�������B
Public Function BasicInfoYenNumberFormat() As String
    Dim yenMark As String
    yenMark = ChrW$(&HA5)   ' \

    BasicInfoYenNumberFormat = yenMark & "#,##0;[��]-" & yenMark & "#,##0"
End Function

Public Function CollectSelectedSubcontractors(ByVal ws As Worksheet, _
                                                ByVal lastRow As Long) As Collection
    Dim result As New Collection
    If lastRow < 2 Then
        Set CollectSelectedSubcontractors = result
        Exit Function
    End If

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)

    Dim aliasMap As Object
    Set aliasMap = Nothing
    If Not wsInfo Is Nothing Then
        Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))
    End If

    Dim sheetByCanonical As Object
    Set sheetByCanonical = CreateObject("Scripting.Dictionary")
    sheetByCanonical.CompareMode = vbTextCompare

    Dim sheetOrderKeys As New Collection

    Dim vendorColumns As Collection
    Set vendorColumns = mod_Construction_OutputLayout.OutputSheetVendorColumnsCore(ws)

    Dim r As Long
    Dim vendorCol As Variant
    For r = 2 To lastRow
        For Each vendorCol In vendorColumns
            Dim vendorName As String
            Dim canonicalKey As String
            vendorName = Trim$(CommonNzText(ws.Cells(r, CLng(vendorCol)).value))
            If vendorName = "" Then GoTo NextVendorCol

            canonicalKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
            If canonicalKey = "" Then GoTo NextVendorCol

            If Not sheetByCanonical.Exists(canonicalKey) Then
                sheetByCanonical.Add canonicalKey, vendorName
                sheetOrderKeys.Add canonicalKey
            End If
NextVendorCol:
        Next vendorCol
    Next r

    If sheetByCanonical.Count = 0 Then
        Set CollectSelectedSubcontractors = result
        Exit Function
    End If

    Dim usedCanonical As Object
    Set usedCanonical = CreateObject("Scripting.Dictionary")
    usedCanonical.CompareMode = vbTextCompare

    If Not wsInfo Is Nothing Then
        If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
            AppendOrderedBasicInfoVendorsByWorkType wsInfo, aliasMap, sheetByCanonical, usedCanonical, result, WELDING_WORK_TYPE_KEYWORD
            AppendOrderedBasicInfoVendorsByWorkType wsInfo, aliasMap, sheetByCanonical, usedCanonical, result, TRACK_WORK_TYPE_KEYWORD
        Else
            Dim vendorCount As Long
            Dim i As Long
            vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

            For i = 1 To vendorCount
                Dim basicInfoName As String
                Dim basicInfoKey As String
                basicInfoName = GetBasicInfoCellText(wsInfo, _
                    wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, BasicInfoVendorColumn(i)).Address)
                If basicInfoName = "" Then GoTo NextBasicInfoVendor

                basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
                If basicInfoKey <> "" And sheetByCanonical.Exists(basicInfoKey) Then
                    If Not usedCanonical.Exists(basicInfoKey) Then
                        usedCanonical.Add basicInfoKey, True
                        result.Add CStr(sheetByCanonical(basicInfoKey))
                    End If
                End If
NextBasicInfoVendor:
            Next i
        End If
    End If

    Dim fallbackKey As Variant
    For Each fallbackKey In sheetOrderKeys
        If Not usedCanonical.Exists(CStr(fallbackKey)) Then
            result.Add CStr(sheetByCanonical(CStr(fallbackKey)))
        End If
    Next fallbackKey

    Set CollectSelectedSubcontractors = result
End Function

Public Sub AppendOrderedBasicInfoVendorsByWorkType( _
    ByVal wsInfo As Worksheet, _
    ByVal aliasMap As Object, _
    ByVal sheetByCanonical As Object, _
    ByVal usedCanonical As Object, _
    ByVal result As Collection, _
    ByVal workTypeKeyword As String)

    Dim vendorNameMap As Object
    Set vendorNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)
    If vendorNameMap Is Nothing Then Exit Sub

    Dim blockIndex As Long
    For blockIndex = 1 To BASIC_INFO_VENDOR_MAX_BLOCKS
        Dim valueCol As Long
        valueCol = BasicInfoVendorColumn(blockIndex)
        If Not BasicInfoBlockMatchesWorkType(wsInfo, valueCol, workTypeKeyword) Then GoTo NextWorkTypeBlock

        Dim basicInfoName As String
        Dim basicInfoKey As String
        Dim mappedName As String
        basicInfoName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).Address)
        If basicInfoName = "" Then GoTo NextWorkTypeBlock

        basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
        If basicInfoKey = "" Then GoTo NextWorkTypeBlock
        If Not sheetByCanonical.Exists(basicInfoKey) Then GoTo NextWorkTypeBlock
        If usedCanonical.Exists(basicInfoKey) Then GoTo NextWorkTypeBlock

        mappedName = Trim$(CommonNzText(sheetByCanonical(basicInfoKey)))
        If mappedName = "" Then GoTo NextWorkTypeBlock

        usedCanonical.Add basicInfoKey, True
        result.Add mappedName
NextWorkTypeBlock:
    Next blockIndex
End Sub

Public Function BasicInfoBlockMatchesWorkType(ByVal wsInfo As Worksheet, _
                                               ByVal valueCol As Long, _
                                               ByVal workTypeKeyword As String) As Boolean
    Dim workTypeText As String
    workTypeText = CommonRemoveAllSpaces(CommonNormalizeText( _
        CommonNzText(wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueCol).value)))
    BasicInfoBlockMatchesWorkType = (workTypeText <> "") And _
        (InStr(1, workTypeText, workTypeKeyword, vbTextCompare) > 0)
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim hit As Range
    Set hit = ws.rows(1).Find(What:=headerText, After:=ws.Cells(1, 1), _
                              LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                              MatchCase:=False)
    If Not hit Is Nothing Then FindHeaderColumn = hit.Column
End Function

Public Function GetVendorUnitPriceRows(ByVal unitPriceSheetName As String, _
                                        ByVal vendorName As String, _
                                        ByVal vendorPriceCaches As Object) As Object
    If unitPriceSheetName = "" Then Exit Function
    If vendorPriceCaches Is Nothing Then Exit Function

    Dim cacheKey As String
    cacheKey = unitPriceSheetName & "|" & NormalizeVendorPriceName(vendorName)
    If vendorPriceCaches.Exists(cacheKey) Then
        Dim cachedRows As Object
        Set cachedRows = vendorPriceCaches(cacheKey)
        If Not cachedRows Is Nothing Then
            Set GetVendorUnitPriceRows = cachedRows
            Exit Function
        End If
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(unitPriceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then
        LogCI "�{�H��ВP��: �V�[�g�����o [" & unitPriceSheetName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If
    If Not mod_MaterialPriceImport.IsConstructionUnitPriceSheet(priceSheet) Then
        LogCI "�{�H��ВP��: �H���P���V�[�g�ł͂Ȃ� [" & unitPriceSheetName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim aliasMap As Object
    Set aliasMap = Nothing
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If Not wsInfo Is Nothing Then
        Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))
    End If

    Dim vendorDayColumn As Long
    vendorDayColumn = FindUnitPriceVendorDayColumn(priceSheet, vendorName, aliasMap)
    If vendorDayColumn = 0 Then
        LogCI "�{�H��ВP��: �Ǝҗ񖢌��o sheet=[" & unitPriceSheetName & "] vendor=[" & vendorName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim priceLastRow As Long
    priceLastRow = GetUnitPriceSheetLastDataRow(priceSheet)

    Dim r As Long
    For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
        Dim recordKey As String
        recordKey = mod_Construction_LineMapping.NormalizeRecordKey(priceSheet.Cells(r, COL_SEIRI).value)
        If recordKey = "" Then GoTo NextVendorPriceRow
        If Not IsUnitPriceVendorRowPriceEligible(priceSheet, r, vendorDayColumn) Then GoTo NextVendorPriceRow

        If Not result.Exists(recordKey) Then
            result.Add recordKey, Array(priceSheet.Cells(r, vendorDayColumn).Value2, _
                                        priceSheet.Cells(r, vendorDayColumn + 1).Value2)
        End If
NextVendorPriceRow:
    Next r

    LogCI "�{�H��ВP��: sheet=[" & unitPriceSheetName & "] vendor=[" & vendorName & _
          "] col=" & vendorDayColumn & " keys=" & result.Count

    StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
    Set GetVendorUnitPriceRows = result
End Function

Public Sub StoreVendorUnitPriceCache(ByVal vendorPriceCaches As Object, _
                                      ByVal cacheKey As String, _
                                      ByVal priceRows As Object)
    If vendorPriceCaches Is Nothing Then Exit Sub
    If vendorPriceCaches.Exists(cacheKey) Then
        Set vendorPriceCaches(cacheKey) = priceRows
    Else
        vendorPriceCaches.Add cacheKey, priceRows
    End If
End Sub

Public Function FindUnitPriceVendorDayColumn(ByVal priceSheet As Worksheet, _
                                              ByVal vendorName As String, _
                                              Optional ByVal aliasMap As Object = Nothing) As Long
    Dim vendorKey As String
    vendorKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
    If vendorKey = "" Then Exit Function

    Dim lastColumn As Long
    lastColumn = priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, _
                                  priceSheet.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = UNIT_PRICE_VENDOR_FIRST_DAY_COL To lastColumn Step 2
        Dim headerCell As Range
        Dim headerKey As String
        Set headerCell = priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, c)
        On Error Resume Next
        If headerCell.MergeCells Then
            Dim mergedHeader As Range
            Set mergedHeader = headerCell.mergeArea.Cells(1, 1)
            If Not mergedHeader Is Nothing Then Set headerCell = mergedHeader
        End If
        On Error GoTo 0
        headerKey = ResolveVendorCanonicalKey(CommonNzText(headerCell.value), aliasMap)
        If headerKey <> "" Then
            If StrComp(headerKey, vendorKey, vbTextCompare) = 0 Then
                FindUnitPriceVendorDayColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Public Function GetUnitPriceSheetLastDataRow(ByVal priceSheet As Worksheet) As Long
    Dim scanStartRow As Long
    scanStartRow = UNIT_PRICE_DATA_START_ROW
    If Not priceSheet.UsedRange Is Nothing Then
        Dim usedLastRow As Long
        usedLastRow = priceSheet.UsedRange.row + priceSheet.UsedRange.rows.Count - 1
        If usedLastRow > scanStartRow Then scanStartRow = usedLastRow
    End If

    Dim rowIndex As Long
    For rowIndex = scanStartRow To UNIT_PRICE_DATA_START_ROW Step -1
        If Trim$(CommonNzText(priceSheet.Cells(rowIndex, COL_SEIRI).value)) <> "" Then
            GetUnitPriceSheetLastDataRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    GetUnitPriceSheetLastDataRow = UNIT_PRICE_DATA_START_ROW - 1
End Function

Public Function IsSanpaiTypeText(ByVal typeText As String) As Boolean
    IsSanpaiTypeText = (InStr(1, CommonRemoveAllSpaces(CommonNormalizeText(typeText)), _
                                SANPAI_KEYWORD, vbTextCompare) > 0)
End Function

Public Function UnitPriceValueIsUsable(ByVal value As Variant) As Boolean
    If isEmpty(value) Or IsError(value) Then Exit Function
    If Len(Trim$(CommonNzText(value))) = 0 Then Exit Function
    UnitPriceValueIsUsable = IsNumeric(value)
End Function

Public Function IsUnitPriceSheetSanpaiRow(ByVal priceSheet As Worksheet, _
                                           ByVal rowIndex As Long) As Boolean
    IsUnitPriceSheetSanpaiRow = IsSanpaiTypeText( _
        CommonNzText(priceSheet.Cells(rowIndex, UNIT_PRICE_WORK_TYPE_COL).value))
End Function

Public Function IsUnitPriceVendorRowPriceEligible(ByVal priceSheet As Worksheet, _
                                                   ByVal rowIndex As Long, _
                                                   ByVal vendorDayColumn As Long) As Boolean
    If IsUnitPriceSheetSanpaiRow(priceSheet, rowIndex) Then Exit Function

    Dim dayPrice As Variant
    Dim nightPrice As Variant
    dayPrice = priceSheet.Cells(rowIndex, vendorDayColumn).Value2
    nightPrice = priceSheet.Cells(rowIndex, vendorDayColumn + 1).Value2

    '  �Ǝ��H��Ȃ� JR �Q��(E/F)����ł��A�{�H��З�֎���͂��ꂽ�P���͍̗p����
    IsUnitPriceVendorRowPriceEligible = UnitPriceValueIsUsable(dayPrice) Or _
                                        UnitPriceValueIsUsable(nightPrice)
End Function

Public Function SelectUsableDayNightPrice(ByVal dayNightText As String, _
                                             ByVal dayNightPrices As Variant) As Variant
    Dim selectedPrice As Variant
    selectedPrice = mod_Construction_LineMapping.SelectDayNightPrice(dayNightText, dayNightPrices)
    If UnitPriceValueIsUsable(selectedPrice) Then
        SelectUsableDayNightPrice = selectedPrice
    End If
End Function

Public Function NormalizeVendorPriceName(ByVal vendorName As String) As String
    NormalizeVendorPriceName = CommonRemoveAllSpaces(CommonNormalizeText(vendorName))
End Function

'  ResolveVendorCanonicalKey
'  �Ǝ҃}�X�^(�ʖ��\)�𐳂Ƃ��āA���K���E���̖��̂ǂ���̕\�L�ł�
'  ����̐��K���L�[�։�������B�}�X�^�ɖ������̂͐��K������������̂܂ܕԂ�
'  ���߃t�H�[���o�b�N����A�Q�ƃG���[�ɂ͂Ȃ�Ȃ��B
Public Function ResolveVendorCanonicalKey(ByVal vendorName As String, _
                                           ByVal aliasMap As Object) As String
    Dim normalizedKey As String
    normalizedKey = NormalizeVendorPriceName(vendorName)
    If normalizedKey = "" Then Exit Function

    If Not aliasMap Is Nothing Then
        If aliasMap.Exists(normalizedKey) Then
            ResolveVendorCanonicalKey = CStr(aliasMap(normalizedKey))
            Exit Function
        End If
    End If

    ResolveVendorCanonicalKey = normalizedKey
End Function

Public Function ResolveBasicInfoVendorInfoIndexCore(ByVal vendorDisplayName As String, _
                                                Optional ByVal workTypeKeyword As String = "") As Long
    ResolveBasicInfoVendorInfoIndexCore = 0

    Dim normalizedDisplay As String
    normalizedDisplay = NormalizeVendorPriceName(vendorDisplayName)
    If normalizedDisplay = "" Then Exit Function

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim aliasMap As Object
    Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))

    Dim canonicalKey As String
    canonicalKey = ResolveVendorCanonicalKey(vendorDisplayName, aliasMap)

    Dim vendorNameMap As Object
    Set vendorNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueCol As Long
        valueCol = BasicInfoVendorColumn(vendorIndex)

        If workTypeKeyword <> "" Then
            If Not BasicInfoBlockMatchesWorkType(wsInfo, valueCol, workTypeKeyword) Then GoTo NextVendorIndex
        End If

        Dim basicInfoName As String
        basicInfoName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).Address)
        If basicInfoName = "" Then GoTo NextVendorIndex

        Dim basicInfoKey As String
        basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
        If canonicalKey <> "" And basicInfoKey = canonicalKey Then
            ResolveBasicInfoVendorInfoIndexCore = vendorIndex
            Exit Function
        End If

        If NormalizeVendorPriceName(basicInfoName) = normalizedDisplay Then
            ResolveBasicInfoVendorInfoIndexCore = vendorIndex
            Exit Function
        End If

        If Not vendorNameMap Is Nothing Then
            Dim nameKey As String
            nameKey = CommonNormalizeText(basicInfoName)
            If vendorNameMap.Exists(nameKey) Then
                Dim mappedDisplayName As String
                mappedDisplayName = Trim$(CommonNzText(vendorNameMap(nameKey)))
                If NormalizeVendorPriceName(mappedDisplayName) = normalizedDisplay Then
                    ResolveBasicInfoVendorInfoIndexCore = vendorIndex
                    Exit Function
                End If
            End If
        End If
NextVendorIndex:
    Next vendorIndex
End Function

'  BuildVendorAliasMap
'  �Ǝ҃}�X�^(�S�Д�).xlsx �́u�x�X��(��{���B6)�v�V�[�g���J���A
'  A��=�ƎҖ�(����) / B��=�����Ҏ���(���K��) ��ǂݍ���ŁA
'  ���K��(����)�E���K��(���K��) �̑o���� ���K��(���K��) �֑Ή��t����������Ԃ��B
'  (1�s�ڂ͌��o���s�����A���ƎҖ��ƈ�v���Ȃ����Q�ȃG���g���ɂȂ邾��)
'  �}�X�^�����o�E�V�[�g�����o�E�Ǎ����s���͋󎫏���Ԃ�(�ˍ��͐��K���݂̂Ōp��)�B
Public Function BuildVendorAliasMap(ByVal branchName As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim connection As Object
    Dim recordset As Object

    On Error GoTo Cleanup

    If Trim$(branchName) = "" Then
        LogCI "�Ǝ҃}�X�^�ʖ�: ��{���B6(�x�X��)����̂��ߖ��񂹂Ȃ�"
        GoTo Cleanup
    End If

    Dim masterPath As String
    masterPath = ResolveVendorMasterPath()
    If masterPath = "" Then
        LogCI "�Ǝ҃}�X�^�����o -> ���񂹂Ȃ�(���K���݂̂œˍ�)"
        GoTo Cleanup
    End If

    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then
        LogCI "�Ǝ҃}�X�^ADO�ڑ��s�� path=[" & masterPath & "]"
        GoTo Cleanup
    End If

    Dim actualSheetName As String
    actualSheetName = mod_Construction_OutputLayout.FindAdoWorksheetName(connection, branchName)
    If actualSheetName = "" Then
        LogCI "�Ǝ҃}�X�^�Ɏx�X�V�[�g[" & branchName & "]��������܂��� -> ���񂹂Ȃ�"
        GoTo Cleanup
    End If

    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT [F" & VENDOR_MASTER_OFFICIAL_COL & "], [F" & VENDOR_MASTER_ABBREV_COL & "] FROM " & _
                   mod_Construction_OutputLayout.BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

    Dim official As String, abbrev As String, canonicalKey As String
    Do Until recordset.EOF
        official = CommonNzText(recordset.fields(0).value)
        abbrev = CommonNzText(recordset.fields(1).value)
        canonicalKey = NormalizeVendorPriceName(official)
        If canonicalKey <> "" Then
            AddVendorAlias result, official, canonicalKey
            AddVendorAlias result, abbrev, canonicalKey
        End If
        recordset.MoveNext
    Loop

    LogCI "�Ǝ҃}�X�^�ʖ� ����=" & result.Count & " �x�X=[" & branchName & _
          "] sheet=[" & actualSheetName & "]"

Cleanup:
    If Err.Number <> 0 Then
        LogCI "�Ǝ҃}�X�^�Ǎ��G���[ Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    Set BuildVendorAliasMap = result
End Function

'  AddVendorAlias
'  ���K�������ʖ��𐳋K���L�[�֓o�^����(��E�d���͖���)�B
Public Sub AddVendorAlias(ByVal aliasMap As Object, _
                           ByVal aliasName As String, _
                           ByVal canonicalKey As String)
    Dim normalizedAlias As String
    normalizedAlias = NormalizeVendorPriceName(aliasName)
    If normalizedAlias = "" Then Exit Sub
    If Not aliasMap.Exists(normalizedAlias) Then aliasMap.Add normalizedAlias, canonicalKey
End Sub

'  ResolveVendorMasterPath
'  �Ǝ҃}�X�^(�S�Д�).xlsx �̃p�X�𕡐���₩���������B
Public Function ResolveVendorMasterPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim candidates As Collection
    Set candidates = New Collection

    Dim unitMasterPath As String
    unitMasterPath = mod_Construction_Import_Load.ResolveMasterFilePath()
    If unitMasterPath <> "" Then
        AddUniqueText candidates, _
            fso.BuildPath(fso.GetParentFolderName(unitMasterPath), VENDOR_MASTER_FILE)
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        AddUniqueText candidates, _
            fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                          MASTER_DATA_FOLDER & "\" & VENDOR_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText candidates, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
            "���H�o�����p_������_�������A�N�Z�X�T�C�g - �h�L�������g\" & _
            MASTER_DATA_FOLDER & "\" & VENDOR_MASTER_FILE
    End If

    Dim candidate As Variant
    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            ResolveVendorMasterPath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Public Sub ApplyWeldingOutputSheetColumnAlignment(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then Exit Sub

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    If lastRow < 1 Then lastRow = 1

    ws.Columns(WELD_COL_WELDING_VENDOR).HorizontalAlignment = xlCenter
    ws.Columns(WELD_COL_TRACK_VENDOR).HorizontalAlignment = xlCenter
    ws.Columns("F").HorizontalAlignment = xlCenter
    ws.Columns("I").HorizontalAlignment = xlCenter
    ws.Columns("R").HorizontalAlignment = xlCenter
End Sub

Public Sub FormatSubcontractorPriceColumns(ByVal ws As Worksheet, _
                                            ByVal lastRow As Long, _
                                            ByVal columnCount As Long, _
                                            Optional ByVal firstColumn As Long = 0)
    Dim lastColumn As Long
    If firstColumn = 0 Then firstColumn = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstColCore(ws)
    lastColumn = firstColumn + columnCount - 1

    With ws.Range(ws.Cells(1, firstColumn), ws.Cells(1, lastColumn))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .shrinkToFit = True
    End With

    If lastRow >= 2 Then
        With ws.Range(ws.Cells(2, firstColumn), ws.Cells(lastRow, lastColumn))
            .NumberFormatLocal = "#,##0;[��]-#,##0"
        End With
        With ws.Range(ws.Cells(1, firstColumn), ws.Cells(lastRow, lastColumn)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
    End If

    ws.Range(ws.Columns(firstColumn), ws.Columns(lastColumn)).AutoFit
End Sub

Public Function IsSanpaiRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    IsSanpaiRow = IsSanpaiTypeText( _
        CommonNzText(ws.Cells(rowIndex, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_TYPE)).value))
End Function

' ��ʗ���ꊇ�ǎ悵�A2�s�ځ`lastRow�̎Y�p�s�t���O�z��(�Y��=���s�ԍ�)��Ԃ��B
' �s�P�ʂ� IsSanpaiRow �A��(����Z���ǎ�)������邽�߂̈ꊇ�ŁB
Public Function BuildSanpaiRowFlags(ByVal ws As Worksheet, ByVal lastRow As Long) As Variant
    Dim flags() As Boolean
    If ws Is Nothing Or lastRow < 2 Then
        ReDim flags(2 To 2)
        BuildSanpaiRowFlags = flags
        Exit Function
    End If
    ReDim flags(2 To lastRow)

    Dim typeCol As Long
    typeCol = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_TYPE)

    Dim typeVals As Variant
    typeVals = ws.Range(ws.Cells(2, typeCol), ws.Cells(lastRow, typeCol)).Value2
    If Not IsArray(typeVals) Then
        flags(2) = IsSanpaiTypeText(CommonNzText(typeVals))
        BuildSanpaiRowFlags = flags
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        flags(r) = IsSanpaiTypeText(CommonNzText(typeVals(r - 1, 1)))
    Next r
    BuildSanpaiRowFlags = flags
End Function

Public Function GetSanpaiFillColor() As Long
    If mSanpaiFillColorCached Then
        GetSanpaiFillColor = mSanpaiFillColorCache
        Exit Function
    End If

    mSanpaiFillColorCache = SANPAI_FALLBACK_FILL_COLOR

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(ws) Then
            Dim priceLastRow As Long
            priceLastRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).row

            Dim r As Long, c As Long
            For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
                For c = 5 To 6
                    If ws.Cells(r, c).Interior.Pattern <> xlPatternNone Then
                        mSanpaiFillColorCache = ws.Cells(r, c).Interior.Color
                        GoTo SanpaiFillColorDone
                    End If
                Next c
            Next r
        End If
    Next ws

SanpaiFillColorDone:
    mSanpaiFillColorCached = True
    GetSanpaiFillColor = mSanpaiFillColorCache
End Function
