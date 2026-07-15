Option Explicit

' �������e���v���[�g�捞(�{�H��Њm�莞�̃V�[�g�}���E���󖾍ד]�L)�̋��ʒ萔�E�ƍ��w���p�[�B
' ���C����: CHANGELOG.md �Q��

Public Const ORDER_TPL_DETAIL_START_ROW As Long = 11      ' ���󖾍ׂ̖��׊J�n�s(�Z�N�V�������o���s)
Public Const ORDER_TPL_DETAIL_DEFAULT_ROWS As Long = 22   ' �e���v���[�g�����̖��׍s��(11:32�s)
Public Const ORDER_TPL_PRINT_TITLE_ROWS As String = "$7:$10"
Public Const ORDER_TPL_BLOCK_VENDOR_CODE_ROW As Long = 16 ' ��{���: �Ǝ҃R�[�h�s
Public Const ORDER_TPL_BLOCK_ORDER_NO_ROW As Long = 27    ' ��{���: �����ԍ��s
Public Const ORDER_TPL_VENDOR_ADO_NAME_FIELD As Long = 0      ' �Ǝ҃}�X�^ A�� �ƎҖ�
Public Const ORDER_TPL_VENDOR_ADO_OFFICIAL_FIELD As Long = 1  ' �Ǝ҃}�X�^ B�� �����Ҏ���
Public Const ORDER_TPL_VENDOR_ADO_WORK_FIELD As Long = 14     ' �Ǝ҃}�X�^ O�� �S���H��
Public Const ORDER_TPL_VENDOR_ADO_ALIAS_FIELD As Long = 15    ' �Ǝ҃}�X�^ P�� ����

Private mVendorInfoCache As Object       ' �x�X|��Ж� �� Array(�ƎҖ�, ����, �S���H��)
Private mBranchOfficeCodeCache As Object ' �x�X|�o���� �� ���X�R�[�h
Private mPlaceholderRepairDone As Boolean ' �Z�b�V�������Ŗh��I�Ȉꊇ��C�����{�ς݂�(Activate���̑S���������)

' �e���v���[�g�t�@�C����
Public Function OrderTplTemplateFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30C6, &H30F3, &H30D7, &H30EC, &H30FC) & _
                 CommonTextFromChars(&H30C8) & _
                 ".xlsx"
    End If
    OrderTplTemplateFileNameText = cached
End Function

Public Function OrderTplBaseNameBreakdownText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5185, &H8A33, &H660E, &H7D30)
    End If
    OrderTplBaseNameBreakdownText = cached
End Function

Public Function OrderTplBaseNameContractorText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H53D7, &H6CE8, &H8005, &H7528)
    End If
    OrderTplBaseNameContractorText = cached
End Function

Public Function OrderTplBaseNameAcceptanceText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H8ACB, &H66F8)
    End If
    OrderTplBaseNameAcceptanceText = cached
End Function

Public Function OrderTplBaseNameBranchCopyText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H652F, &H5E97, &H63A7)
    End If
    OrderTplBaseNameBranchCopyText = cached
End Function

Public Function OrderTplBaseNameAttachment3Text() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5225, &H7D19, &H2162)
    End If
    OrderTplBaseNameAttachment3Text = cached
End Function

' 条件書(取込後の基準名。テンプレートは工事区分別の「条件書(◯◯工事)」)
Public Function OrderTplBaseNameConditionText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6761, &H4EF6, &H66F8)
    End If
    OrderTplBaseNameConditionText = cached
End Function

Public Function OrderTplRailWeldingLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H30EC, &H30FC, &H30EB, &H6EB6, &H63A5)
    End If
    OrderTplRailWeldingLabelText = cached
End Function

' �{�H��ЕʒP����w�b�_�[�̐ڔ���
Public Function OrderTplUnitPriceHeaderSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5358, &H4FA1)
    End If
    OrderTplUnitPriceHeaderSuffixText = cached
End Function

Public Function OrderTplDayFirstText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H663C)
    End If
    OrderTplDayFirstText = cached
End Function

Public Function OrderTplSubtotalLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5C0F, &H8A08)
    End If
    OrderTplSubtotalLabelText = cached
End Function

' ���؂�Ȃ������̒P�ʈꗗ
Private Function OrderTplIntegerUnitListText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H53E3) & _
                 "|" & _
                 CommonTextFromChars(&H7A74) & _
                 "|" & _
                 CommonTextFromChars(&H56DE) & _
                 "|" & _
                 CommonTextFromChars(&H500B) & _
                 "|" & _
                 CommonTextFromChars(&H7B87, &H6240) & _
                 "|" & _
                 CommonTextFromChars(&H672C) & _
                 "|" & _
                 CommonTextFromChars(&H7D44) & _
                 "|" & _
                 CommonTextFromChars(&H5F0F) & _
                 "|" & _
                 CommonTextFromChars(&H679A)
    End If
    OrderTplIntegerUnitListText = cached
End Function

' ����2���̒P�ʈꗗ
Private Function OrderTplDecimalUnitListText() As String
    Static cached As String
    If cached = "" Then
        cached = "m|m3|M|m2|m" & _
                 CommonTextFromChars(&HB2) & _
                 "|" & _
                 CommonTextFromChars(&H33A1) & _
                 "|m" & _
                 CommonTextFromChars(&HB3) & _
                 "|M3|t"
    End If
    OrderTplDecimalUnitListText = cached
End Function

' �a��\���`��
Public Function OrderTplEraDateNumberFormatText() As String
    Static cached As String
    If cached = "" Then
        cached = "[$-411]ggge""" & _
                 CommonTextFromChars(&H5E74) & _
                 """m""" & _
                 CommonTextFromChars(&H6708) & _
                 """d""" & _
                 CommonTextFromChars(&H65E5) & _
                 """"
    End If
    OrderTplEraDateNumberFormatText = cached
End Function

Public Function OrderTplVendorNotFoundMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H696D, &H8005, &H30DE, &H30B9, &H30BF) & _
                 "(" & _
                 CommonTextFromChars(&H5168, &H793E, &H7248) & _
                 ")" & _
                 CommonTextFromChars(&H306B, &H8A72, &H5F53, &H3059, &H308B, &H4F1A, &H793E, &H304C) & _
                 CommonTextFromChars(&H898B, &H3064, &H304B, &H3089, &H306A, &H3044, &H305F, &H3081) & _
                 CommonTextFromChars(&H3001, &H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3092) & _
                 CommonTextFromChars(&H4F5C, &H6210, &H3067, &H304D, &H307E, &H305B, &H3093, &H3002)
    End If
    OrderTplVendorNotFoundMessageText = cached
End Function

Public Function OrderTplAliasEmptyMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H696D, &H8005, &H30DE, &H30B9, &H30BF) & _
                 "(" & _
                 CommonTextFromChars(&H5168, &H793E, &H7248) & _
                 ")" & _
                 CommonTextFromChars(&H306E) & _
                 "P" & _
                 CommonTextFromChars(&H5217) & _
                 "(" & _
                 CommonTextFromChars(&H7565, &H79F0) & _
                 ")" & _
                 CommonTextFromChars(&H304C, &H672A, &H5165, &H529B, &H306E, &H305F, &H3081, &H3001) & _
                 CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3092, &H4F5C) & _
                 CommonTextFromChars(&H6210, &H3067, &H304D, &H307E, &H305B, &H3093, &H3002)
    End If
    OrderTplAliasEmptyMessageText = cached
End Function

Public Function OrderTplTemplateNotFoundMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30C6, &H30F3, &H30D7, &H30EC, &H30FC) & _
                 CommonTextFromChars(&H30C8) & _
                 ".xlsx " & _
                 CommonTextFromChars(&H304C, &H898B, &H3064, &H304B, &H308A, &H307E, &H305B, &H3093) & _
                 CommonTextFromChars(&H3002)
    End If
    OrderTplTemplateNotFoundMessageText = cached
End Function

Public Function OrderTplDuplicateAliasMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H540C, &H3058, &H65BD, &H5DE5, &H4F1A, &H793E, &H304C, &H8907) & _
                 CommonTextFromChars(&H6570, &H306E, &H696D, &H8005, &H60C5, &H5831, &H30D6, &H30ED) & _
                 CommonTextFromChars(&H30C3, &H30AF, &H306B, &H5165, &H529B, &H3055, &H308C, &H3066) & _
                 CommonTextFromChars(&H3044, &H308B, &H305F, &H3081, &H3001, &H6CE8, &H6587, &H66F8) & _
                 CommonTextFromChars(&H30B7, &H30FC, &H30C8, &H306E, &H4F5C, &H6210, &H3092, &H4E2D) & _
                 CommonTextFromChars(&H6B62, &H3057, &H307E, &H3059, &H3002)
    End If
    OrderTplDuplicateAliasMessageText = cached
End Function

Public Function OrderTplRefreshDoneMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3078, &H306E) & _
                 CommonTextFromChars(&H518D, &H8EE2, &H8A18, &H304C, &H5B8C, &H4E86, &H3057, &H307E) & _
                 CommonTextFromChars(&H3057, &H305F, &H3002)
    End If
    OrderTplRefreshDoneMessageText = cached
End Function

Public Sub OrderTplClearCaches()
    Set mVendorInfoCache = Nothing
    Set mBranchOfficeCodeCache = Nothing
End Sub

Public Function OrderTplTemplateSheetBaseNames() As Variant
    OrderTplTemplateSheetBaseNames = Array(OrderTplBaseNameBreakdownText(), _
                                           OrderTplBaseNameContractorText(), _
                                           OrderTplBaseNameAcceptanceText(), _
                                           OrderTplBaseNameBranchCopyText(), _
                                           OrderTplBaseNameAttachment3Text())
End Function

' 生成シートの認識用の基準名(取込5シート + 条件書)。
' 条件書は工事区分別テンプレートから取り込むが、生成後は "条件書(略称)" となるため、
' 認識・後始末・再転記の判定はこの一覧で行う(取込のコピー対象は上の5固定のまま)。
Public Function OrderTplGeneratedSheetBaseNames() As Variant
    OrderTplGeneratedSheetBaseNames = Array(OrderTplBaseNameBreakdownText(), _
                                            OrderTplBaseNameContractorText(), _
                                            OrderTplBaseNameAcceptanceText(), _
                                            OrderTplBaseNameBranchCopyText(), _
                                            OrderTplBaseNameConditionText(), _
                                            OrderTplBaseNameAttachment3Text())
End Function

' 工事区分(基本情報10行)に一致する条件書テンプレートシートを template から探す。
' "条件書" で始まり工事区分名を含むシートを返す(パーレン表記の揺れに強い)。無ければ Nothing。
Public Function OrderTplFindConditionTemplateSheet(ByVal templateBook As Workbook, _
                                                   ByVal workType As String) As Worksheet
    If templateBook Is Nothing Then Exit Function
    Dim wt As String
    wt = CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(workType)))
    If wt = "" Then Exit Function

    Dim condBase As String
    condBase = CommonRemoveAllSpaces(OrderTplBaseNameConditionText())

    Dim ws As Worksheet
    For Each ws In templateBook.Worksheets
        Dim nm As String
        nm = CommonRemoveAllSpaces(CommonNormalizeText(ws.Name))
        If Len(nm) > Len(condBase) Then
            If Left$(nm, Len(condBase)) = condBase Then
                If InStr(1, nm, wt, vbTextCompare) > 0 Then
                    Set OrderTplFindConditionTemplateSheet = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

' �}�X�^�f�[�^�t�H���_���t�@�C���̃p�X����(���L�t�H���_ �� �u�b�N�e�t�H���_ �� �u�b�N����)
Public Function OrderTplMasterDataFilePath(ByVal fileName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")

    Dim candidatePath As String
    If Len(Trim$(userProfilePath)) > 0 Then
        candidatePath = userProfilePath & Chr$(92) & CommonCompanyNameText() & Chr$(92) & _
                        CommonOrderInvoiceDocumentFolderText() & Chr$(92) & _
                        CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then
            OrderTplMasterDataFilePath = candidatePath
            Exit Function
        End If
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        candidatePath = fso.GetParentFolderName(ThisWorkbook.Path) & Chr$(92) & _
                        CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then
            OrderTplMasterDataFilePath = candidatePath
            Exit Function
        End If

        candidatePath = ThisWorkbook.Path & Chr$(92) & CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then OrderTplMasterDataFilePath = candidatePath
    End If
End Function

' ��{���V�[�g�̎{�H��ЃZ��(11�s��)�̒l���擾����
Public Function OrderTplGetVendorCompanyName(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As String
    If wsInfo Is Nothing Then Exit Function
    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    OrderTplGetVendorCompanyName = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))
End Function

' �Ǝ҃}�X�^(�S�Д�)�����Ж�(B�� �����Ҏ���)�ŏƍ����AA�� �ƎҖ��EP�� ���́EO�� �S���H����Ԃ�
Public Function OrderTplResolveVendorMasterInfo(ByVal branchName As String, _
                                                ByVal officialName As String, _
                                                ByRef vendorName As String, _
                                                ByRef aliasText As String, _
                                                ByRef workText As String) As Boolean
    vendorName = ""
    aliasText = ""
    workText = ""

    Dim normalizedBranch As String
    Dim normalizedOfficial As String
    normalizedBranch = CommonNormalizeText(branchName)
    normalizedOfficial = CommonNormalizeText(officialName)
    If normalizedBranch = "" Or normalizedOfficial = "" Then Exit Function

    If mVendorInfoCache Is Nothing Then
        Set mVendorInfoCache = CreateObject("Scripting.Dictionary")
        mVendorInfoCache.CompareMode = vbTextCompare
    End If

    Dim cacheKey As String
    cacheKey = normalizedBranch & "|" & normalizedOfficial
    If mVendorInfoCache.Exists(cacheKey) Then
        Dim cachedInfo As Variant
        cachedInfo = mVendorInfoCache(cacheKey)
        vendorName = CStr(cachedInfo(0))
        aliasText = CStr(cachedInfo(1))
        workText = CStr(cachedInfo(2))
        OrderTplResolveVendorMasterInfo = (Len(vendorName) > 0)
        Exit Function
    End If

    Dim sourceFilePath As String
    sourceFilePath = OrderTplMasterDataFilePath(VENDOR_MASTER_FILE)
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim sheetName As String
    sheetName = ResolveAdoWorksheetName(connection, normalizedBranch)
    If sheetName = "" Then GoTo Cleanup

    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & Replace$(sheetName, "]", "]]") & "$A2:P500]", connection, 0, 1, 1

    Do Until recordset.EOF
        Dim rowOfficial As String
        rowOfficial = CommonNormalizeText(CommonNzText( _
            CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_OFFICIAL_FIELD)))
        If StrComp(rowOfficial, normalizedOfficial, vbTextCompare) = 0 Then
            vendorName = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_NAME_FIELD)))
            aliasText = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_ALIAS_FIELD)))
            workText = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_WORK_FIELD)))
            Exit Do
        End If
        recordset.MoveNext
    Loop
    CommonCloseAdoRecordset recordset

    mVendorInfoCache.Add cacheKey, Array(vendorName, aliasText, workText)
    OrderTplResolveVendorMasterInfo = (Len(vendorName) > 0)

Cleanup:
    CommonCloseAdoConnection connection
End Function

' �o������_�P���K�p����.xlsx �̒P���K�p����V�[�g���畔�X�R�[�h(G��)���擾����
Public Function OrderTplResolveBranchOfficeCode(ByVal branchName As String, _
                                                ByVal officeName As String) As String
    Dim normalizedBranch As String
    Dim normalizedOffice As String
    normalizedBranch = CommonNormalizeText(branchName)
    normalizedOffice = CommonNormalizeText(officeName)
    If normalizedBranch = "" Or normalizedOffice = "" Then Exit Function

    If mBranchOfficeCodeCache Is Nothing Then
        Set mBranchOfficeCodeCache = CreateObject("Scripting.Dictionary")
        mBranchOfficeCodeCache.CompareMode = vbTextCompare
    End If

    Dim cacheKey As String
    cacheKey = normalizedBranch & "|" & normalizedOffice
    If mBranchOfficeCodeCache.Exists(cacheKey) Then
        OrderTplResolveBranchOfficeCode = CStr(mBranchOfficeCodeCache(cacheKey))
        Exit Function
    End If

    Dim sourceFilePath As String
    sourceFilePath = OrderTplMasterDataFilePath(UNIT_PRICE_LINE_MASTER_FILE)
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim resultCode As String
    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & PRICE_LINE_SHEET & "$A2:G500]", connection, 0, 1, 1

    Do Until recordset.EOF
        If StrComp(CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 1))), _
                   normalizedBranch, vbTextCompare) = 0 Then
            If StrComp(CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 2))), _
                       normalizedOffice, vbTextCompare) = 0 Then
                resultCode = CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 6)))
                Exit Do
            End If
        End If
        recordset.MoveNext
    Loop
    CommonCloseAdoRecordset recordset

    mBranchOfficeCodeCache.Add cacheKey, resultCode
    OrderTplResolveBranchOfficeCode = resultCode

Cleanup:
    CommonCloseAdoConnection connection
End Function

Private Function ResolveAdoWorksheetName(ByVal connection As Object, ByVal targetSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)
    If sheetNames Is Nothing Then Exit Function

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(CommonNormalizeText(CStr(sheetName)), targetSheetName, vbTextCompare) = 0 Then
            ResolveAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

' �e���v���[�g�V�[�g�� + ����(P��)����V�[�g����g�ݗ��Ă�
Public Function OrderTplBuildSheetName(ByVal baseName As String, ByVal aliasText As String) As String
    OrderTplBuildSheetName = mod_Construction_OutputLayout.SanitizeSheetName(baseName & aliasText)
End Function

' �����ς݃e���v���[�g�V�[�g�����肵�A��{���Ɨ��̂�Ԃ�
Public Function OrderTplIsGeneratedSheet(ByVal ws As Worksheet, _
                                         ByRef baseName As String, _
                                         ByRef aliasText As String) As Boolean
    baseName = ""
    aliasText = ""
    If ws Is Nothing Then Exit Function

    Dim normalizedName As String
    normalizedName = mod_Construction_Import_Load.NormalizeSheetNameParentheses(CommonNormalizeText(ws.Name))

    Dim baseNames As Variant
    baseNames = OrderTplGeneratedSheetBaseNames()

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        Dim candidateBase As String
        candidateBase = CStr(baseNames(i))
        If Len(normalizedName) > Len(candidateBase) + 1 Then
            If StrComp(Left$(normalizedName, Len(candidateBase)), candidateBase, vbTextCompare) = 0 Then
                Dim suffixText As String
                suffixText = Mid$(normalizedName, Len(candidateBase) + 1)
                If Left$(suffixText, 1) = "(" And Right$(suffixText, 1) = ")" Then
                    baseName = candidateBase
                    aliasText = suffixText
                    OrderTplIsGeneratedSheet = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Function OrderTplSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    OrderTplSheetExists = Not ws Is Nothing
End Function

' �{�H�w����(�H��)/�{�H�ʒm��(�H��)�̎捞�ς݃V�[�g��T��(���݂�����������g�p)
Public Function OrderTplFindWorksSourceSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_Import_Load.IsManagedImportOutputSheet(ws) Then
            If Not mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
                If mod_Construction_Import_Load.SheetNameEndsWithSuffixText(ws.Name, CONSTRUCTION_SHEET_SUFFIX_WORKS) Then
                    If Not mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                        Set OrderTplFindWorksSourceSheet = ws
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws
End Function

' �{�H�w����(�n��)/�{�H�ʒm��(�n��)�̎捞�ς݃V�[�g��T��(���݂�����������g�p)
Public Function OrderTplFindWeldingSourceSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_Import_Load.IsManagedImportOutputSheet(ws) Then
            If Not mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
                If mod_Construction_Import_Load.SheetNameEndsWithSuffixText(ws.Name, CONSTRUCTION_SHEET_SUFFIX_WELDING) Then
                    If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                        Set OrderTplFindWeldingSourceSheet = ws
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws
End Function

' �_����於����ڔ���((�O��)/(�n�ڎw�����p)/(�n�ڒʒm���p))����������
' �����̃}�[�J�[�����֐��ֈϏ����A�w�����E�ʒm���̗��\�L�ɑΉ�����
Public Function OrderTplStripLineSuffix(ByVal lineText As String, ByVal isWeldingSource As Boolean) As String
    Dim t As String
    t = mod_Construction_Import_Load.NormalizeSheetNameParentheses(CommonNormalizeText(lineText))

    If isWeldingSource Then
        t = mod_Construction_LineMapping.RemoveWeldingInstructionMarker(t)
    Else
        t = mod_Construction_LineMapping.RemoveTrackDesignationMarker(t)
    End If
    OrderTplStripLineSuffix = Trim$(t)
End Function

' ���ʂ̒P�ʂ�����(���؂�Ȃ�)�Ώۂ����肷��
Public Function OrderTplIsIntegerUnit(ByVal unitText As String) As Boolean
    OrderTplIsIntegerUnit = UnitListContains(OrderTplIntegerUnitListText(), unitText)
End Function

' ���ʂ̒P�ʂ�����2���Ώۂ����肷��
Public Function OrderTplIsDecimalUnit(ByVal unitText As String) As Boolean
    OrderTplIsDecimalUnit = UnitListContains(OrderTplDecimalUnitListText(), unitText)
End Function

Private Function UnitListContains(ByVal listText As String, ByVal unitText As String) As Boolean
    Dim normalizedUnit As String
    normalizedUnit = CommonRemoveAllSpaces(CommonNormalizeText(unitText))
    If normalizedUnit = "" Then Exit Function
    ' �P���\�E�{�H�w�����̒P�ʂ͑S�p�p��(��/�l/��)�œ��邱�Ƃ����邽�ߔ��p�֐��K�����ďƍ�����
    normalizedUnit = StrConv(normalizedUnit, vbNarrow)

    Dim items As Variant
    items = Split(listText, "|")

    Dim i As Long
    For i = LBound(items) To UBound(items)
        If StrComp(normalizedUnit, CStr(items(i)), vbBinaryCompare) = 0 Then
            UnitListContains = True
            Exit Function
        End If
    Next i
End Function

' �����ς݃e���v���[�g(�x�X�T/�󒍎җp/��������)�Ɏc��v���[�X�z���_�[��������������B
' �e���v���[�g xlsx �̎��ȎQ��(��: ='�x�X�T(����)'!E20)�� #REF!�A�x�X�T�ւ̃~���[������
' �u�b�N�Čv�Z���ɏz�Q�ƃ_�C�A���O���o�����߁AVBA �]�L�O�ɒl�Z���֖߂��B
' �����ς݃e���v���[�g(�x�X�T/�󒍎җp/��������)�̃v���[�X�z���_�[�������ꊇ��C����B
' ������(GenerateVendorOrderSheets)�E�ē]�L���EF9���C�A�E�g�ύX���͊e�V�[�g���ʂɕ�C�ς݂̂��߁A
' �����Ăяo���� Force:=True �ŕK�����s���A��{���V�[�g�� Activate ����̖h��I�Ăяo��(Force�ȗ�)��
' �Z�b�V��������̂ݎ��s���āA�V�[�g�ؑւ̓s�x�̑S����(�̊��t���[�Y�̈��)���������B
Public Sub OrderTplRepairAllGeneratedPlaceholderFormulas(Optional ByVal Force As Boolean = False)
    If Not Force And mPlaceholderRepairDone Then Exit Sub

    Dim ws As Worksheet
    Dim baseName As String
    Dim aliasText As String

    For Each ws In ThisWorkbook.Worksheets
        If OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then
            OrderTplSanitizePlaceholderFormulas ws
        End If
    Next ws

    mPlaceholderRepairDone = True
End Sub

Public Sub OrderTplSanitizePlaceholderFormulas(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    Dim aliasText As String
    If Not OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then Exit Sub

    If baseName <> OrderTplBaseNameBranchCopyText() And _
       baseName <> OrderTplBaseNameContractorText() And _
       baseName <> OrderTplBaseNameAcceptanceText() Then Exit Sub

    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaRange Is Nothing Then Exit Sub

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In formulaRange.Cells
        Dim repCell As Range
        Set repCell = cell.MergeArea.Cells(1, 1)

        Dim repKey As String
        repKey = repCell.Address(True, True)
        If processed.Exists(repKey) Then GoTo ContinueCell
        processed.Add repKey, True

        ' ��\�Z��(�����̍���)���琔����1�񂾂��擾������֓n��(MergeArea�����̏d�������)
        Dim formulaText As String
        formulaText = ""
        On Error Resume Next
        formulaText = CStr(repCell.Formula)
        On Error GoTo 0

        If OrderTplShouldClearPlaceholderFormula(ws, baseName, repCell, formulaText) Then
            OrderTplClearFormulaCell repCell
        End If
ContinueCell:
    Next cell
End Sub

Private Function OrderTplShouldClearPlaceholderFormula(ByVal ws As Worksheet, _
                                                       ByVal baseName As String, _
                                                       ByVal cell As Range, _
                                                       ByVal formulaText As String) As Boolean
    If Len(formulaText) = 0 Then Exit Function
    If Left$(formulaText, 1) <> "=" Then Exit Function

    If InStr(1, formulaText, "#REF!", vbTextCompare) > 0 Then
        OrderTplShouldClearPlaceholderFormula = True
        Exit Function
    End If

    If OrderTplFormulaReferencesSameCell(ws, cell, formulaText) Then
        OrderTplShouldClearPlaceholderFormula = True
        Exit Function
    End If

    If baseName = OrderTplBaseNameContractorText() Or _
       baseName = OrderTplBaseNameAcceptanceText() Then
        If OrderTplFormulaReferencesBranchCopySheet(formulaText) Then
            OrderTplShouldClearPlaceholderFormula = True
        End If
    End If
End Function

Private Function OrderTplFormulaReferencesSameCell(ByVal ws As Worksheet, _
                                                 ByVal cell As Range, _
                                                 ByVal formulaText As String) As Boolean
    Dim normalizedFormula As String
    normalizedFormula = UCase$(Replace$(formulaText, " ", ""))

    Dim quotedSheet As String
    quotedSheet = "'" & Replace$(ws.Name, "'", "''") & "'!"

    Dim addr As Variant
    For Each addr In Array(cell.Address(True, False), cell.Address(False, False), _
                           cell.Address(True, True), cell.Address(False, True))
        If InStr(1, normalizedFormula, UCase$(quotedSheet & CStr(addr)), vbBinaryCompare) > 0 Then
            OrderTplFormulaReferencesSameCell = True
            Exit Function
        End If
    Next addr

    If InStr(1, normalizedFormula, "!", vbBinaryCompare) = 0 Then
        For Each addr In Array(cell.Address(True, False), cell.Address(False, False), _
                               cell.Address(True, True), cell.Address(False, True))
            If normalizedFormula = "=" & UCase$(CStr(addr)) Then
                OrderTplFormulaReferencesSameCell = True
                Exit Function
            End If
        Next addr
    End If
End Function

Private Function OrderTplFormulaReferencesBranchCopySheet(ByVal formulaText As String) As Boolean
    Dim branchBase As String
    branchBase = OrderTplBaseNameBranchCopyText()
    If InStr(1, formulaText, "'" & branchBase, vbTextCompare) > 0 Then
        OrderTplFormulaReferencesBranchCopySheet = True
    End If
End Function

Private Sub OrderTplClearFormulaCell(ByVal cell As Range)
    Dim writeCell As Range
    On Error Resume Next
    Set writeCell = cell.MergeArea.Cells(1, 1)
    If writeCell Is Nothing Then Set writeCell = cell
    On Error Resume Next
    writeCell.ClearContents
    On Error GoTo 0
End Sub

Public Sub OrderTplLog(ByVal msg As String)
    mod_DebugLog.Log "[OrderTpl] " & msg
End Sub
