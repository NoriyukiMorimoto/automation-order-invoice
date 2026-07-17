Option Explicit

' �{�H��Њm�莞�ɒ������e���v���[�g5�V�[�g(���󖾍�/�󒍎җp/��������/�x�X�T/�ʎ��V)��
' �w���[���w��(�ʒm)�V�[�g�̉E���։�Ђ��Ƃɑ}���E�폜����B
' ���C����: CHANGELOG.md �Q��

Private mGenerating As Boolean
Private mDetailRefreshScheduledTime As Date
Private mPendingVendorIndexes As Object        ' �x�������҂��̃u���b�N�ԍ�(Dictionary)
Private mVendorGenScheduledTime As Date

Private Function GenerateErrorText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H306E, &H4F5C) & _
                 CommonTextFromChars(&H6210, &H4E2D, &H306B, &H30A8, &H30E9, &H30FC, &H304C, &H767A) & _
                 CommonTextFromChars(&H751F, &H3057, &H307E, &H3057, &H305F, &H3002)
    End If
    GenerateErrorText = cached
End Function

Private Function RefreshErrorText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3078, &H306E) & _
                 CommonTextFromChars(&H518D, &H8EE2, &H8A18, &H4E2D, &H306B, &H30A8, &H30E9, &H30FC) & _
                 CommonTextFromChars(&H304C, &H767A, &H751F, &H3057, &H307E, &H3057, &H305F, &H3002)
    End If
    RefreshErrorText = cached
End Function

' Sheet1(��{���)��Worksheet_Change����Ă΂������B
' 11�s��(�{�H���)�̊m��Ńe���v���[�g�V�[�g���č쐬���A�N���A���͌ǎ��V�[�g���폜����B
Public Sub HandleVendorNameCellChange(ByVal wsInfo As Worksheet, ByVal changedCell As Range)
    If wsInfo Is Nothing Then Exit Sub
    If changedCell Is Nothing Then Exit Sub
    If mGenerating Then Exit Sub

    Dim vendorIndex As Long
    vendorIndex = mod_VendorMaster.GetVendorIndexFromValueColumnPublic(changedCell.Column)
    If vendorIndex < 1 Then Exit Sub

    ' ActiveX�R���{���̃R���g���[���C�x���g������ Worksheets.Copy �����s�����
    ' ���s���G���[1004(���̃V�[�g���R�s�[�ł��܂���ł���)�ɂȂ邽�߁A
    ' OnTime �ŃC�x���g������(�t�H�[�J�X���퉻��)�ɒx�����s����
    ScheduleVendorSheetGeneration vendorIndex
End Sub

' �w��u���b�N�̃V�[�g������x���\�񂷂�(�����u���b�N�̗\��͂܂Ƃ߂Ď��s�����)
Public Sub ScheduleVendorSheetGeneration(ByVal vendorIndex As Long)
    If mPendingVendorIndexes Is Nothing Then
        Set mPendingVendorIndexes = CreateObject("Scripting.Dictionary")
    End If
    If Not mPendingVendorIndexes.Exists(vendorIndex) Then
        mPendingVendorIndexes.Add vendorIndex, True
    End If

    CancelScheduledVendorSheetGeneration

    mVendorGenScheduledTime = Now
    On Error Resume Next
    Application.OnTime EarliestTime:=mVendorGenScheduledTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledVendorSheetGeneration"
    If Err.Number <> 0 Then mVendorGenScheduledTime = 0
    On Error GoTo 0
End Sub

Public Sub CancelScheduledVendorSheetGeneration()
    If mVendorGenScheduledTime = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mVendorGenScheduledTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledVendorSheetGeneration", _
                       Schedule:=False
    On Error GoTo 0
    mVendorGenScheduledTime = 0
End Sub

' Application.OnTime ����Ă΂��x�������̎��s��
Public Sub RunScheduledVendorSheetGeneration()
    mVendorGenScheduledTime = 0
    If mPendingVendorIndexes Is Nothing Then Exit Sub
    If mPendingVendorIndexes.Count = 0 Then Exit Sub

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    ' ���s�O�ɗ\�񃊃X�g��ޔ����ăN���A����(���s���̍ė\��ƏՓ˂����Ȃ�)
    Dim pendingKeys As Variant
    pendingKeys = mPendingVendorIndexes.Keys
    mPendingVendorIndexes.RemoveAll

    Dim prevEnableEvents As Boolean
    prevEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim i As Long
    For i = LBound(pendingKeys) To UBound(pendingKeys)
        GenerateVendorOrderSheets wsInfo, CLng(pendingKeys(i))
    Next i

    ' �����������󖾍ׂ́u�v�v����{���33�s�֔��f����
    mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals

    Application.EnableEvents = prevEnableEvents
End Sub

' �w��u���b�N�̎{�H��Ђ̃e���v���[�g5�V�[�g���폜���čč쐬���A���󖾍ׂ֓]�L����
Public Sub GenerateVendorOrderSheets(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If mGenerating Then Exit Sub

    Dim prevDisplayAlerts As Boolean
    Dim prevScreenUpdating As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    prevScreenUpdating = Application.ScreenUpdating

    Dim previousActiveSheet As Object
    Set previousActiveSheet = ThisWorkbook.ActiveSheet

    On Error GoTo ErrorHandler
    mGenerating = True
    Application.ScreenUpdating = False

    Dim companyName As String
    companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
    mod_OrderTpl_Shared.OrderTplLog "GenerateVendorOrderSheets index=" & vendorIndex & " company=[" & companyName & "]"

    RemoveOrphanGeneratedSheets wsInfo
    If companyName = "" Then GoTo Cleanup

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim vendorName As String
    Dim aliasText As String
    Dim workText As String
    If Not mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then
        mod_OrderTpl_Shared.OrderTplLog "vendor not found in master: " & companyName
        MsgBox mod_OrderTpl_Shared.OrderTplVendorNotFoundMessageText() & vbCrLf & companyName, vbExclamation
        GoTo Cleanup
    End If
    If aliasText = "" Then
        mod_OrderTpl_Shared.OrderTplLog "alias empty: " & companyName
        MsgBox mod_OrderTpl_Shared.OrderTplAliasEmptyMessageText() & vbCrLf & companyName, vbExclamation
        GoTo Cleanup
    End If

    If DuplicateAliasExists(wsInfo, vendorIndex, branchName, aliasText) Then
        mod_OrderTpl_Shared.OrderTplLog "duplicate alias: " & aliasText
        MsgBox mod_OrderTpl_Shared.OrderTplDuplicateAliasMessageText() & vbCrLf & companyName, vbExclamation
        GoTo Cleanup
    End If

    RemoveGeneratedSheetsByAlias aliasText

    Dim anchorSheet As Worksheet
    Set anchorSheet = ResolveInsertAnchor(wsInfo, vendorIndex, branchName)

    Dim templatePath As String
    templatePath = mod_OrderTpl_Shared.OrderTplMasterDataFilePath(mod_OrderTpl_Shared.OrderTplTemplateFileNameText())
    If templatePath = "" Then
        mod_OrderTpl_Shared.OrderTplLog "template not found"
        MsgBox mod_OrderTpl_Shared.OrderTplTemplateNotFoundMessageText(), vbExclamation
        GoTo Cleanup
    End If

    Dim openedHere As Boolean
    Dim templateWorkbook As Workbook
    Set templateWorkbook = mod_Construction_Import_Load.OpenWorkbookReadOnly(templatePath, openedHere)
    If templateWorkbook Is Nothing Then
        mod_OrderTpl_Shared.OrderTplLog "template open failed: " & templatePath
        MsgBox mod_OrderTpl_Shared.OrderTplTemplateNotFoundMessageText() & vbCrLf & templatePath, vbExclamation
        GoTo Cleanup
    End If

    Dim baseNames As Variant
    baseNames = mod_OrderTpl_Shared.OrderTplTemplateSheetBaseNames()

    Application.DisplayAlerts = False
    mod_OrderTpl_Shared.OrderTplLog "step: copy template after=" & anchorSheet.Name
    templateWorkbook.Worksheets(baseNames).Copy After:=anchorSheet

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        Dim wsGenerated As Worksheet
        Set wsGenerated = ThisWorkbook.Sheets(anchorSheet.Index + 1 + i - LBound(baseNames))
        wsGenerated.Name = mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(i)), aliasText)
        mod_OrderTpl_Shared.OrderTplLog "step: sanitize " & wsGenerated.Name
        mod_OrderTpl_Shared.OrderTplSanitizePlaceholderFormulas wsGenerated
    Next i

    ' �H���敪(��{���10�s)�ɉ�������������1���荞�݁A�x�X�T�̉E(�ʎ�III�̍�)�֔z�u����
    Dim workTypeText As String
    workTypeText = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, _
                     mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)).value))
    If workTypeText = "" Then workTypeText = workText

    ' �ʎ�III�͋O���H���̎{�H��Ђ݂̂ɐ�������B
    ' �H���敪(��{���10�s��)���O���H���łȂ���΁A�R�s�[�ς݂̕ʎ�III�폜����B
    If Not mod_OrderTpl_Shared.OrderTplIsRailWorkType(workTypeText) Then
        Dim attach3Name As String
        attach3Name = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
            mod_OrderTpl_Shared.OrderTplBaseNameAttachment3Text(), aliasText)
        If mod_OrderTpl_Shared.OrderTplSheetExists(attach3Name) Then
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(attach3Name).Delete
            mod_OrderTpl_Shared.OrderTplLog _
                "attachment3 skipped (non-rail workType=[" & workTypeText & "]) removed " & attach3Name
        End If
    End If

    Dim condTemplateSheet As Worksheet
    Set condTemplateSheet = mod_OrderTpl_Shared.OrderTplFindConditionTemplateSheet(templateWorkbook, workTypeText)
    If Not condTemplateSheet Is Nothing Then
        Dim wsBranchCopy As Worksheet
        Set wsBranchCopy = ThisWorkbook.Worksheets( _
            mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBranchCopyText(), aliasText))
        condTemplateSheet.Copy After:=wsBranchCopy
        Dim wsCondition As Worksheet
        Set wsCondition = ThisWorkbook.Sheets(wsBranchCopy.Index + 1)
        wsCondition.Name = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
            mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
        mod_OrderTpl_Shared.OrderTplLog "step: condition sheet " & wsCondition.Name & " from " & condTemplateSheet.Name
        mod_OrderTpl_Header.ApplyConditionSheetHeader wsInfo, wsCondition, vendorIndex
        mod_OrderTpl_Header.SetupConditionCheckboxExclusivity wsCondition
    Else
        mod_OrderTpl_Shared.OrderTplLog "condition template not found for workType=[" & workTypeText & "]"
    End If

    If openedHere Then templateWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = prevDisplayAlerts

    Dim wsBreakdown As Worksheet
    Set wsBreakdown = ThisWorkbook.Worksheets( _
        mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(LBound(baseNames))), aliasText))

    mod_OrderTpl_Shared.OrderTplLog "step: apply headers alias=" & aliasText
    mod_OrderTpl_Header.ApplyVendorSheetHeaders wsInfo, vendorIndex, aliasText
    mod_OrderTpl_Shared.OrderTplLog "step: apply breakdown " & wsBreakdown.Name
    mod_OrderTpl_Detail.ApplyBreakdownDetails wsBreakdown, vendorName, companyName, branchName, workTypeText

    Dim condSheetNameFinal As String
    condSheetNameFinal = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
        mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetNameFinal) Then
        mod_OrderTpl_Header.SetupConditionCheckboxExclusivity _
            ThisWorkbook.Worksheets(condSheetNameFinal)
    End If

    mod_OrderTpl_Shared.OrderTplLog "GenerateVendorOrderSheets done alias=" & aliasText

Cleanup:
    On Error Resume Next
    If Not previousActiveSheet Is Nothing Then previousActiveSheet.Activate
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    mGenerating = False
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "GenerateVendorOrderSheets error: " & Err.Number & " " & Err.Description
    MsgBox GenerateErrorText() & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Sub

' �S�m���Ђ̓��󖾍ׂ��ē]�L����(�{�H�w�������̎捞��Ɏ蓮���s������J�}�N��)
Public Sub RefreshAllVendorOrderDetails()
    RefreshAllVendorOrderDetailsCore True
End Sub

' �S�m���Ђ̓��󖾍ׂ��ē]�L����(�������b�Z�[�W�Ȃ��B�������f�p)
Public Sub RefreshAllVendorOrderDetailsSilent()
    RefreshAllVendorOrderDetailsCore False
End Sub

Private Sub RefreshAllVendorOrderDetailsCore(ByVal showCompletionMessage As Boolean)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevEnableEvents As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevEnableEvents = Application.EnableEvents

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    mod_OrderTpl_Shared.OrderTplClearCaches

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
                Dim breakdownSheetName As String
                breakdownSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
                If mod_OrderTpl_Shared.OrderTplSheetExists(breakdownSheetName) Then
                    Dim wsBreakdown As Worksheet
                    Set wsBreakdown = ThisWorkbook.Worksheets(breakdownSheetName)

                    Dim workTypeText As String
                    workTypeText = CommonNormalizeText(CommonNzText( _
                        wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, _
                                     mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)).value))
                    If workTypeText = "" Then workTypeText = workText

                    mod_OrderTpl_Header.ApplyVendorSheetHeaders wsInfo, vendorIndex, aliasText
                    mod_OrderTpl_Detail.ApplyBreakdownDetails wsBreakdown, vendorName, companyName, branchName, workTypeText

                    Dim condSheetName As String
                    condSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                        mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
                    If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetName) Then
                        mod_OrderTpl_Header.ApplyConditionSheetHeader wsInfo, _
                            ThisWorkbook.Worksheets(condSheetName), vendorIndex
                        mod_OrderTpl_Header.SetupConditionCheckboxExclusivity _
                            ThisWorkbook.Worksheets(condSheetName)
                    End If
                Else
                    GenerateVendorOrderSheets wsInfo, vendorIndex
                End If
            Else
                mod_OrderTpl_Shared.OrderTplLog "Refresh: vendor not resolved: " & companyName
            End If
        End If
    Next vendorIndex

    Application.Calculation = prevCalculation
    mod_OrderTpl_Shared.OrderTplRepairAllGeneratedPlaceholderFormulas True
    Application.Calculate
    ' ���󖾍ׂ́u�v�v����{���33�s�֔��f����(���󖾍ׂ̍č\�z��)
    mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating

    If showCompletionMessage Then
        MsgBox mod_OrderTpl_Shared.OrderTplRefreshDoneMessageText(), vbInformation
    End If
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorOrderDetails error: " & Err.Number & " " & Err.Description
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    If showCompletionMessage Then
        MsgBox RefreshErrorText() & vbCrLf & Err.Description, vbExclamation
    End If
End Sub

' �{�s�w�����E�{�s�ʒm���V�[�g�̎{�H��З�(�H��:A��/�n��:A�EB��)�ύX���ɁA
' �x�����s�őS�Ѝē]�L��\�񂷂�(ThisWorkbook.Workbook_SheetChange ����Ă΂��)
Public Sub HandleSourceSheetVendorCellChange(ByVal sh As Object, ByVal target As Range)
    If mGenerating Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim ws As Worksheet
    Set ws = sh

    If Not mod_Construction_Import_Load.IsManagedImportOutputSheet(ws) Then Exit Sub
    If mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then Exit Sub

    Dim vendorColumns As Range
    If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
        Set vendorColumns = ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 2))
    Else
        Set vendorColumns = ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 1))
    End If
    If Intersect(target, vendorColumns) Is Nothing Then Exit Sub

    ScheduleOrderDetailRefresh
    Exit Sub

Quiet:
    Err.Clear
End Sub

' �S�Ѝē]�L�̒x�����s��\�񂷂�(�A���ύX��1��ɂ܂Ƃ߂�)
Public Sub ScheduleOrderDetailRefresh()
    CancelScheduledOrderDetailRefresh

    mDetailRefreshScheduledTime = Now + TimeSerial(0, 0, 1)
    On Error Resume Next
    Application.OnTime EarliestTime:=mDetailRefreshScheduledTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledOrderDetailRefresh"
    If Err.Number <> 0 Then mDetailRefreshScheduledTime = 0
    On Error GoTo 0
End Sub

Public Sub CancelScheduledOrderDetailRefresh()
    If mDetailRefreshScheduledTime = 0 Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=mDetailRefreshScheduledTime, _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledOrderDetailRefresh", _
                       Schedule:=False
    On Error GoTo 0
    mDetailRefreshScheduledTime = 0
End Sub

' Application.OnTime ����Ă΂��x���ē]�L�̎��s��
Public Sub RunScheduledOrderDetailRefresh()
    mDetailRefreshScheduledTime = 0
    RefreshAllVendorOrderDetailsCore False
End Sub

' �����ς݂̃e���v���[�g�V�[�g�����ׂč폜����(��{���N���A���̌�n���p)
Public Sub RemoveAllGeneratedOrderTemplateSheets()
    Dim sheetNamesToDelete As Collection
    Set sheetNamesToDelete = New Collection

    Dim ws As Worksheet
    Dim baseName As String
    Dim sheetAlias As String
    For Each ws In ThisWorkbook.Worksheets
        If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, sheetAlias) Then
            sheetNamesToDelete.Add ws.Name
        End If
    Next ws

    DeleteSheetsByNameList sheetNamesToDelete
End Sub

' �w�藪�̂̃e���v���[�g5�V�[�g���폜����
Public Sub RemoveGeneratedSheetsByAlias(ByVal aliasText As String)
    Dim normalizedAlias As String
    normalizedAlias = mod_Construction_Import_Load.NormalizeSheetNameParentheses(CommonNormalizeText(aliasText))
    If normalizedAlias = "" Then Exit Sub

    Dim sheetNamesToDelete As Collection
    Set sheetNamesToDelete = New Collection

    Dim ws As Worksheet
    Dim baseName As String
    Dim sheetAlias As String
    For Each ws In ThisWorkbook.Worksheets
        If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, sheetAlias) Then
            If StrComp(sheetAlias, normalizedAlias, vbTextCompare) = 0 Then
                sheetNamesToDelete.Add ws.Name
            End If
        End If
    Next ws

    DeleteSheetsByNameList sheetNamesToDelete
End Sub

' ��{���̊m���ЂɑΉ����Ȃ������ς݃V�[�g���폜����
Public Sub RemoveOrphanGeneratedSheets(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim expectedAliases As Object
    Set expectedAliases = CreateObject("Scripting.Dictionary")
    expectedAliases.CompareMode = vbTextCompare

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
                If aliasText <> "" Then
                    Dim aliasKey As String
                    aliasKey = mod_Construction_Import_Load.NormalizeSheetNameParentheses(aliasText)
                    If Not expectedAliases.Exists(aliasKey) Then expectedAliases.Add aliasKey, True
                End If
            Else
                ' �}�X�^�������̊m���Ђ�����ꍇ�͌�폜������邽�ߌ�n���𒆎~����
                mod_OrderTpl_Shared.OrderTplLog "RemoveOrphan skipped (unresolved): " & companyName
                Exit Sub
            End If
        End If
    Next vendorIndex

    Dim sheetNamesToDelete As Collection
    Set sheetNamesToDelete = New Collection

    Dim ws As Worksheet
    Dim baseName As String
    Dim sheetAlias As String
    For Each ws In ThisWorkbook.Worksheets
        If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, sheetAlias) Then
            If Not expectedAliases.Exists(sheetAlias) Then
                sheetNamesToDelete.Add ws.Name
            End If
        End If
    Next ws

    DeleteSheetsByNameList sheetNamesToDelete
End Sub

Private Sub DeleteSheetsByNameList(ByVal sheetNames As Collection)
    If sheetNames Is Nothing Then Exit Sub
    If sheetNames.Count = 0 Then Exit Sub

    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        mod_OrderTpl_Shared.OrderTplLog "delete sheet: " & CStr(sheetName)
        mod_Construction_Import_Load.DeleteSheetByName CStr(sheetName)
    Next sheetName

    Application.DisplayAlerts = prevDisplayAlerts
End Sub

' �������̂����̊m��u���b�N�ɑ��݂��邩
Private Function DuplicateAliasExists(ByVal wsInfo As Worksheet, _
                                      ByVal vendorIndex As Long, _
                                      ByVal branchName As String, _
                                      ByVal aliasText As String) As Boolean
    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        If i <> vendorIndex Then
            Dim companyName As String
            companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, i)
            If companyName <> "" Then
                Dim vendorName As String
                Dim otherAlias As String
                Dim workText As String
                If mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, otherAlias, workText) Then
                    If StrComp(CommonNormalizeText(otherAlias), CommonNormalizeText(aliasText), vbTextCompare) = 0 Then
                        DuplicateAliasExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' �}���ʒu�̌���: ���O�u���b�N�̕ʎ��V(����) �� �w���[���w��/�ʒm �� �捞�ςݍŉE�V�[�g �� ��{���
Private Function ResolveInsertAnchor(ByVal wsInfo As Worksheet, _
                                     ByVal vendorIndex As Long, _
                                     ByVal branchName As String) As Worksheet
    Dim i As Long
    For i = vendorIndex - 1 To 1 Step -1
        Dim companyName As String
        companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, i)
        If companyName <> "" Then
            Dim vendorName As String
            Dim aliasText As String
            Dim workText As String
            If mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then
                Dim lastSheetName As String
                lastSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameAttachment3Text(), aliasText)
                If mod_OrderTpl_Shared.OrderTplSheetExists(lastSheetName) Then
                    Set ResolveInsertAnchor = ThisWorkbook.Worksheets(lastSheetName)
                    Exit Function
                End If
            End If
        End If
    Next i

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
            Set ResolveInsertAnchor = ws
            Exit Function
        End If
    Next ws

    Dim anchorSheet As Worksheet
    Set anchorSheet = mod_Construction_Import_Load.GetImportSheetAnchorSheet()

    Dim rightmostSheet As Worksheet
    Set rightmostSheet = mod_Construction_Import_Load.FindRightmostManagedImportSheetAfterAnchor(anchorSheet)
    If Not rightmostSheet Is Nothing Then
        Set ResolveInsertAnchor = rightmostSheet
    ElseIf Not anchorSheet Is Nothing Then
        Set ResolveInsertAnchor = anchorSheet
    Else
        Set ResolveInsertAnchor = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
End Function
