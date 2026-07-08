Option Explicit

' 施工会社確定時に注文書テンプレート5シート(内訳明細/受注者用/注文請書/支店控/別紙Ⅲ)を
' 購入充当指示(通知)シートの右側へ会社ごとに挿入・削除する。
' 改修履歴: CHANGELOG.md 参照

Private mGenerating As Boolean

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

' Sheet1(基本情報)のWorksheet_Changeから呼ばれる入口。
' 11行目(施工会社)の確定でテンプレートシートを再作成し、クリア時は孤児シートを削除する。
Public Sub HandleVendorNameCellChange(ByVal wsInfo As Worksheet, ByVal changedCell As Range)
    If wsInfo Is Nothing Then Exit Sub
    If changedCell Is Nothing Then Exit Sub
    If mGenerating Then Exit Sub

    Dim vendorIndex As Long
    vendorIndex = mod_VendorMaster.GetVendorIndexFromValueColumnPublic(changedCell.Column)
    If vendorIndex < 1 Then Exit Sub

    GenerateVendorOrderSheets wsInfo, vendorIndex
End Sub

' 指定ブロックの施工会社のテンプレート5シートを削除して再作成し、内訳明細へ転記する
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
    templateWorkbook.Worksheets(baseNames).Copy After:=anchorSheet

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        ThisWorkbook.Worksheets(anchorSheet.Index + 1 + i - LBound(baseNames)).Name = _
            mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(i)), aliasText)
    Next i

    If openedHere Then templateWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = prevDisplayAlerts

    Dim wsBreakdown As Worksheet
    Set wsBreakdown = ThisWorkbook.Worksheets( _
        mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(LBound(baseNames))), aliasText))

    Dim workTypeText As String
    workTypeText = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, _
                     mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)).value))
    If workTypeText = "" Then workTypeText = workText

    mod_OrderTpl_Header.ApplyBreakdownHeader wsInfo, wsBreakdown, vendorIndex
    mod_OrderTpl_Detail.ApplyBreakdownDetails wsBreakdown, vendorName, companyName, branchName, workTypeText

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

' 全確定会社の内訳明細を再転記する(施工指示書等の取込後に手動実行する公開マクロ)
Public Sub RefreshAllVendorOrderDetails()
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

                    mod_OrderTpl_Header.ApplyBreakdownHeader wsInfo, wsBreakdown, vendorIndex
                    mod_OrderTpl_Detail.ApplyBreakdownDetails wsBreakdown, vendorName, companyName, branchName, workTypeText
                Else
                    GenerateVendorOrderSheets wsInfo, vendorIndex
                End If
            Else
                mod_OrderTpl_Shared.OrderTplLog "Refresh: vendor not resolved: " & companyName
            End If
        End If
    Next vendorIndex

    Application.Calculation = prevCalculation
    Application.Calculate
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating

    MsgBox mod_OrderTpl_Shared.OrderTplRefreshDoneMessageText(), vbInformation
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorOrderDetails error: " & Err.Number & " " & Err.Description
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox RefreshErrorText() & vbCrLf & Err.Description, vbExclamation
End Sub

' 指定略称のテンプレート5シートを削除する
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

' 基本情報の確定会社に対応しない生成済みシートを削除する
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
                ' マスタ未解決の確定会社がある場合は誤削除を避けるため後始末を中止する
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

' 同じ略称が他の確定ブロックに存在するか
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

' 挿入位置の決定: 直前ブロックの別紙Ⅲ(略称) → 購入充当指示/通知 → 取込済み最右シート → 基本情報
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
