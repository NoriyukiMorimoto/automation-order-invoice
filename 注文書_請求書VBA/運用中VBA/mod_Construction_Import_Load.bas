Option Explicit

' ????: CHANGELOG.md ??
' mod_Construction_Import_Load (split from mod_Construction_Order_Import)

Public Sub ImportConstructionDocumentCore()
    Dim scrn As Boolean, evt As Boolean, alerts As Boolean

    scrn = Application.screenUpdating
    evt = Application.EnableEvents
    alerts = Application.DisplayAlerts

    mSuppressOverwritePrompt = False
    Set mLastCreatedImportSheet = Nothing
    mod_Construction_LineMapping.ClearProjectLineNameAliasCache

    On Error GoTo Cleanup

    ' 既存シートの確認・削除は UI 抑制前に行う(「はい」で削除後にファイル選択ダイアログを開く)
    Dim cancelImport As Boolean
    PrepareExistingOutputSheets cancelImport
    If cancelImport Then GoTo Cleanup

    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim srcPaths As Collection
    Set srcPaths = PickSourceFiles()
    If srcPaths Is Nothing Then GoTo Cleanup
    If srcPaths.Count = 0 Then GoTo Cleanup

    Dim tWorks As Long, tWeld As Long, tPurch As Long
    Dim lastSheet As Worksheet
    Dim fileIndex As Long

    mSuppressOverwritePrompt = True
    Dim srcPathVar As Variant
    For Each srcPathVar In srcPaths
        fileIndex = fileIndex + 1
        LogCI "===== 取込ファイル " & fileIndex & "/" & srcPaths.Count & " [" & CStr(srcPathVar) & "] ====="
        ImportOneConstructionDocument CStr(srcPathVar), tWorks, tWeld, tPurch, lastSheet
    Next srcPathVar
    mSuppressOverwritePrompt = False

    mod_Construction_BasicTotals.RefreshBasicInfoConstructionTotalsCore

    NormalizeManagedImportSheetOrder

    If Not lastSheet Is Nothing Then
        lastSheet.Activate
        lastSheet.Range("A1").Select
    End If

    Application.screenUpdating = scrn
    RestoreAutomaticCalculation
    Application.EnableEvents = evt
    Application.DisplayAlerts = alerts

    LogCI "取込み完了 ファイル数=" & srcPaths.Count & _
          " 工事=" & tWorks & " 溶接=" & tWeld & " 購入=" & tPurch
    Exit Sub

Cleanup:
    mSuppressOverwritePrompt = False
    Set mLastCreatedImportSheet = Nothing
    Application.screenUpdating = scrn
    RestoreAutomaticCalculation
    Application.EnableEvents = evt
    Application.DisplayAlerts = alerts

    If Err.Number <> 0 Then
        MsgBox "取込み処理でエラーが発生しました。" & vbCrLf & _
               "Err " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub

Public Sub RestoreAutomaticCalculation()
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
End Sub

Public Sub ImportOneConstructionDocument(ByVal srcPath As String, _
                                          ByRef tWorks As Long, _
                                          ByRef tWeld As Long, _
                                          ByRef tPurch As Long, _
                                          ByRef lastSheet As Worksheet)
    Dim srcWb As Workbook
    Dim srcOpenedHere As Boolean

    On Error GoTo Cleanup

    LogCI "srcPath=[" & srcPath & "]"

    Dim docType As Long
    docType = DetermineDocType(srcPath)
    If docType = 0 Then GoTo Cleanup
    LogCI "docType=" & docType

    If Not ValidateBasicInfoForImport(docType) Then GoTo Cleanup

    Set srcWb = OpenWorkbookReadOnly(srcPath, srcOpenedHere)
    If srcWb Is Nothing Then
        MsgBox "取込対象ブックを開けませんでした。" & vbCrLf & srcPath, vbExclamation
        GoTo Cleanup
    End If

    Dim srcWs As Worksheet
    Set srcWs = srcWb.ActiveSheet

    Dim refValueSourceCell As String
    Dim refValue As Variant
    If docType = DOC_NOTICE Then
        refValueSourceCell = REF_VALUE_SOURCE_CELL_NOTICE
    Else
        refValueSourceCell = REF_VALUE_SOURCE_CELL_ORDER
    End If
    refValue = srcWs.Range(refValueSourceCell).value
    LogCI "参照シート " & refValueSourceCell & "=[" & CommonNzText(refValue) & "]"

    Dim sourceA3Text As String
    sourceA3Text = CommonNzText(srcWs.Range(SOURCE_SHEET_NAME_CELL).value)

    Dim baseSheetName As String
    baseSheetName = mod_Construction_OutputLayout.SanitizeSheetName(sourceA3Text)
    If baseSheetName = "" Then
        MsgBox "取込ブックの " & SOURCE_SHEET_NAME_CELL & " が空のため、シート名を決定できません。", vbExclamation
        GoTo Cleanup
    End If
    LogCI "baseSheetName=[" & baseSheetName & "]"

    Dim guidanceDocumentName As String
    guidanceDocumentName = mod_Construction_LineMapping.ResolveGuidanceDocumentName(sourceA3Text, docType)
    LogCI "再取込み案内文書名=[" & guidanceDocumentName & "] A3=[" & sourceA3Text & "]"

    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.rows.Count, "A").End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "取込対象データ(" & DATA_START_ROW & "行目以降)が見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    Dim colType As String, colDayNight As String, colUnit As String
    Dim colQty As String, colLine As String, colMgr As String
    Dim colPrice As String, colAmount As String, colKind As String

    If docType = DOC_ORDER Then
        colType = "E": colDayNight = "AM": colUnit = "AP": colQty = "AZ"
        colLine = "BF": colMgr = "BP": colPrice = "CR": colAmount = "CY": colKind = "BX"
    Else
        colType = "O": colDayNight = "BH": colUnit = "BK": colQty = "CA"
        colLine = "DK": colMgr = "": colPrice = "CS": colAmount = "DA": colKind = "EC"
    End If

    Dim mgrSet As Object   ' Scripting.Dictionary
    If docType = DOC_ORDER Then
        Set mgrSet = BuildManagerRoomSet()
        If mgrSet Is Nothing Then GoTo Cleanup
        If mgrSet.Count = 0 Then
            MsgBox "基本情報 " & BASIC_INFO_BRANCH_CELL & "/" & BASIC_INFO_OFFICE_CELL & _
                   " に対応する管理室が管理室マスタに見つかりませんでした。" & vbCrLf & _
                   "支店名・出張所名をご確認ください。", vbExclamation
            GoTo Cleanup
        End If
    End If

    Dim n As Long
    n = lastRow - DATA_START_ROW + 1

    Dim vSeiri As Variant, vType As Variant, vDN As Variant, vUnit As Variant
    Dim vQty As Variant, vLine As Variant, vMgr As Variant
    Dim vPrice As Variant, vAmount As Variant, vKind As Variant

    vSeiri = ReadColumnValues(srcWs, "A", DATA_START_ROW, lastRow)
    vType = ReadColumnValues(srcWs, colType, DATA_START_ROW, lastRow)
    vDN = ReadColumnValues(srcWs, colDayNight, DATA_START_ROW, lastRow)
    vUnit = ReadColumnValues(srcWs, colUnit, DATA_START_ROW, lastRow)
    vQty = ReadColumnValues(srcWs, colQty, DATA_START_ROW, lastRow)
    vLine = ReadColumnValues(srcWs, colLine, DATA_START_ROW, lastRow)
    vPrice = ReadColumnValues(srcWs, colPrice, DATA_START_ROW, lastRow)
    vAmount = ReadColumnValues(srcWs, colAmount, DATA_START_ROW, lastRow)
    vKind = ReadColumnValues(srcWs, colKind, DATA_START_ROW, lastRow)
    If docType = DOC_ORDER Then vMgr = ReadColumnValues(srcWs, colMgr, DATA_START_ROW, lastRow)

    Dim worksRows As Collection, purchRows As Collection, weldRows As Collection
    Set worksRows = New Collection
    Set purchRows = New Collection
    Set weldRows = New Collection

    Dim seenMgr As Object
    Set seenMgr = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim seiriText As String, kindText As String, mgrOut As String, mraw As String
    For i = 1 To n
        seiriText = Trim$(CommonNzText(vSeiri(i)))
        If seiriText <> "" Then
            Dim keepRow As Boolean
            keepRow = True

            If docType = DOC_ORDER Then
                mgrOut = CommonNzText(vMgr(i))
                mraw = Trim$(mgrOut)
                If mraw <> "" Then
                    If Not seenMgr.Exists(mraw) Then seenMgr.Add mraw, True
                End If
                If Not mgrSet.Exists(CommonRemoveAllSpaces(mgrOut)) Then keepRow = False
            Else
                mgrOut = ""
            End If

            If keepRow Then
                kindText = CommonRemoveAllSpaces(CommonNzText(vKind(i)))
                If InStr(1, kindText, PURCHASE_KEYWORD) > 0 Then
                    Dim purchArr(0 To OUTPUT_COL_COUNT - 1) As Variant
                    purchArr(COL_VENDOR - 1) = ""
                    purchArr(COL_SEIRI - 1) = vSeiri(i)
                    purchArr(COL_TYPE - 1) = vType(i)
                    purchArr(COL_DAYNIGHT - 1) = vDN(i)
                    purchArr(COL_UNIT - 1) = vUnit(i)
                    purchArr(COL_QTY - 1) = vQty(i)
                    purchArr(COL_LINE - 1) = vLine(i)
                    purchArr(COL_MGR - 1) = mgrOut
                    purchArr(COL_JR_PRICE - 1) = vPrice(i)
                    purchArr(COL_JR_AMOUNT - 1) = vAmount(i)
                    purchArr(COL_OUT_PRICE - 1) = ""
                    purchArr(COL_OUT_AMOUNT - 1) = ""
                    purchArr(COL_KIND - 1) = vKind(i)
                    purchRows.Add purchArr
                ElseIf InStr(1, kindText, WELDING_KEYWORD) > 0 Then
                    Dim weldArr(0 To WELDING_OUTPUT_COL_COUNT - 1) As Variant
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_VENDOR)) = ""
                    weldArr(WELD_COL_TRACK_VENDOR - 1) = ""
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_SEIRI)) = vSeiri(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_TYPE)) = vType(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_DAYNIGHT)) = vDN(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_UNIT)) = vUnit(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_QTY)) = vQty(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_LINE)) = vLine(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_MGR)) = mgrOut
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_JR_PRICE)) = vPrice(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_JR_AMOUNT)) = vAmount(i)
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_OUT_PRICE)) = ""
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_OUT_AMOUNT)) = ""
                    weldArr(mod_Construction_OutputLayout.WeldingRowArrayIndex(COL_KIND)) = vKind(i)
                    weldRows.Add weldArr
                Else
                    Dim rowArr(0 To OUTPUT_COL_COUNT - 1) As Variant
                    rowArr(COL_VENDOR - 1) = ""
                    rowArr(COL_SEIRI - 1) = vSeiri(i)
                    rowArr(COL_TYPE - 1) = vType(i)
                    rowArr(COL_DAYNIGHT - 1) = vDN(i)
                    rowArr(COL_UNIT - 1) = vUnit(i)
                    rowArr(COL_QTY - 1) = vQty(i)
                    rowArr(COL_LINE - 1) = vLine(i)
                    rowArr(COL_MGR - 1) = mgrOut
                    rowArr(COL_JR_PRICE - 1) = vPrice(i)
                    rowArr(COL_JR_AMOUNT - 1) = vAmount(i)
                    rowArr(COL_OUT_PRICE - 1) = ""
                    rowArr(COL_OUT_AMOUNT - 1) = ""
                    rowArr(COL_KIND - 1) = vKind(i)
                    worksRows.Add rowArr
                End If
            End If
        End If
    Next i
    LogCI "工事側=" & worksRows.Count & " / 溶接側=" & weldRows.Count & " / 購入充当側=" & purchRows.Count

    Dim filteredOutAll As Boolean
    filteredOutAll = (docType = DOC_ORDER) And (worksRows.Count = 0) And _
                     (weldRows.Count = 0) And (purchRows.Count = 0)
    If filteredOutAll Then
        MsgBox "管理室フィルタで全ての行が除外されました(取込0件)。" & vbCrLf & vbCrLf & _
               "■基本情報(支店/出張所)に対応する管理室:" & vbCrLf & mod_Construction_OutputLayout.JoinKeys(mgrSet) & vbCrLf & vbCrLf & _
               "■取込データに含まれる管理室:" & vbCrLf & mod_Construction_OutputLayout.JoinKeys(seenMgr) & vbCrLf & vbCrLf & _
               "両者が違う地域の場合は、基本情報シートの支店・出張所と取込データの地域が" & _
               "一致しているかご確認ください。" & vbCrLf & _
               "「取込データに含まれる管理室」が(なし)の場合は、取込元の管理室セル位置(BP列)を" & _
               "ご確認ください。", vbExclamation, "取込み診断"
    End If

    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Set srcWb = Nothing: srcOpenedHere = False

    WriteReferenceValueToBasicInfo refValue, refValueSourceCell

    Dim wsWorks As Worksheet
    Dim wsWeld As Worksheet
    Dim wsPurch As Worksheet
    Dim wsActivate As Worksheet
    Set mLastCreatedImportSheet = GetImportSheetAnchorSheet()

    If worksRows.Count > 0 Then
        Set wsWorks = BuildConstructionOutputSheet( _
            mod_Construction_OutputLayout.BuildConstructionSheetName(sourceA3Text, False), _
            worksRows, docType, guidanceDocumentName, False)
        If wsActivate Is Nothing Then Set wsActivate = wsWorks
    End If

    If weldRows.Count > 0 Then
        Set wsWeld = BuildConstructionOutputSheet( _
            mod_Construction_OutputLayout.BuildConstructionSheetName(sourceA3Text, True), _
            weldRows, docType, guidanceDocumentName, True)
        If wsActivate Is Nothing Then Set wsActivate = wsWeld
    End If

    If purchRows.Count > 0 Then
        Dim purchName As String
        If docType = DOC_ORDER Then
            purchName = CommonPurchaseOrderOutputSheetName()
        Else
            purchName = CommonPurchaseNoticeOutputSheetName()
        End If
        Set wsPurch = CreateOrReplaceSheet(purchName)
        If Not wsPurch Is Nothing Then
            wsPurch.Tab.Color = RGB(233, 241, 123)   ' #E9F17B
            WriteRecordsToSheet wsPurch, purchRows
            ApplyPurchaseNoticeLayout wsPurch
            SortPurchaseSheet wsPurch, PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_KIND_COL
            WritePurchaseNoticeAdditionalHeaders wsPurch
            FillPurchaseUnitPrices wsPurch
            FormatPurchaseNoticeSheet wsPurch
            ApplyPurchaseNoticeColumnExclusions wsPurch
            WritePurchaseNoticeJrTotalRow wsPurch
            ApplyOutputSheetHeaderAutoFilter wsPurch, _
                PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_KIND_COL, _
                PURCHASE_NOTICE_AUTO_AMOUNT_COL, PURCHASE_NOTICE_PRICE_COMPARE_COL, _
                PURCHASE_NOTICE_PRICE_GUIDANCE_COL
            ApplyOutputSheetHeaderFreezePanes wsPurch
            If wsActivate Is Nothing Then Set wsActivate = wsPurch
        End If
    End If

    NormalizeManagedImportSheetOrder

    tWorks = tWorks + worksRows.Count
    tWeld = tWeld + weldRows.Count
    tPurch = tPurch + purchRows.Count
    If Not wsActivate Is Nothing Then Set lastSheet = wsActivate
    Exit Sub

Cleanup:
    Dim errNo As Long, errDesc As String
    errNo = Err.Number: errDesc = Err.Description

    On Error Resume Next
    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0

    If errNo <> 0 Then
        LogCI "取込みエラー [" & srcPath & "] Err " & errNo & ": " & errDesc
        MsgBox "ファイルの取込みでエラーが発生しました。" & vbCrLf & _
               "ファイル: " & srcPath & vbCrLf & _
               "Err " & errNo & ": " & errDesc, vbExclamation
    End If
End Sub

Public Function BuildConstructionOutputSheet(ByVal sheetName As String, _
                                              ByVal rows As Collection, _
                                              ByVal docType As Long, _
                                              ByVal guidanceDocumentName As String, _
                                              ByVal isWelding As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = CreateOrReplaceSheet(sheetName)
    If ws Is Nothing Then Exit Function

    ws.Tab.Color = RGB(255, 255, 0)
    If isWelding Then
        WriteWeldingRecordsToSheet ws, rows
    Else
        WriteRecordsToSheet ws, rows
    End If
    SortWorksSheet ws
    If isWelding Then
        WriteAdditionalHeadersAtColumns _
            ws, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_AUTO_PRICE), mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_AUTO_AMOUNT), _
            mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_COMPARE), True
    Else
        WriteAdditionalHeaders ws
    End If
    FillReferenceUnitPrices ws, guidanceDocumentName
    If isWelding Then
        FormatSheet ws, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_SEIRI)
        ApplyPriceGuidanceColumnLayoutAtColumns _
            ws, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_COMPARE), mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_GUIDANCE)
        mod_subcontractorselector.ApplySubcontractorDropdowns ws
    Else
        FormatSheet ws
        ApplyPriceGuidanceColumnLayout ws
        mod_subcontractorselector.ApplySubcontractorDropdowns ws
    End If
    mod_Construction_OutputFormat.ApplySanpaiRowRestrictionsCore ws
    RefreshOutputSheetVendorColumnColors ws, mod_Construction_LineMapping.GetLastDataRow(ws)

    If Not isWelding Then
        ws.Columns(COL_VENDOR).HorizontalAlignment = xlCenter
    End If

    If docType = DOC_NOTICE Then ws.Columns(mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_MGR)).Hidden = True

    ws.Range(ws.Cells(1, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_OUT_PRICE)), _
             ws.Cells(1, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_OUT_AMOUNT))).EntireColumn.Delete Shift:=xlToLeft

    WriteJrTotalRow ws
    RedrawOutputSheetDataBorders ws

    If isWelding Then
        ApplyWeldingOutputSheetColumnAlignment ws
    End If

    If isWelding Then
        ApplyOutputSheetHeaderAutoFilter ws, _
            mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_SEIRI), mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_KIND), _
            mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_AUTO_AMOUNT), mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_COMPARE), _
            mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_GUIDANCE)
    Else
        ApplyOutputSheetHeaderAutoFilter ws
    End If

    ApplyOutputSheetHeaderFreezePanes ws

    Set BuildConstructionOutputSheet = ws
End Function

Public Sub PrepareExistingOutputSheets(ByRef cancelImport As Boolean)
    cancelImport = False

    Dim existingSheets As Collection
    Set existingSheets = CollectExistingManagedOutputSheetNames()
    If existingSheets.Count = 0 Then Exit Sub

    Dim existingList As String
    Dim sheetName As Variant
    For Each sheetName In existingSheets
        existingList = existingList & "  ・" & CStr(sheetName) & vbCrLf
    Next sheetName

    Dim ans As VbMsgBoxResult
    ans = MsgBox("既に作成済みの取込シートがあります。" & vbCrLf & vbCrLf & _
                 existingList & vbCrLf & _
                 "これらのシートを消去してから取り込みますか？" & vbCrLf & _
                 "「はい」  : 上記の既存シートを消去してから取り込みます。" & vbCrLf & _
                 "「いいえ」: 既存シートを残したまま取り込みます。" & vbCrLf & _
                 "「キャンセル」: 取込みを中止します。", _
                 vbYesNoCancel + vbQuestion, "既存シートの確認")

    Select Case ans
        Case vbYes
            Dim prevDisplayAlerts As Boolean
            Dim prevScreenUpdating As Boolean
            prevDisplayAlerts = Application.DisplayAlerts
            prevScreenUpdating = Application.ScreenUpdating
            Application.DisplayAlerts = False
            Application.ScreenUpdating = True
            For Each sheetName In existingSheets
                DeleteSheetByName CStr(sheetName)
            Next sheetName
            Application.ScreenUpdating = prevScreenUpdating
            Application.DisplayAlerts = prevDisplayAlerts
            DoEvents
            mSuppressOverwritePrompt = True
        Case vbNo
            ' 既存シートはそのまま残す(同名で作成するシートのみ上書きされる)
        Case Else
            cancelImport = True
    End Select
End Sub

Public Sub DeleteSheetByName(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.worksheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
    End If
    On Error GoTo 0
End Sub

Public Function CollectExistingManagedOutputSheetNames() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsManagedImportOutputSheet(ws) Then result.Add ws.Name
    Next ws

    Set CollectExistingManagedOutputSheetNames = result
End Function

Public Function IsManagedImportOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    If mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
        IsManagedImportOutputSheet = True
        Exit Function
    End If

    If Not SheetNameHasConstructionOutputSuffix(ws.Name) Then Exit Function

    If mod_Construction_BasicTotals.IsConstructionOutputSheet(ws) Then
        IsManagedImportOutputSheet = True
        Exit Function
    End If

    IsManagedImportOutputSheet = _
        (mod_Construction_BasicTotals.FindHeaderColumn(ws, "整理番号") > 0) And _
        (mod_Construction_BasicTotals.FindHeaderColumn(ws, "JR金額") > 0) And _
        (mod_Construction_BasicTotals.FindHeaderColumn(ws, "工種分類") > 0)
End Function

Public Function SheetNameHasConstructionOutputSuffix(ByVal sheetName As String) As Boolean
    SheetNameHasConstructionOutputSuffix = _
        SheetNameEndsWithSuffixText(sheetName, CONSTRUCTION_SHEET_SUFFIX_WORKS) Or _
        SheetNameEndsWithSuffixText(sheetName, CONSTRUCTION_SHEET_SUFFIX_WELDING)
End Function

Public Function SheetNameEndsWithSuffixText(ByVal sheetName As String, ByVal suffixText As String) As Boolean
    Dim normalizedName As String
    Dim normalizedSuffix As String
    normalizedName = NormalizeSheetNameParentheses(sheetName)
    normalizedSuffix = NormalizeSheetNameParentheses(suffixText)

    If Len(normalizedName) < Len(normalizedSuffix) Then Exit Function
    SheetNameEndsWithSuffixText = _
        (StrComp(Right$(normalizedName, Len(normalizedSuffix)), normalizedSuffix, vbTextCompare) = 0)
End Function

Public Function NormalizeSheetNameParentheses(ByVal text As String) As String
    Dim t As String
    t = text
    t = Replace$(t, ChrW$(&HFF08), "(")
    t = Replace$(t, ChrW$(&HFF09), ")")
    NormalizeSheetNameParentheses = t
End Function

Public Function PickSourceFiles() As Collection
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "施工指示書・施工通知書ブックを選択してください(複数選択可)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel ブック", "*.xlsx; *.xlsm; *.xls"
        Dim root As String
        root = GetImportRootFolder()
        If root <> "" Then .InitialFileName = root & "\"
        If .Show = -1 Then
            Dim result As Collection
            Set result = New Collection
            Dim k As Long
            For k = 1 To .SelectedItems.Count
                result.Add .SelectedItems(k)
            Next k
            Set PickSourceFiles = result
        End If
    End With
End Function

Public Function PickSourceFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "施工指示書・施工通知書ブックを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック", "*.xlsx; *.xlsm; *.xls"
        Dim root As String
        root = GetImportRootFolder()
        If root <> "" Then .InitialFileName = root & "\"
        If .Show = -1 Then
            PickSourceFile = .SelectedItems(1)
        Else
            PickSourceFile = ""
        End If
    End With
End Function

Public Function GetImportRootFolder() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim docFolder As String
    docFolder = "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"

    Dim subPath As String
    subPath = "注文書_請求書VBA\施工指示書・施工通知書"

    Dim candidates As Collection
    Set candidates = New Collection

    Dim up As String
    up = Environ$("USERPROFILE")
    If Len(Trim$(up)) = 0 Then up = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(up)) > 0 Then
        candidates.Add up & "\" & CommonCompanyNameText() & "\" & docFolder & "\" & subPath
    End If
    candidates.Add "C:\Users\n-morimoto\" & CommonCompanyNameText() & "\" & docFolder & "\" & subPath

    Dim p As Variant
    For Each p In candidates
        If fso.FolderExists(CStr(p)) Then
            GetImportRootFolder = CStr(p)
            Exit Function
        End If
    Next p

    If candidates.Count > 0 Then GetImportRootFolder = CStr(candidates(1))
End Function

Public Function DetermineDocType(ByVal filePath As String) As Long
    Dim nm As String
    nm = Mid$(filePath, InStrRev(filePath, "\") + 1)

    If InStr(1, nm, "通知") > 0 Then
        DetermineDocType = DOC_NOTICE
    ElseIf InStr(1, nm, "指示") > 0 Then
        DetermineDocType = DOC_ORDER
    Else
        Dim ans As VbMsgBoxResult
        ans = MsgBox("ファイル名から種別を判定できませんでした。" & vbCrLf & _
                     "施工指示書として取り込みますか？" & vbCrLf & _
                     "(「いいえ」を選ぶと施工通知書として取り込みます)", _
                     vbYesNoCancel + vbQuestion, "取込種別の確認")
        Select Case ans
            Case vbYes: DetermineDocType = DOC_ORDER
            Case vbNo: DetermineDocType = DOC_NOTICE
            Case Else: DetermineDocType = 0
        End Select
    End If
End Function

Public Function ValidateBasicInfoForImport(ByVal docType As Long) As Boolean
    ValidateBasicInfoForImport = False

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりません。", vbExclamation
        Exit Function
    End If

    Dim chkCells(1 To 4) As String, chkLabels(1 To 4) As String, n As Long
    n = 0
    n = n + 1: chkCells(n) = BASIC_INFO_LINE_TYPE_CELL:    chkLabels(n) = "在来線・新幹線区分"
    n = n + 1: chkCells(n) = BASIC_INFO_PROJECT_NAME_CELL: chkLabels(n) = "工事種別"
    If docType = DOC_ORDER Then
        n = n + 1: chkCells(n) = BASIC_INFO_BRANCH_CELL:   chkLabels(n) = "支店"
        n = n + 1: chkCells(n) = BASIC_INFO_OFFICE_CELL:   chkLabels(n) = "出張所"
    End If

    Dim missingMsg As String, firstCell As String
    Dim i As Long, v As String
    For i = 1 To n
        v = Trim$(CommonNzText(wsInfo.Range(chkCells(i)).value))
        If v = "" Then
            missingMsg = missingMsg & "  ・" & chkCells(i) & "  " & chkLabels(i) & vbCrLf
            If firstCell = "" Then firstCell = chkCells(i)
        End If
    Next i

    If missingMsg <> "" Then
        On Error Resume Next
        Application.screenUpdating = True
        wsInfo.Activate
        wsInfo.Range(firstCell).Select
        On Error GoTo 0
        MsgBox "取込に必要な基本情報が未入力です。" & vbCrLf & vbCrLf & _
               "以下のセルを入力してから取り込んでください。" & vbCrLf & vbCrLf & _
               missingMsg, vbExclamation, "入力チェック"
        Exit Function
    End If

    ValidateBasicInfoForImport = True
End Function

Public Function BuildManagerRoomSet() As Object
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim BranchName As String, officeName As String
    BranchName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    officeName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    If BranchName = "" Or officeName = "" Then
        MsgBox "基本情報シート " & BASIC_INFO_BRANCH_CELL & "(支店) または " & _
               BASIC_INFO_OFFICE_CELL & "(出張所) が空です。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim masterPath As String
    Dim connection As Object
    Set connection = mod_Construction_OutputLayout.OpenUnitPriceMasterAdoConnection(masterPath)
    If connection Is Nothing Then
        MsgBox "出張所別_単価適用線区.xlsx が見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    On Error GoTo ErrorHandler

    Dim actualSheetName As String
    actualSheetName = mod_Construction_OutputLayout.FindAdoWorksheetName(connection, MGR_MASTER_SHEET)
    If actualSheetName = "" Then
        MsgBox "管理室マスタに「" & MGR_MASTER_SHEET & "」シートが見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        GoTo Cleanup
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT [F2], [F3], [F6] FROM " & _
                   mod_Construction_OutputLayout.BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

    Dim b As String, c As String, room As String
    Do Until recordset.EOF
        b = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(0).value))
        c = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(1).value))
        If b = BranchName And c = officeName Then
            room = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(2).value))
            If room <> "" Then
                If Not dict.Exists(room) Then dict.Add room, True
            End If
        End If
        recordset.MoveNext
    Loop

    LogCI "管理室集合 件数=" & dict.Count & " (支店=" & BranchName & " 出張所=" & officeName & ")"
    Set BuildManagerRoomSet = dict
    GoTo Cleanup

ErrorHandler:
    MsgBox "管理室マスタをADOで読み込めませんでした。" & vbCrLf & _
           masterPath & vbCrLf & Err.Description, vbExclamation
    Set BuildManagerRoomSet = Nothing

Cleanup:
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
End Function

Public Function ResolveMasterFilePath() As String
    Dim p As String
    On Error Resume Next
    p = mod_MaterialPriceImport.GetMasterFilePath()
    On Error GoTo 0
    If p <> "" Then
        ResolveMasterFilePath = p
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim up As String
    up = Environ$("USERPROFILE")
    If Len(Trim$(up)) = 0 Then up = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(up)) > 0 Then
        Dim cand As String
        cand = up & "\" & CommonCompanyNameText() & "\" & _
               "線路出張所用_注文書_請求書アクセスサイト - ドキュメント\マスタデータ\出張所別_単価適用線区.xlsx"
        If fso.FileExists(cand) Then ResolveMasterFilePath = cand
    End If
End Function

Public Function OpenWorkbookReadOnly(ByVal filePath As String, _
                                      ByRef openedHere As Boolean) As Workbook
    openedHere = False

    Dim nm As String
    nm = Mid$(filePath, InStrRev(filePath, "\") + 1)

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nm, vbTextCompare) = 0 Then
            Set OpenWorkbookReadOnly = wb
            Exit Function
        End If
    Next wb

    On Error Resume Next
    Set wb = Application.Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    If Not wb Is Nothing Then openedHere = True
    Set OpenWorkbookReadOnly = wb
End Function

Public Function ReadColumnValues(ByVal ws As Worksheet, ByVal col As String, _
                                  ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim out() As Variant, v As Variant, i As Long, cnt As Long, colIdx As Long
    cnt = r2 - r1 + 1
    ReDim out(1 To cnt)

    colIdx = ColLetterToNum(col)
    If colIdx = 0 Or colIdx > ws.Columns.Count Then
        LogCI "列[" & col & "](=" & colIdx & ")はシート範囲外(列数=" & ws.Columns.Count & ") -> 空白で取込"
        For i = 1 To cnt
            out(i) = ""
        Next i
        ReadColumnValues = out
        Exit Function
    End If

    If cnt = 1 Then
        out(1) = ws.Cells(r1, colIdx).value
    Else
        v = ws.Range(ws.Cells(r1, colIdx), ws.Cells(r2, colIdx)).value
        For i = 1 To cnt
            out(i) = v(i, 1)
        Next i
    End If
    ReadColumnValues = out
End Function

Public Function ColLetterToNum(ByVal col As String) As Long
    Dim i As Long, ch As String, result As Long
    col = UCase$(Trim$(col))
    If Len(col) = 0 Then
        ColLetterToNum = 0
        Exit Function
    End If
    result = 0
    For i = 1 To Len(col)
        ch = Mid$(col, i, 1)
        If ch < "A" Or ch > "Z" Then
            ColLetterToNum = 0
            Exit Function
        End If
        result = result * 26 + (Asc(ch) - Asc("A") + 1)
    Next i
    ColLetterToNum = result
End Function

Public Function GetImportSheetAnchorSheet() As Worksheet
    On Error Resume Next
    Set GetImportSheetAnchorSheet = Sheet1
    On Error GoTo 0

    If GetImportSheetAnchorSheet Is Nothing Then
        Set GetImportSheetAnchorSheet = CommonGetBasicInfoWorksheet(ThisWorkbook)
    End If
End Function

Public Function FindRightmostManagedImportSheetAfterAnchor( _
        ByVal anchorSheet As Worksheet) As Worksheet
    If anchorSheet Is Nothing Then Exit Function

    Dim ws As Worksheet
    Dim candidate As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Index > anchorSheet.Index Then
            If IsManagedImportOutputSheet(ws) Then
                If Not mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
                    If candidate Is Nothing Or ws.Index > candidate.Index Then
                        Set candidate = ws
                    End If
                End If
            End If
        End If
    Next ws

    Set FindRightmostManagedImportSheetAfterAnchor = candidate
End Function

Public Function ResolveImportSheetInsertAfter( _
        Optional ByVal replaceSheet As Worksheet = Nothing) As Worksheet
    If Not mLastCreatedImportSheet Is Nothing Then
        Set ResolveImportSheetInsertAfter = mLastCreatedImportSheet
        Exit Function
    End If

    If Not replaceSheet Is Nothing Then
        If replaceSheet.Index > 1 Then
            Set ResolveImportSheetInsertAfter = _
                ThisWorkbook.Worksheets(replaceSheet.Index - 1)
            Exit Function
        End If
    End If

    Dim anchorSheet As Worksheet
    Set anchorSheet = GetImportSheetAnchorSheet()

    Dim rightmostSheet As Worksheet
    Set rightmostSheet = FindRightmostManagedImportSheetAfterAnchor(anchorSheet)
    If Not rightmostSheet Is Nothing Then
        Set ResolveImportSheetInsertAfter = rightmostSheet
    ElseIf Not anchorSheet Is Nothing Then
        Set ResolveImportSheetInsertAfter = anchorSheet
    Else
        Set ResolveImportSheetInsertAfter = _
            ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
End Function

Public Function CreateOrReplaceSheet(ByVal sheetName As String) As Worksheet
    Dim existing As Worksheet
    Dim replaceSheet As Worksheet
    On Error Resume Next
    Set existing = ThisWorkbook.worksheets(sheetName)
    On Error GoTo 0

    If Not existing Is Nothing Then
        If Not mSuppressOverwritePrompt Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("シート「" & sheetName & "」は既に存在します。" & vbCrLf & _
                         "削除して作り直しますか？", vbYesNo + vbQuestion, "シートの上書き確認")
            If ans <> vbYes Then
                Set CreateOrReplaceSheet = Nothing
                Exit Function
            End If
        End If
        Set replaceSheet = existing
        existing.Delete
    End If

    Dim insertAfter As Worksheet
    Set insertAfter = ResolveImportSheetInsertAfter(replaceSheet)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.worksheets.Add(After:=insertAfter)
    ws.Name = sheetName
    Set mLastCreatedImportSheet = ws
    Set CreateOrReplaceSheet = ws
End Function

Public Sub NormalizeManagedImportSheetOrder()
    Dim anchorSheet As Worksheet
    Set anchorSheet = GetImportSheetAnchorSheet()
    If anchorSheet Is Nothing Then Exit Sub

    Dim orderedSheets As Collection
    Set orderedSheets = New Collection

    AppendOutputSheetsBySuffix orderedSheets, CONSTRUCTION_SHEET_SUFFIX_WORKS
    AppendOutputSheetsBySuffix orderedSheets, CONSTRUCTION_SHEET_SUFFIX_WELDING
    AppendPurchaseOutputSheets orderedSheets

    If orderedSheets.Count = 0 Then Exit Sub

    Dim prev As Worksheet
    Set prev = anchorSheet

    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In orderedSheets
        ws.Move After:=prev
        Set prev = ws
    Next ws
    On Error GoTo 0

    If Not prev Is Nothing Then Set mLastCreatedImportSheet = prev
End Sub

Public Sub AppendOutputSheetsBySuffix(ByVal target As Collection, ByVal suffixText As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsManagedImportOutputSheet(ws) Then
            If SheetNameEndsWithSuffixText(ws.Name, suffixText) Then
                target.Add ws
            End If
        End If
    Next ws
End Sub

Public Sub AppendPurchaseOutputSheets(ByVal target As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
            target.Add ws
        End If
    Next ws
End Sub

Public Sub WriteRecordsToSheet(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim headers As Variant
    headers = Array("施工業者", "整理番号", "工事種類", "昼夜別", "単位", "数量", _
                    "契約線区名", "管理室", "JR単価", "JR金額", "外注単価", "外注金額", "工種分類")
    ws.Range("A1").Resize(1, OUTPUT_COL_COUNT).value = headers

    If rows.Count = 0 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To rows.Count, 1 To OUTPUT_COL_COUNT)

    Dim i As Long, cc As Long, item As Variant
    i = 0
    For Each item In rows
        i = i + 1
        For cc = 1 To OUTPUT_COL_COUNT
            arr(i, cc) = item(cc - 1)
        Next cc
    Next item

    ws.Range("A2").Resize(rows.Count, OUTPUT_COL_COUNT).value = arr
End Sub

Public Sub WriteWeldingRecordsToSheet(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim headers As Variant
    headers = Array(WELDING_VENDOR_HEADER, TRACK_VENDOR_HEADER, "整理番号", "工事種類", "昼夜別", "単位", "数量", _
                    "契約線区名", "管理室", "JR単価", "JR金額", "外注単価", "外注金額", "工種分類")
    ws.Range("A1").Resize(1, WELDING_OUTPUT_COL_COUNT).value = headers

    If rows.Count = 0 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To rows.Count, 1 To WELDING_OUTPUT_COL_COUNT)

    Dim i As Long, cc As Long, item As Variant
    i = 0
    For Each item In rows
        i = i + 1
        For cc = 1 To WELDING_OUTPUT_COL_COUNT
            arr(i, cc) = item(cc - 1)
        Next cc
    Next item

    ws.Range("A2").Resize(rows.Count, WELDING_OUTPUT_COL_COUNT).value = arr
End Sub

Public Sub ApplyPurchaseNoticeLayout(ByVal ws As Worksheet)
    ws.Columns(COL_VENDOR).Delete Shift:=xlToLeft
    ws.Cells(1, PURCHASE_NOTICE_SEIRI_COL).value = "整理番号"
End Sub

Public Sub WriteAdditionalHeaders(ByVal ws As Worksheet, _
                                   Optional ByVal includePriceComparison As Boolean = True)
    WriteAdditionalHeadersAtColumns _
        ws, COL_AUTO_PRICE, COL_AUTO_AMOUNT, COL_PRICE_COMPARE, includePriceComparison
End Sub

Public Sub WritePurchaseNoticeAdditionalHeaders(ByVal ws As Worksheet)
    WriteAdditionalHeadersAtColumns _
        ws, PURCHASE_NOTICE_AUTO_PRICE_COL, PURCHASE_NOTICE_AUTO_AMOUNT_COL, _
        PURCHASE_NOTICE_PRICE_COMPARE_COL, True
End Sub

Public Sub WriteAdditionalHeadersAtColumns( _
    ByVal ws As Worksheet, _
    ByVal autoPriceColumn As Long, _
    ByVal autoAmountColumn As Long, _
    ByVal comparisonColumn As Long, _
    ByVal includePriceComparison As Boolean)

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つからないため、追加ヘッダーを設定できませんでした。", vbExclamation
        Exit Sub
    End If

    Dim amountName As String
    amountName = CommonNzText(wsInfo.Range(BASIC_INFO_AMOUNT_CELL).value)
    ws.Cells(1, autoPriceColumn).value = _
        CommonNzText(wsInfo.Range(BASIC_INFO_PUBLIC_CELL).value) & "公開" & amountName
    ws.Cells(1, autoAmountColumn).value = amountName & "金額"
    If includePriceComparison Then ws.Cells(1, comparisonColumn).value = "単価比較"
End Sub

Public Function GetOutputLastColumn( _
    ByVal ws As Worksheet, _
    Optional ByVal kindColumn As Long = COL_KIND, _
    Optional ByVal autoAmountColumn As Long = COL_AUTO_AMOUNT, _
    Optional ByVal comparisonColumn As Long = COL_PRICE_COMPARE, _
    Optional ByVal guidanceColumn As Long = COL_PRICE_GUIDANCE) As Long

    If CommonNzText(ws.Cells(1, comparisonColumn).value) <> "" Then
        GetOutputLastColumn = guidanceColumn
    ElseIf CommonNzText(ws.Cells(1, autoAmountColumn).value) <> "" Then
        GetOutputLastColumn = autoAmountColumn
    Else
        GetOutputLastColumn = kindColumn
    End If
End Function

