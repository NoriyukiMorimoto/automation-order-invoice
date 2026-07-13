Option Explicit

' 基本情報シート施工会社ブロックの2セクション選択セル(40/41行目)ダブルクリックで
' frmSubconSelector を2セクションモードで表示する。
'   行40: 有償/無償 + 支給材料(単一選択)   → 受注者用 F32(F32:V32 結合)
'   行41: 有償/無償 + 貸与品(なし排他の複数可) → 受注者用 F33(F33:V33 結合)
' 下段の候補は出張所別_単価適用線区.xlsx の各シートA列(ヘッダー除く)を ADO で読み込む。
' 対象列は F9(施工会社数)に応じ F/I/L/O/R/U/X/AA/AD/AG(3列おき・最大10社)。
' モーダルをイベント内で表示するとハングするため Application.OnTime で遅延起動する。
' 確定内容は「(ONの字形)有償/無償 + 全角空白 + 下段選択(複数は読点連結)」を
' 基本情報セルと受注者用シートの結合セルへ左詰めで書き込む。
' 改修履歴: CHANGELOG.md 参照

Private Const SUPPLY_ROW As Long = 40
Private Const LOAN_ROW As Long = 41

Private mScheduled As Boolean
Private mPendingSheetName As String
Private mPendingCellAddress As String
Private mPendingMasterSheet As String
Private mPendingAllowMulti As Boolean
Private mPendingAcceptanceCell As String
Private mPendingCaption As String

' --- 行40: 支給材料 --------------------------------------------------
Public Function IsSupplyMaterialTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    IsSupplyMaterialTarget = IsRowVendorTarget(wsInfo, target, SUPPLY_ROW)
End Function

Public Sub RequestSupplyMaterialSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
    ScheduleSelection wsInfo, target, SupplySheetNameText(), False, "F32", SupplyCaptionText()
End Sub

' --- 行41: 貸与品 ----------------------------------------------------
Public Function IsLoanItemTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    IsLoanItemTarget = IsRowVendorTarget(wsInfo, target, LOAN_ROW)
End Function

Public Sub RequestLoanItemSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
    ScheduleSelection wsInfo, target, LoanSheetNameText(), True, "F33", LoanCaptionText()
End Sub

' --- 対象判定(指定行・施工会社列) -----------------------------------
Private Function IsRowVendorTarget(ByVal wsInfo As Worksheet, ByVal target As Range, ByVal rowNo As Long) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)
    If topLeft.Row <> rowNo Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        If topLeft.Column = mod_Construction_BasicTotals.BasicInfoVendorColumn(i) Then
            IsRowVendorTarget = True
            Exit Function
        End If
    Next i
End Function

' --- OnTime 遅延起動 -------------------------------------------------
Private Sub ScheduleSelection(ByVal wsInfo As Worksheet, ByVal target As Range, _
                              ByVal masterSheet As String, ByVal allowMulti As Boolean, _
                              ByVal acceptanceCell As String, ByVal formCaption As String)
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If mScheduled Then Exit Sub

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    mPendingSheetName = wsInfo.Name
    mPendingCellAddress = targetArea.Cells(1, 1).Address(False, False)
    mPendingMasterSheet = masterSheet
    mPendingAllowMulti = allowMulti
    mPendingAcceptanceCell = acceptanceCell
    mPendingCaption = formCaption
    mScheduled = True

    mod_DebugLog.Log "[SupplyLoan] scheduled cell=" & mPendingCellAddress

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledSupplyLoanSelection", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:="mod_BasicInfoSupplyLoan.RunScheduledSupplyLoanSelection", _
                           Schedule:=True
        If Err.Number <> 0 Then
            Err.Clear
            mScheduled = False
        End If
    End If
    On Error GoTo 0
End Sub

' Application.OnTime から呼び出されるエントリーポイント(Public 必須)
Public Sub RunScheduledSupplyLoanSelection()
    mScheduled = False
    DoEvents
    ShowForPendingCell
End Sub

Private Sub ShowForPendingCell()
    If Len(mPendingSheetName) = 0 Then Exit Sub
    If Len(mPendingCellAddress) = 0 Then Exit Sub

    Dim wsInfo As Worksheet
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets(mPendingSheetName)
    On Error GoTo 0
    If wsInfo Is Nothing Then Exit Sub

    Dim targetCell As Range
    On Error Resume Next
    Set targetCell = wsInfo.Range(mPendingCellAddress)
    On Error GoTo 0

    Dim masterSheet As String, allowMulti As Boolean, acceptanceCell As String, formCaption As String
    masterSheet = mPendingMasterSheet
    allowMulti = mPendingAllowMulti
    acceptanceCell = mPendingAcceptanceCell
    formCaption = mPendingCaption

    mPendingSheetName = ""
    mPendingCellAddress = ""

    If targetCell Is Nothing Then Exit Sub

    ShowSelection wsInfo, targetCell, masterSheet, allowMulti, acceptanceCell, formCaption
End Sub

Public Sub ShowSelection(ByVal wsInfo As Worksheet, ByVal targetCell As Range, _
                         ByVal masterSheet As String, ByVal allowMulti As Boolean, _
                         ByVal acceptanceCell As String, ByVal formCaption As String)
    If wsInfo Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    Dim items As Variant
    items = GetMasterColumnAList(masterSheet)
    If Not IsArray(items) Then
        MsgBox MasterEmptyText(), vbExclamation
        Exit Sub
    End If

    Dim anchor As Range
    Set anchor = targetCell
    On Error Resume Next
    If anchor.MergeCells Then Set anchor = anchor.MergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim topInit As String, bottomInit As Variant
    ParseCurrent CommonNzText(anchor.Value), topInit, bottomInit

    Dim f As frmSubconSelector
    Set f = New frmSubconSelector
    f.ConfigureTwoSectionMode formCaption, PaidText(), FreeText(), items, allowMulti, NoneText(), topInit, bottomInit, NoneText()
    f.Show vbModal

    Dim isConfirmed As Boolean, topSel As String, bottomSel As String
    isConfirmed = f.confirmed
    topSel = f.resultTop
    bottomSel = f.resultBottom

    Unload f
    Set f = Nothing

    If Not isConfirmed Then Exit Sub

    WriteSelection wsInfo, anchor, topSel, bottomSel, acceptanceCell
End Sub

Private Sub WriteSelection(ByVal wsInfo As Worksheet, ByVal anchor As Range, _
                           ByVal topSel As String, ByVal bottomSel As String, ByVal acceptanceCell As String)
    Dim combined As String
    If StrComp(topSel, NoneText(), vbTextCompare) = 0 Then
        combined = NoneText()
    Else
        combined = ChrW$(&H2611) & topSel & ChrW$(&H3000) & bottomSel
    End If

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo CleanExit
    Application.EnableEvents = False

    anchor.Value = combined
    anchor.HorizontalAlignment = xlLeft
    ' 施工会社列セルは幅に収まるよう縮小表示(結合セルは Excel 仕様上無効のため無視)
    On Error Resume Next
    anchor.ShrinkToFit = False
    anchor.WrapText = True
    anchor.EntireColumn.AutoFit
    anchor.EntireRow.AutoFit
    On Error GoTo CleanExit

    Dim ws As Worksheet
    Set ws = ResolveAcceptanceSheetForColumn(wsInfo, anchor.Column)
    If Not ws Is Nothing Then
        Dim cell As Range
        Set cell = ws.Range(acceptanceCell)
        If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
        cell.Value = combined
        cell.HorizontalAlignment = xlLeft
    End If

    mod_DebugLog.Log "[SupplyLoan] applied cell=" & anchor.Address(False, False) & " value=" & combined

CleanExit:
    Application.EnableEvents = prevEvents
End Sub

' 既存セル値から上段(有償/無償)と下段(配列)を復元する
Private Sub ParseCurrent(ByVal current As String, ByRef topInit As String, ByRef bottomInit As Variant)
    topInit = ""
    Dim fw As String
    fw = ChrW$(&H3000)
    Dim leftPart As String, rightPart As String
    Dim p As Long
    p = InStr(current, fw)
    If p > 0 Then
        leftPart = Left$(current, p - 1)
        rightPart = Mid$(current, p + 1)
    Else
        leftPart = current
        rightPart = ""
    End If
    leftPart = Replace$(leftPart, ChrW$(&H2611), "")
    leftPart = Replace$(leftPart, ChrW$(&H2610), "")
    topInit = Trim$(leftPart)
    If rightPart <> "" Then bottomInit = Split(rightPart, ChrW$(&H3001))
End Sub

' 会社列(値列)から業者インデックスを求め、対応する受注者用(略称)シートを解決する
Private Function ResolveAcceptanceSheetForColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Worksheet
    Dim vendorIndex As Long
    vendorIndex = mod_VendorMaster.GetVendorIndexFromValueColumnPublic(valueColumn)
    If vendorIndex < 1 Then Exit Function

    Dim companyName As String
    companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
    If companyName = "" Then Exit Function

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).Value))

    Dim vendorName As String, aliasText As String, workText As String
    If Not mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then Exit Function
    If aliasText = "" Then Exit Function

    Dim sheetName As String
    sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If Not mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then Exit Function

    Set ResolveAcceptanceSheetForColumn = ThisWorkbook.Worksheets(sheetName)
End Function

' 出張所別_単価適用線区.xlsx の指定シートA列(ヘッダー除く・重複除去)を配列で返す
Private Function GetMasterColumnAList(ByVal sheetNameText As String) As Variant
    Dim filePath As String
    filePath = mod_MaterialPriceImport.GetMasterFilePath()
    If filePath = "" Then Exit Function

    Dim cn As Object
    Set cn = CommonOpenExcelAdoConnection(filePath)
    If cn Is Nothing Then Exit Function

    Dim rs As Object
    On Error GoTo Cleanup

    Dim sheetName As String
    sheetName = ResolveMasterSheetName(cn, sheetNameText)
    If sheetName = "" Then GoTo Cleanup

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & Replace$(sheetName, "]", "]]") & "$A:A]", cn, 0, 1, 1

    Dim names As Collection
    Set names = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim isHeader As Boolean
    isHeader = True
    Dim nm As String
    Do Until rs.EOF
        If isHeader Then
            isHeader = False
        Else
            nm = CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
            If nm <> "" Then
                If Not seen.Exists(nm) Then
                    seen.Add nm, True
                    names.Add nm
                End If
            End If
        End If
        rs.MoveNext
    Loop

    If names.Count > 0 Then
        Dim arr() As String
        Dim i As Long
        ReDim arr(0 To names.Count - 1)
        For i = 1 To names.Count
            arr(i - 1) = CStr(names(i))
        Next i
        GetMasterColumnAList = arr
    End If

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
End Function

Private Function ResolveMasterSheetName(ByVal cn As Object, ByVal sheetNameText As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)
    If sheetNames Is Nothing Then Exit Function

    Dim targetName As String
    targetName = CommonNormalizeText(sheetNameText)

    Dim sName As Variant
    For Each sName In sheetNames
        If StrComp(CommonNormalizeText(CStr(sName)), targetName, vbTextCompare) = 0 Then
            ResolveMasterSheetName = CStr(sName)
            Exit Function
        End If
    Next sName
End Function

' "支給材料"
Private Function SupplySheetNameText() As String
    SupplySheetNameText = ChrW$(&H652F) & ChrW$(&H7D66) & ChrW$(&H6750) & ChrW$(&H6599)
End Function

' "貸与品"
Private Function LoanSheetNameText() As String
    LoanSheetNameText = ChrW$(&H8CB8) & ChrW$(&H4E0E) & ChrW$(&H54C1)
End Function

' "有償"
Private Function PaidText() As String
    PaidText = ChrW$(&H6709) & ChrW$(&H511F)
End Function

' "無償"
Private Function FreeText() As String
    FreeText = ChrW$(&H7121) & ChrW$(&H511F)
End Function

' "なし"
Private Function NoneText() As String
    NoneText = ChrW$(&H306A) & ChrW$(&H3057)
End Function

' "支給材料 有償・無償選択"
Private Function SupplyCaptionText() As String
    SupplyCaptionText = SupplySheetNameText() & " " & PaidText() & ChrW$(&H30FB) & FreeText() & _
                        ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "貸与品 有償・無償選択"
Private Function LoanCaptionText() As String
    LoanCaptionText = LoanSheetNameText() & " " & PaidText() & ChrW$(&H30FB) & FreeText() & _
                      ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "マスタ(支給材料/貸与品)を読み込めませんでした。"
Private Function MasterEmptyText() As String
    MasterEmptyText = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H3092) & _
                      ChrW$(&H8AAD) & ChrW$(&H307F) & ChrW$(&H8FBC) & ChrW$(&H3081) & _
                      ChrW$(&H307E) & ChrW$(&H305B) & ChrW$(&H3093) & ChrW$(&H3067) & _
                      ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H3002)
End Function

' 施工会社名の変更/受注者用シート再生成時に、基本情報 F40/F41 の現在値から
' 受注者用シート F32/F33 へ再転記する(ApplyContractorHeader から呼ぶ)。
Public Sub ReapplySupplyLoan(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim vc As Long
    vc = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    Dim ws As Worksheet
    Set ws = ResolveAcceptanceSheetForColumn(wsInfo, vc)
    If ws Is Nothing Then Exit Sub

    ReapplyStringCell ws, wsInfo.Cells(SUPPLY_ROW, vc), "F32"
    ReapplyStringCell ws, wsInfo.Cells(LOAN_ROW, vc), "F33"
End Sub

Private Sub ReapplyStringCell(ByVal ws As Worksheet, ByVal srcCell As Range, ByVal targetCell As String)
    Dim cur As String
    cur = CommonNzText(srcCell.mergeArea.Cells(1, 1).value)
    Dim c As Range
    Set c = ws.Range(targetCell)
    On Error Resume Next
    If c.MergeCells Then Set c = c.mergeArea.Cells(1, 1)
    On Error GoTo 0
    c.value = cur
    c.HorizontalAlignment = xlLeft
End Sub
