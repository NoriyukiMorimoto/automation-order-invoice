Attribute VB_Name = "mod_PrefectureSelector"
Option Explicit

' ========================================================
' mod_PrefectureSelector
' 基本情報シート C13 の都道府県選択（複数選択）。
'   ・マスタ: 出張所別_単価適用線区.xlsx の「都道府県」シート A2:A
'     （パスは mod_MaterialPriceImport.GetMasterFilePath を流用）
'   ・フォーム: frmPrefectureSelector（lstCompanies を MultiSelect 化）
'   ・区切り: ・（カタカナ中点）で連結し C13 に書き込む。
'   ・色塗り: mod_BasicInfoGuide に委譲（未入力=黄色 / 入力済=濃色）。
'   ・ダブルクリック直後のハング回避のため OnTime で UI 解放後に起動。
' ========================================================

Private Const PREF_CELL As String = "C13"
Private mScheduled As Boolean

' ダブルクリックからのエントリーポイント。
' イベント中にモーダルフォームを同期表示するとハングするため、OnTime で遅延起動する。
Public Sub RequestPrefectureSelection()
    If mScheduled Then Exit Sub
    mScheduled = True

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledPrefectureSelection", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:="mod_PrefectureSelector.RunScheduledPrefectureSelection", _
                           Schedule:=True
        If Err.Number <> 0 Then
            Err.Clear
            mScheduled = False
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub RunScheduledPrefectureSelection()
    mScheduled = False
    DoEvents
    SelectPrefectureForBasicInfo
End Sub

Public Sub SelectPrefectureForBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim names As Variant
    names = GetPrefectureList()
    If Not IsArray(names) Then
        MsgBox PrefectureListEmptyText(), vbExclamation
        Exit Sub
    End If

    Dim f As frmPrefectureSelector
    Set f = New frmPrefectureSelector
    f.SetSelectedPrefectures SplitPrefectureCell(wsInfo)
    f.SetPrefectures names
    f.Show vbModal

    Dim isConfirmed As Boolean
    isConfirmed = f.confirmed

    Dim selected As Collection
    Set selected = f.GetSelectedPrefectures()
    Unload f

    If Not isConfirmed Then Exit Sub
    If selected Is Nothing Then Exit Sub
    If selected.Count = 0 Then Exit Sub

    Dim joined As String
    joined = JoinPrefectures(selected)

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo CleanExit
    Application.EnableEvents = False

    Dim anchor As Range
    Set anchor = wsInfo.Range(PREF_CELL)
    If anchor.MergeCells Then Set anchor = anchor.mergeArea.Cells(1, 1)
    anchor.value = joined

    ' セルからはみ出す場合は縮小して表示（結合セルでは Excel 仕様上無効）
    On Error Resume Next
    wsInfo.Range(PREF_CELL).ShrinkToFit = True
    On Error GoTo CleanExit

    mod_BasicInfoGuide.OnCellChanged wsInfo, wsInfo.Range(PREF_CELL)

CleanExit:
    Application.EnableEvents = prevEvents
End Sub

' 「都道府県」シート A2 以降を読み込んで配列で返す。
Public Function GetPrefectureList() As Variant
    Dim filePath As String
    filePath = mod_MaterialPriceImport.GetMasterFilePath()
    If filePath = "" Then Exit Function

    Dim cn As Object
    Set cn = CommonOpenExcelAdoConnection(filePath)
    If cn Is Nothing Then Exit Function

    Dim rs As Object

    On Error GoTo Cleanup

    Dim sheetName As String
    sheetName = ResolvePrefectureSheetName(cn)
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
            ' A1 は見出し行のためスキップ
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
        GetPrefectureList = arr
    End If

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
End Function

Private Function ResolvePrefectureSheetName(ByVal cn As Object) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)
    If sheetNames Is Nothing Then Exit Function

    Dim targetName As String
    targetName = CommonNormalizeText(PrefectureSheetNameText())

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(CommonNormalizeText(CStr(sheetName)), targetName, vbTextCompare) = 0 Then
            ResolvePrefectureSheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

' C13 の現在値を「・」で分割して事前選択用の配列で返す。
Private Function SplitPrefectureCell(ByVal wsInfo As Worksheet) As Variant
    Dim anchor As Range
    Set anchor = wsInfo.Range(PREF_CELL)
    If anchor.MergeCells Then Set anchor = anchor.mergeArea.Cells(1, 1)

    Dim raw As String
    raw = Trim$(CommonNzText(anchor.value))
    If raw = "" Then Exit Function

    SplitPrefectureCell = Split(raw, PrefectureDelimiterText())
End Function

Private Function JoinPrefectures(ByVal selected As Collection) As String
    Dim result As String
    Dim item As Variant
    For Each item In selected
        If result <> "" Then result = result & PrefectureDelimiterText()
        result = result & CStr(item)
    Next item
    JoinPrefectures = result
End Function

' "・"（カタカナ中点）
Private Function PrefectureDelimiterText() As String
    PrefectureDelimiterText = ChrW$(&H30FB)
End Function

' "都道府県"
Private Function PrefectureSheetNameText() As String
    PrefectureSheetNameText = ChrW$(&H90FD) & ChrW$(&H9053) & ChrW$(&H5E9C) & ChrW$(&H770C)
End Function

' "都道府県マスタを読み込めませんでした。"
Private Function PrefectureListEmptyText() As String
    PrefectureListEmptyText = ChrW$(&H90FD) & ChrW$(&H9053) & ChrW$(&H5E9C) & ChrW$(&H770C) & _
                              ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H3092) & _
                              ChrW$(&H8AAD) & ChrW$(&H307F) & ChrW$(&H8FBC) & ChrW$(&H3081) & _
                              ChrW$(&H307E) & ChrW$(&H305B) & ChrW$(&H3093) & ChrW$(&H3067) & _
                              ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H3002)
End Function
