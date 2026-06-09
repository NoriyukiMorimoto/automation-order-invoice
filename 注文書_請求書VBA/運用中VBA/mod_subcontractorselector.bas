Attribute VB_Name = "mod_subcontractorselector"
Option Explicit

'==========================================================================
'  施工会社(施工業者)選択モジュール  mod_SubcontractorSelector
'    ・施工指示書/施工通知書シートのA列(施工業者)にドロップダウンを設定
'    ・候補は基本情報シートの行11・F列から3列おき(F11,I11,L11…)の社名
'    ・frmSubconSelector を使い、選択行(複数可)へ一括適用(Project_Number_Posting
'      の複数行一括選択の発想を踏襲)
'==========================================================================

' 出力シート列(A=施工業者, B=整理番号)
Private Const COL_VENDOR As Long = 1
Private Const COL_SEIRI As Long = 2
Private Const DATA_START_ROW As Long = 2

' 基本情報シートの業者ブロック(行11・F列から3列おき、最大20ブロック)
Private Const VENDOR_NAME_ROW As Long = 11
Private Const VENDOR_FIRST_COL As Long = 6      ' F列
Private Const VENDOR_STEP_COLS As Long = 3      ' F,I,L,O…
Private Const VENDOR_MAX_BLOCKS As Long = 20

' リストが長い(255文字超)場合に使う非表示の補助列
Private Const HELPER_COL As Long = 100

'==========================================================================
'  基本情報シートから施工会社リスト(1次元配列)を取得
'==========================================================================
Public Function GetSubcontractorList() As Variant
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim names As Collection
    Set names = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim k As Long, col As Long, nm As String
    For k = 0 To VENDOR_MAX_BLOCKS - 1
        col = VENDOR_FIRST_COL + k * VENDOR_STEP_COLS
        nm = Trim$(CommonNzText(wsInfo.Cells(VENDOR_NAME_ROW, col).value))
        If nm <> "" Then
            If Not seen.Exists(nm) Then
                seen.Add nm, True
                names.Add nm
            End If
        End If
    Next k

    If names.Count = 0 Then Exit Function

    Dim arr() As String, i As Long
    ReDim arr(1 To names.Count)
    For i = 1 To names.Count
        arr(i) = CStr(names(i))
    Next i
    GetSubcontractorList = arr
End Function

'==========================================================================
'  アクティブシート(施工指示書/施工通知書)のA列にドロップダウンを設定
'==========================================================================
Public Sub ApplySubcontractorDropdownsToActiveSheet()
    ApplySubcontractorDropdowns ActiveSheet
End Sub

'==========================================================================
'  指定シートのA列(B列に値がある行)にドロップダウンを設定
'==========================================================================
Public Sub ApplySubcontractorDropdowns(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim names As Variant
    names = GetSubcontractorList()
    If Not IsArray(names) Then
        MsgBox "基本情報シートに施工会社(F11,I11,L11…)が登録されていません。", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).Row
    If lastRow < DATA_START_ROW Then Exit Sub

    Dim listFormula As String
    listFormula = BuildValidationListFormula(ws, names)
    If listFormula = "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim r As Long
    For r = DATA_START_ROW To lastRow
        If Trim$(CommonNzText(ws.Cells(r, COL_SEIRI).value)) <> "" Then
            With ws.Cells(r, COL_VENDOR).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=listFormula
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowError = False
            End With
        End If
    Next r

    Application.EnableEvents = prevEvents
End Sub

'==========================================================================
'  入力規則の Formula1 を構築
'    255文字以内 かつ 社名にカンマが無い -> 直接リスト
'    それ以外 -> 非表示の補助列に縦並びで書き出して範囲参照
'==========================================================================
Private Function BuildValidationListFormula(ByVal ws As Worksheet, ByVal names As Variant) As String
    Dim joined As String, i As Long, hasComma As Boolean
    For i = LBound(names) To UBound(names)
        joined = joined & IIf(joined = "", "", ",") & CStr(names(i))
        If InStr(1, CStr(names(i)), ",") > 0 Then hasComma = True
    Next i

    If Len(joined) <= 255 And Not hasComma Then
        BuildValidationListFormula = joined
        Exit Function
    End If

    Dim n As Long
    n = UBound(names) - LBound(names) + 1

    Dim helper As Range
    Set helper = ws.Range(ws.Cells(1, HELPER_COL), ws.Cells(n, HELPER_COL))
    helper.ClearContents

    Dim arr2() As Variant
    ReDim arr2(1 To n, 1 To 1)
    For i = 1 To n
        arr2(i, 1) = CStr(names(LBound(names) + i - 1))
    Next i
    helper.value = arr2
    ws.Columns(HELPER_COL).Hidden = True

    BuildValidationListFormula = "=" & helper.Address(True, True)
End Function

'==========================================================================
'  選択行(複数可)に施工会社を一括適用(frmSubconSelector を使用)
'    使い方: A列を含む対象行を選択 -> 本マクロ実行 -> フォームで1社選択
'==========================================================================
Public Sub SelectSubcontractorForSelection()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "対象データがありません。", vbExclamation
        Exit Sub
    End If

    ' 対象行(選択範囲のうちB列に値がある行)を収集
    Dim targetRows As Collection
    Set targetRows = New Collection

    If TypeName(Selection) = "Range" Then
        Dim usedRange As Range
        Set usedRange = ws.Range(ws.Cells(DATA_START_ROW, COL_SEIRI), ws.Cells(lastRow, COL_SEIRI))

        Dim hit As Range
        On Error Resume Next
        Set hit = Application.Intersect(Selection.EntireRow, usedRange)
        On Error GoTo 0

        Dim c As Range
        If Not hit Is Nothing Then
            For Each c In hit.Cells
                If Trim$(CommonNzText(ws.Cells(c.Row, COL_SEIRI).value)) <> "" Then targetRows.Add c.Row
            Next c
        End If
    End If

    If targetRows.Count = 0 Then
        ' 選択が無効ならアクティブセルの行
        If ActiveCell.Row >= DATA_START_ROW And ActiveCell.Row <= lastRow Then
            If Trim$(CommonNzText(ws.Cells(ActiveCell.Row, COL_SEIRI).value)) <> "" Then targetRows.Add ActiveCell.Row
        End If
    End If

    If targetRows.Count = 0 Then
        MsgBox "施工会社を設定する行を選択してください(B列に整理番号がある行)。", vbExclamation
        Exit Sub
    End If

    Dim names As Variant
    names = GetSubcontractorList()
    If Not IsArray(names) Then
        MsgBox "基本情報シートに施工会社(F11,I11,L11…)が登録されていません。", vbExclamation
        Exit Sub
    End If

    Dim f As frmSubconSelector
    Set f = New frmSubconSelector
    f.SetCompanies names
    f.Show

    Dim confirmed As Boolean, chosen As String
    confirmed = f.confirmed
    chosen = f.SelectedCompany
    Unload f
    If Not confirmed Or chosen = "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim rIdx As Variant
    For Each rIdx In targetRows
        ws.Cells(CLng(rIdx), COL_VENDOR).value = chosen
    Next rIdx

    Application.EnableEvents = prevEvents

    MsgBox targetRows.Count & " 行に「" & chosen & "」を設定しました。", vbInformation
End Sub


