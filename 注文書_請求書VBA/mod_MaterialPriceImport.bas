Option Explicit

'==========================================================================
'  購入充当材料単価インポートモジュール
'    改修内容：
'      #4  : PasteFromFileFast の中間配列詰め替えループを廃止し、
'            読み込み時点で必要範囲だけを取得して dataArr をそのまま貼付。
'      #5  : ClearTargetData の H/M 列数式投入を、1 行書込＋FillDown に
'            変更。セル単位ループを廃止。
'      #6  : マスター（単価適用線区）の Workbook を 1 回開いて
'            GetSenku / LoadUnitPriceProjectNames で共用する経路を新設。
'            また GetSenku 内部は C 列線形ループを Match→失敗時のみ
'            ループへフォールバックする実装に。
'      #9  : 共通関数（年抽出・日本語名・基本情報シート取得）を mod_Common
'            経由で取得するよう統一。
'      #13 : hasData 判定を Cells.Find から WorksheetFunction.CountA に置換。
'      #14 : GetMasterFilePath のハードコードを ChrW 構築 + Environ で再定義。
'      #15 : C25 変更時の工事件名リスト取得を Workbooks.Open から ADO に変更。
'==========================================================================

' 工事番号データを保持する共有変数
Public SharedMasterData As Variant

Private Const UNIT_PRICE_TABLE As String = "★購入充当材料単価表"
Private Const MASTER_SHEET_NAME As String = "単価適用線区"
Private Const HAYAKI_FOLDER As String = "早期発注"
Private Const SEKKEI_FOLDER As String = "設計変更"
Private Const ZAIRAISEN_FOLDER As String = "在来線"
Private Const SHINKANSEN_FOLDER As String = "新幹線"
Private Const TANKA_SHEET_NAME As String = "年初単価(購入充当)"
Private Const SEKKEI_TANKA_SHEET_NAME As String = "設計変更単価(購入充当)"
Private Const TANKA_DISPLAY_NAME As String = "年初単価(購入充当)"
Private Const SEKKEI_DISPLAY_NAME As String = "設計変更単価(購入充当)"
Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C25"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C26"
Private Const PROJECT_NAME_LIST_COL As String = "AE"
Private Const LINE_TYPE_LIST_COL As String = "AF"
Private Const PROJECT_NAME_LIST_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_LAST_ROW As Long = 1048576
'--------------------------------------------------------------------------
'  マスタファイルパス
'    改修（#14）：他モジュールに合わせて Environ + ChrW 構築に統一。
'--------------------------------------------------------------------------
Public Function GetMasterFilePath() As String
    Dim sep As String
    sep = Application.PathSeparator

    GetMasterFilePath = Environ$("USERPROFILE") & sep & _
                        CommonCompanyNameText() & sep & _
                        OrderInvoiceDocumentFolderText() & sep & _
                        UnitPriceMasterFolderText() & sep & _
                        UnitPriceMasterReferenceFolderText() & sep & _
                        UnitPriceMasterFileNameText()
End Function

Public Sub RefreshUnitPriceProjectNameValidation(Optional ByVal wsInfo As Worksheet, _
                                                  Optional ByVal keepProjectName As Boolean = True)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    WriteUnitPriceLineTypeValidation wsInfo

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    If lineType = "" Then
        ClearUnitPriceProjectNameValidation wsInfo, Not keepProjectName
        Exit Sub
    End If

    Dim projectNames As Collection
    Set projectNames = LoadUnitPriceProjectNames(GetMasterFilePath, lineType)

    If projectNames Is Nothing Then
        ClearUnitPriceProjectNameValidation wsInfo, True
        Exit Sub
    End If

    If projectNames.Count = 0 Then
        ClearUnitPriceProjectNameValidation wsInfo, True
        Exit Sub
    End If

    WriteUnitPriceProjectNameValidation wsInfo, projectNames, keepProjectName
    Exit Sub

ErrorHandler:
    ClearUnitPriceProjectNameValidation wsInfo, True
End Sub

Private Function LoadUnitPriceProjectNames(ByVal sourceFilePath As String, _
                                           ByVal lineType As String) As Collection
    Dim sheetName As String

    On Error GoTo ErrorHandler

    sheetName = ResolveUnitPriceProjectNameMasterSheetName(lineType)
    If sheetName = "" Then Exit Function

    Set LoadUnitPriceProjectNames = LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName)
    Exit Function

ErrorHandler:
    Set LoadUnitPriceProjectNames = Nothing
End Function

Private Function ResolveUnitPriceProjectNameMasterSheetName(ByVal lineType As String) As String
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_FOLDER
            ResolveUnitPriceProjectNameMasterSheetName = CommonUnitPriceProjectNameMasterSheetNameText()
        Case SHINKANSEN_FOLDER
            ResolveUnitPriceProjectNameMasterSheetName = ShinkansenUnitPriceProjectNameMasterSheetNameText()
    End Select
End Function

' C25 変更時にマスタブックを画面表示しないよう、工事件名リストは ADO で読む。
Private Function LoadUnitPriceProjectNamesByAdo(ByVal sourceFilePath As String, _
                                                ByVal sheetName As String) As Collection
    Dim cn As Object
    Dim rs As Object
    Dim result As Collection

    On Error GoTo ErrorHandler

    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Dim sql As String
    sql = "SELECT F1 FROM " & BuildAdoSheetRangeName(sheetName, "A", PROJECT_NAME_MASTER_START_ROW, "A", PROJECT_NAME_MASTER_LAST_ROW) & _
          " WHERE F1 IS NOT NULL"

    Set rs = cn.Execute(sql)
    Set result = New Collection

    Do Until rs.EOF
        Dim itemText As String
        itemText = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If itemText <> "" Then result.Add itemText
        rs.MoveNext
    Loop

    Set LoadUnitPriceProjectNamesByAdo = result

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Set LoadUnitPriceProjectNamesByAdo = Nothing
    Resume Cleanup
End Function

Private Function BuildAdoSheetRangeName(ByVal sheetName As String, _
                                        ByVal startCol As String, _
                                        ByVal startRow As Long, _
                                        ByVal endCol As String, _
                                        ByVal endRow As Long) As String
    BuildAdoSheetRangeName = "[" & Replace$(sheetName, "]", "]]") & "$" & _
                             startCol & CStr(startRow) & ":" & endCol & CStr(endRow) & "]"
End Function

Private Sub WriteUnitPriceLineTypeValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(LINE_TYPE_LIST_COL & PROJECT_NAME_LIST_START_ROW).Resize(2, 1)

    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).value = ZAIRAISEN_FOLDER
    listRange.Cells(2, 1).value = SHINKANSEN_FOLDER

    ResetUnitPriceProjectNameValidation wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL), listRange
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                ByVal projectNames As Collection, _
                                                ByVal keepProjectName As Boolean)
    Dim maxProjectRows As Long
    maxProjectRows = Application.Max(1, projectNames.Count)

    Dim listRange As Range
    Set listRange = wsInfo.Range(PROJECT_NAME_LIST_COL & PROJECT_NAME_LIST_START_ROW).Resize(maxProjectRows, 1)

    Dim currentProjectName As String
    currentProjectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value))

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents

    Dim i As Long
    For i = 1 To projectNames.Count
        listRange.Cells(i, 1).value = projectNames(i)
    Next i

    ResetUnitPriceProjectNameValidation wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL), _
                                        listRange.Resize(Application.Max(1, projectNames.Count))

    If Not keepProjectName Or currentProjectName = "" Or Not CollectionContainsText(projectNames, currentProjectName) Then
        wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    End If

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
End Sub

Private Sub ResetUnitPriceProjectNameValidation(ByVal targetCell As Range, ByVal listRange As Range)
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Sub ClearUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                Optional ByVal clearProjectName As Boolean = True)
    On Error Resume Next
    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents
    wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Validation.Delete
    If clearProjectName Then wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
    On Error GoTo 0
End Sub

Private Function CollectionContainsText(ByVal values As Collection, ByVal searchText As String) As Boolean
    If values Is Nothing Then Exit Function

    Dim item As Variant
    For Each item In values
        If StrComp(CommonNormalizeText(CStr(item)), CommonNormalizeText(searchText), vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next item
End Function

Private Function ShinkansenUnitPriceProjectNameMasterSheetNameText() As String
    Static cached As String
    If cached = "" Then cached = SHINKANSEN_FOLDER & "_" & CommonUnitPriceProjectNameMasterSheetNameText()
    ShinkansenUnitPriceProjectNameMasterSheetNameText = cached
End Function
Private Function GetBasicInfoOfficeName(ByVal ws As Worksheet) As String
    GetBasicInfoOfficeName = Trim(CStr(ws.Range(BASIC_INFO_OFFICE_CELL).value))
End Function

Private Function GetBasicInfoNendo(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetBasicInfoNendo = CLng(ws.Range(BASIC_INFO_YEAR_CELL).value)
    On Error GoTo 0
End Function

Private Function GetFormulaSheetRef(ByVal sheetName As String) As String
    GetFormulaSheetRef = "'" & Replace(sheetName, "'", "''") & "'!"
End Function

Private Function BuildMissingUnitPriceDataMessage(ByVal nendo As Long, _
                                                   ByVal displayName As String) As String
    BuildMissingUnitPriceDataMessage = CStr(nendo) & "年度の" & displayName & "データはまだ存在しません。"
End Function

Public Sub ClearAndImportUnitPriceForBasicInfo(ByVal ws As Worksheet)
    ClearUnitPriceSheets ws.Parent
    RunF1Process ws

    Dim completeMsg As String
    completeMsg = BuildCompleteMessage(ws.Parent)
    If completeMsg <> "" Then MsgBox completeMsg, vbInformation, "完了"
End Sub

'===========================================================
' 指定された列のデータを7行目から最終行までクリアする
'  改修（#5）：H/M 列の数式設定をセル単位ループから先頭セル書込＋FillDown
'              に変更し、数千行でも一瞬で展開できるようにする。
'===========================================================
Public Sub ClearTargetData(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim col As Variant

    ' 原図シート以外からの呼び出しは処理しない
    If ws.Name <> "原図" Then Exit Sub

    ' D列を基準に最終行を判定
    lastRow = ws.Cells(ws.rows.Count, "D").End(xlUp).Row
    If lastRow < 7 Then Exit Sub

    On Error Resume Next
    ' 【修正箇所】指定された列（A, C, G, I, J, L, T）をクリア
    ' ※N列?Q列（4列）が挿入されたため、元の「P列」を「T列」に変更しています。
    For Each col In Array("A", "C", "G", "I", "J", "L", "T")
        ws.Range(ws.Cells(7, col), ws.Cells(lastRow, col)).ClearContents
    Next col
    On Error GoTo 0

    ' H列・M列に数式を入力（これらはN列より左側のため変更不要）
    ' 改修：相対参照のまま 1 行目（7 行目）にだけ数式を書いて FillDown する。
    Dim hFormula As String
    hFormula = "=IF(A7=0,""""," & _
               "VLOOKUP($A7," & GetFormulaSheetRef(TANKA_SHEET_NAME) & "$A:$J,6,FALSE))"

    Dim mFormula As String
    mFormula = "=IF($A7=0,""""," & "ROUNDDOWN(L7*$I7,0))"

    ws.Range("H7").formula = hFormula
    ws.Range("M7").formula = mFormula

    If lastRow > 7 Then
        ws.Range("H7:H" & lastRow).FillDown
        ws.Range("M7:M" & lastRow).FillDown
    End If

    ' D2・D3・E3もクリア
    Application.EnableEvents = False
    ws.Range("D2").value = ""
    ws.Range("D2").Validation.Delete
    ws.Range("D3").value = ""
    ws.Range("E3").value = ""
    Application.EnableEvents = True

    ' 工事番号キャッシュをリセット
    SharedMasterData = Empty

    ' シート名を「原図」に戻す
    If ws.Name <> "原図" Then
        On Error Resume Next
        ws.Name = "原図"
        If Err.Number <> 0 Then
            MsgBox "シート名を「原図」に戻せませんでした。" & vbCrLf & _
                   "同名のシートが既に存在する可能性があります。", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' 単価シート・設計変更単価シートをクリア
    Call ClearUnitPriceSheets(ws.Parent)
End Sub

'===========================================================
' 単価・設計変更単価シートの内容をクリア（シートは削除しない）
'===========================================================
Private Sub ClearUnitPriceSheets(ByVal wb As Workbook)
    Dim sheetNames(1) As String
    sheetNames(0) = TANKA_SHEET_NAME
    sheetNames(1) = SEKKEI_TANKA_SHEET_NAME

    Dim i As Integer
    Dim ws As Worksheet
    For i = 0 To 1
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.EnableEvents = False
            ws.Cells.Clear
            Application.EnableEvents = True
            Set ws = Nothing
        End If
    Next i
End Sub

'===========================================================
' ボタン用：単価シートへの早期発注データ取込
'===========================================================
Public Sub ImportSoukiData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim officeName As String
    Dim nendo As Long
    officeName = GetBasicInfoOfficeName(ws)
    If officeName = "" Then
        MsgBox "C6に出張所名を入力してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    nendo = GetBasicInfoNendo(ws)
    On Error GoTo 0
    If nendo = 0 Then
        MsgBox "B4に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    Dim wsTanka As Worksheet
    If Not GetOrErrorSheet(ws.Parent, TANKA_SHEET_NAME, wsTanka) Then Exit Sub

    ' 改修（#13）：Cells.Find は ActiveCell 位置に依存することがあるため
    ' CountA に変更（空セル/数式/エラーすべて反映される）
    Dim hasData As Boolean
    hasData = (Application.WorksheetFunction.CountA(wsTanka.Cells) > 0)

    Dim b1c1 As String
    b1c1 = Trim(CStr(ws.Range(BASIC_INFO_YEAR_CELL).value)) & Trim(CStr(ws.Range(BASIC_INFO_OFFICE_CELL).value))
    If hasData Then
        Dim tankaHeader As String
        tankaHeader = GetSheetHeaderText(wsTanka.Parent, TANKA_SHEET_NAME)
        Dim resp As VbMsgBoxResult
        resp = MsgBox("既に「" & tankaHeader & "」のデータが入力されていますが、取込を実行してもいいですか？", _
                      vbOKCancel + vbExclamation, "確認")
        If resp = vbCancel Then Exit Sub
    Else
        MsgBox b1c1 & "_早期発注購入充当データを取り込みます。", vbInformation, "取込開始"
    End If

    Dim shisha As String
    Dim tankaFolder As String
    Dim isShinkansen As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo CleanupSouki

    Dim hayakiPath As String
    hayakiPath = BuildFolderPath(nendo, HAYAKI_FOLDER, isShinkansen, shisha, tankaFolder)
    If hayakiPath = "" Then
        MsgBox BuildMissingUnitPriceDataMessage(nendo, TANKA_DISPLAY_NAME), vbExclamation
        GoTo CleanupSouki
    End If

    wsTanka.Cells.Clear
    Dim soukiOk As Boolean
    soukiOk = ImportUnitPriceFiles(hayakiPath, wsTanka, "TableStyleMedium2")
    If Not soukiOk Then
        MsgBox BuildMissingUnitPriceDataMessage(nendo, TANKA_DISPLAY_NAME), vbExclamation
        GoTo CleanupSouki
    End If

CleanupSouki:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    If soukiOk Then
        Dim msgS As String
        msgS = BuildCompleteMessage(ws.Parent, True, False)
        If msgS <> "" Then MsgBox msgS, vbInformation, "完了"
    End If
End Sub

'===========================================================
' ボタン用：設計変更単価シートへの設計変更データ取込
'===========================================================
Public Sub ImportSekkeiData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim officeName As String
    Dim nendo As Long
    officeName = GetBasicInfoOfficeName(ws)
    If officeName = "" Then
        MsgBox "C6に出張所名を入力してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    nendo = GetBasicInfoNendo(ws)
    On Error GoTo 0
    If nendo = 0 Then
        MsgBox "B4に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    Dim wsSekkei As Worksheet
    If Not GetOrErrorSheet(ws.Parent, SEKKEI_TANKA_SHEET_NAME, wsSekkei) Then Exit Sub

    ' 改修（#13）：Cells.Find を CountA に置換
    Dim hasData As Boolean
    hasData = (Application.WorksheetFunction.CountA(wsSekkei.Cells) > 0)

    Dim b1c1 As String
    b1c1 = Trim(CStr(ws.Range(BASIC_INFO_YEAR_CELL).value)) & Trim(CStr(ws.Range(BASIC_INFO_OFFICE_CELL).value))
    If hasData Then
        Dim sekkeiHeader As String
        sekkeiHeader = GetSheetHeaderText(wsSekkei.Parent, SEKKEI_TANKA_SHEET_NAME)
        Dim resp As VbMsgBoxResult
        resp = MsgBox("既に「" & sekkeiHeader & "」のデータが入力されていますが、取込を実行してもいいですか？", _
                      vbOKCancel + vbExclamation, "確認")
        If resp = vbCancel Then Exit Sub
    Else
        MsgBox b1c1 & "_設計変更購入充当データを取り込みます。", vbInformation, "取込開始"
    End If

    Dim shisha As String
    Dim tankaFolder As String
    Dim isShinkansen As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo CleanupSekkei

    Dim sekkeiPath As String
    sekkeiPath = BuildFolderPath(nendo, SEKKEI_FOLDER, isShinkansen, shisha, tankaFolder)
    If sekkeiPath = "" Then
        MsgBox BuildMissingUnitPriceDataMessage(nendo, SEKKEI_DISPLAY_NAME), vbExclamation
        GoTo CleanupSekkei
    End If

    wsSekkei.Cells.Clear
    Dim sekkeiOk As Boolean
    sekkeiOk = ImportUnitPriceFiles(sekkeiPath, wsSekkei, "TableStyleMedium7")
    If Not sekkeiOk Then
        MsgBox BuildMissingUnitPriceDataMessage(nendo, SEKKEI_DISPLAY_NAME), vbExclamation
        GoTo CleanupSekkei
    End If

CleanupSekkei:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    If sekkeiOk Then
        Dim msgK As String
        msgK = BuildCompleteMessage(ws.Parent, False, True)
        If msgK <> "" Then MsgBox msgK, vbInformation, "完了"
    End If
End Sub

'===========================================================
' F1確定時のメイン処理
'  改修（#6）：マスタ Workbook を 1 度だけ開き、GetSenkuFromMaster と
'              ExtractProjectNamesFromMasterSheet を共用する。
'              （マスタは Read Only で 2 回開かれていた）
'===========================================================
Public Sub RunF1Process(ByVal ws As Worksheet)
    Dim officeName As String
    Dim nendo As Long
    officeName = GetBasicInfoOfficeName(ws)
    If officeName = "" Then Exit Sub

    nendo = GetBasicInfoNendo(ws)
    If nendo = 0 Then
        MsgBox "B4に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim shisha As String
    Dim tankaFolder As String
    Dim isShinkansen As Boolean
    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo Cleanup

    Dim hayakiPath As String
    hayakiPath = BuildFolderPath(nendo, HAYAKI_FOLDER, isShinkansen, shisha, tankaFolder)

    Dim wsT As Worksheet
    If Not GetOrErrorSheet(ws.Parent, TANKA_SHEET_NAME, wsT) Then GoTo Cleanup
    If hayakiPath <> "" Then
        If Not ImportUnitPriceFiles(hayakiPath, wsT, "TableStyleMedium2") Then
            MsgBox BuildMissingUnitPriceDataMessage(nendo, TANKA_DISPLAY_NAME), vbExclamation
        End If
    Else
        MsgBox BuildMissingUnitPriceDataMessage(nendo, TANKA_DISPLAY_NAME), vbExclamation
    End If

    Dim sekkeiPath As String
    sekkeiPath = BuildFolderPath(nendo, SEKKEI_FOLDER, isShinkansen, shisha, tankaFolder)

    Dim wsS As Worksheet
    If Not GetOrErrorSheet(ws.Parent, SEKKEI_TANKA_SHEET_NAME, wsS) Then GoTo Cleanup
    If sekkeiPath <> "" Then
        If Not ImportUnitPriceFiles(sekkeiPath, wsS, "TableStyleMedium7") Then
            MsgBox BuildMissingUnitPriceDataMessage(nendo, SEKKEI_DISPLAY_NAME), vbExclamation
        End If
    Else
        MsgBox BuildMissingUnitPriceDataMessage(nendo, SEKKEI_DISPLAY_NAME), vbExclamation
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
End Sub

'===========================================================
' 完了メッセージ文字列を組み立てて返す
'===========================================================
Public Function BuildCompleteMessage(ByVal wb As Workbook, _
                                      Optional ByVal includeTanka As Boolean = True, _
                                      Optional ByVal includeSekkei As Boolean = True) As String
    Dim msgTanka As String
    Dim msgSekkei As String
    If includeTanka Then msgTanka = GetSheetHeaderText(wb, TANKA_SHEET_NAME)
    If includeSekkei Then msgSekkei = GetSheetHeaderText(wb, SEKKEI_TANKA_SHEET_NAME)

    If msgTanka <> "" And msgSekkei <> "" Then
        BuildCompleteMessage = msgTanka & vbCrLf & _
                               msgSekkei & vbCrLf & _
                               "上記2つの購入充当データを適用しました。"
    ElseIf msgTanka <> "" Then
        BuildCompleteMessage = msgTanka & "の購入充当データを適用しました。"
    ElseIf msgSekkei <> "" Then
        BuildCompleteMessage = msgSekkei & "の購入充当データを適用しました。"
    End If
End Function

'===========================================================
' 指定シートのA1:G1の空でないセル値を連結して返す
'===========================================================
Private Function GetSheetHeaderText(ByVal wb As Workbook, _
                                     ByVal sheetName As String) As String
    GetSheetHeaderText = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim arr As Variant
    arr = ws.Range("A1:G1").value

    Dim c As Long
    Dim parts() As String
    ReDim parts(0)
    Dim cnt As Long
    cnt = 0

    For c = 1 To 7
        Dim v As String
        v = Trim(CStr(arr(1, c)))
        If v <> "" Then
            ReDim Preserve parts(cnt)
            parts(cnt) = v
            cnt = cnt + 1
        End If
    Next c

    If cnt > 0 Then
        GetSheetHeaderText = Join(parts, " ")
    End If
End Function

'===========================================================
' マスタファイルから出張所名に対応する線区情報を取得
'  改修（#6）：
'    - 検索を Application.Match ベースに切替。Match 失敗時のみ線形ループへ
'      フォールバックすることで、シートが大きくなっても即応する。
'    - エラー時にも masterWb を確実に閉じる。
'===========================================================
Private Function GetSenku(ByVal officeName As String, _
                           ByRef shisha As String, _
                           ByRef tankaFolder As String, _
                           ByRef isShinkansen As Boolean) As Boolean
    GetSenku = False
    shisha = ""
    tankaFolder = ""
    isShinkansen = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(GetMasterFilePath) Then
        MsgBox "マスタファイルが見つかりませんでした。" & vbCrLf & GetMasterFilePath, vbCritical
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing

    Dim masterWb As Workbook
    Dim masterWs As Worksheet
    On Error Resume Next
    Set masterWb = Workbooks.Open(GetMasterFilePath, ReadOnly:=True)
    On Error GoTo 0
    If masterWb Is Nothing Then
        MsgBox "マスタファイルを開けませんでした。", vbCritical
        Exit Function
    End If

    On Error GoTo Cleanup

    Set masterWs = masterWb.Sheets(MASTER_SHEET_NAME)
    If masterWs Is Nothing Then
        MsgBox "マスタファイルに「" & MASTER_SHEET_NAME & "」シートが見つかりません。", vbCritical
        GoTo Cleanup
    End If

    If GetSenkuFromMasterSheet(masterWs, officeName, shisha, tankaFolder) Then
        isShinkansen = (InStr(officeName, "新幹線") > 0)
        GetSenku = True
    Else
        MsgBox "出張所「" & officeName & "」がマスタに見つかりませんでした。" & vbCrLf & _
               "シート：" & MASTER_SHEET_NAME & " / C列を確認してください。", vbExclamation
    End If

    If tankaFolder = "" And GetSenku Then
        MsgBox "出張所「" & officeName & "」の単価適用保線区（F列）が空です。", vbExclamation
        GetSenku = False
    End If

Cleanup:
    On Error Resume Next
    If Not masterWb Is Nothing Then masterWb.Close False
    On Error GoTo 0
End Function

' すでに開いた masterWs を使って線区情報を抜き出す。
' Application.Match で C 列を直引きし、見つかった行から D 列・F 列を読む。
' 完全一致前提（元実装も Trim 後の文字列完全一致）。Match が失敗したら
' Trim 値が異なる可能性に備え、念のため線形フォールバックする。
Private Function GetSenkuFromMasterSheet(ByVal masterWs As Worksheet, _
                                          ByVal officeName As String, _
                                          ByRef shisha As String, _
                                          ByRef tankaFolder As String) As Boolean
    GetSenkuFromMasterSheet = False

    Dim lastRow As Long
    lastRow = masterWs.Cells(masterWs.rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim searchRange As Range
    Set searchRange = masterWs.Range("C2:C" & lastRow)

    ' まず Match で完全一致を試す
    Dim matchedRow As Variant
    On Error Resume Next
    matchedRow = Application.Match(officeName, searchRange, 0)
    On Error GoTo 0

    Dim foundRow As Long
    If IsNumeric(matchedRow) Then
        foundRow = CLng(matchedRow) + 1 ' C2 オフセット分加算
    Else
        ' Trim ずれを想定してフォールバック線形検索（C 列 D 列 F 列を配列読込）
        Dim arr As Variant
        arr = masterWs.Range("C2:F" & lastRow).value
        Dim i As Long
        For i = 1 To UBound(arr, 1)
            If Trim$(CStr(arr(i, 1))) = officeName Then
                foundRow = i + 1
                Exit For
            End If
        Next i
    End If

    If foundRow = 0 Then Exit Function

    shisha = Trim$(CStr(masterWs.Cells(foundRow, "D").value))
    tankaFolder = Trim$(CStr(masterWs.Cells(foundRow, "F").value))
    GetSenkuFromMasterSheet = True
End Function

'===========================================================
' フォルダパスを組み立てて返す
'===========================================================
Private Function BuildFolderPath(ByVal nendo As Long, _
                                  ByVal category As String, _
                                  ByVal isShinkansen As Boolean, _
                                  ByVal shisha As String, _
                                  ByVal tankaFolder As String) As String
    BuildFolderPath = ""

    Dim sep As String
    sep = Application.PathSeparator

    Dim userName As String
    userName = Environ("USERNAME")

    Dim basePath As String
    basePath = "C:" & sep & "Users" & sep & userName & sep & CommonCompanyNameText() & sep & _
               UnitPriceFolderText() & sep & UNIT_PRICE_TABLE & sep & _
               CStr(nendo) & sep & category & sep

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fullPath As String
    If isShinkansen Then
        fullPath = basePath & SHINKANSEN_FOLDER & sep & tankaFolder
        If fso.FolderExists(fullPath) Then BuildFolderPath = fullPath & sep
    Else
        BuildFolderPath = ResolveZairaisenUnitPriceFolder(fso, basePath, shisha, tankaFolder, sep)
    End If

    Set fso = Nothing
End Function

Private Function ResolveZairaisenUnitPriceFolder(ByVal fso As Object, _
                                                  ByVal basePath As String, _
                                                  ByVal shisha As String, _
                                                  ByVal tankaFolder As String, _
                                                  ByVal sep As String) As String
    ResolveZairaisenUnitPriceFolder = ""

    Dim zairaisenPath As String
    zairaisenPath = basePath & ZAIRAISEN_FOLDER & sep
    If Not fso.FolderExists(zairaisenPath) Then Exit Function

    Dim directPath As String
    directPath = zairaisenPath & shisha & sep & tankaFolder
    If fso.FolderExists(directPath) Then
        ResolveZairaisenUnitPriceFolder = directPath & sep
        Exit Function
    End If

    Dim parentFolder As Object
    Dim branchFolder As Object
    Dim targetPath As String
    Set parentFolder = fso.GetFolder(zairaisenPath)

    For Each branchFolder In parentFolder.SubFolders
        If IsMatchingBranchFolder(branchFolder.Name, shisha) Then
            targetPath = branchFolder.Path & sep & tankaFolder
            If fso.FolderExists(targetPath) Then
                ResolveZairaisenUnitPriceFolder = targetPath & sep
                Exit Function
            End If
        End If
    Next branchFolder

    For Each branchFolder In parentFolder.SubFolders
        targetPath = branchFolder.Path & sep & tankaFolder
        If fso.FolderExists(targetPath) Then
            ResolveZairaisenUnitPriceFolder = targetPath & sep
            Exit Function
        End If
    Next branchFolder
End Function

Private Function IsMatchingBranchFolder(ByVal folderName As String, _
                                         ByVal shisha As String) As Boolean
    IsMatchingBranchFolder = False
    If folderName = shisha Then
        IsMatchingBranchFolder = True
    ElseIf InStr(1, folderName, shisha, vbTextCompare) > 0 Then
        IsMatchingBranchFolder = True
    ElseIf InStr(1, shisha, folderName, vbTextCompare) > 0 Then
        IsMatchingBranchFolder = True
    End If
End Function

'===========================================================
' 対象シートを取得する
'===========================================================
Private Function GetOrErrorSheet(ByVal wb As Workbook, _
                                  ByVal sheetName As String, _
                                  ByRef ws As Worksheet) As Boolean
    GetOrErrorSheet = False
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "取込先シート「" & sheetName & "」が見つかりません。" & vbCrLf & _
               "シートは自動作成しません。", vbCritical
        Exit Function
    End If
    GetOrErrorSheet = True
End Function

'===========================================================
' 指定フォルダのExcelファイルを取込み、対象シートに貼り付け
'===========================================================
Private Function ImportUnitPriceFiles(ByVal folderPath As String, _
                                       ByVal wsDest As Worksheet, _
                                       Optional ByVal tableStyleName As String = "TableStyleMedium2") As Boolean
    ImportUnitPriceFiles = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Set fso = Nothing
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    Dim normalFile As String
    Dim sub2File As String
    normalFile = ""
    sub2File = ""
    Dim f As Object
    Dim ext As String
    For Each f In folder.Files
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
            If InStr(f.Name, "-2") > 0 Then
                sub2File = f.Path
            Else
                normalFile = f.Path
            End If
        End If
    Next f
    Set fso = Nothing

    If normalFile = "" Then Exit Function

    Dim tableStyle As String
    Dim tableHeaderRow As Long
    tableStyle = tableStyleName
    tableHeaderRow = 5

    wsDest.Cells.Clear

    Dim normalLastRow As Long
    normalLastRow = 0
    Call PasteFromFileFast(normalFile, wsDest, 1, normalLastRow)

    Dim totalLastRow As Long
    totalLastRow = normalLastRow
    If sub2File <> "" And normalLastRow > 0 Then
        Call PasteFromFileFast(sub2File, wsDest, normalLastRow + 1, totalLastRow)
    End If

    If tableStyle <> "" And totalLastRow >= tableHeaderRow + 1 Then
        Call ResetAsTable(wsDest, tableHeaderRow, totalLastRow, tableStyle)
    End If

    wsDest.PageSetup.PrintArea = ""
    wsDest.Parent.Windows(1).View = xlNormalView

    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(totalLastRow, 7))
        .Font.Name = "BIZ UDゴシック"
        .Font.Size = 11
    End With

    With wsDest.Range("A1:G1")
        .Merge
        .Font.Name = "BIZ UDゴシック"
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsDest.Range("A5:G5")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Dim centerCols As Variant
    Dim cc As Variant
    For Each cc In Array(1, 5, 7)
        With wsDest.Range(wsDest.Cells(6, cc), wsDest.Cells(totalLastRow, cc))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next cc

    ImportUnitPriceFiles = True
End Function

'===========================================================
' 取込元ファイルからテーブル情報を取得
'===========================================================
Private Sub GetTableInfo(ByVal filePath As String, _
                          ByRef tableStyle As String, _
                          ByRef tableHeaderRow As Long)
    tableStyle = ""
    tableHeaderRow = 1

    Dim srcWb As Workbook
    On Error Resume Next
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If srcWb Is Nothing Then Exit Sub

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = srcWb.Sheets("単価")
    On Error GoTo 0
    If Not srcWs Is Nothing Then
        If srcWs.ListObjects.Count > 0 Then
            tableStyle = srcWs.ListObjects(1).tableStyle
            tableHeaderRow = srcWs.ListObjects(1).HeaderRowRange.Row
        End If
    End If

    srcWb.Close False
End Sub

'===========================================================
' ファイルを開いて高速転記
'  改修（#4）：destStartRow=1 のときと >1 のときで読み込み範囲を分岐し、
'              中間配列詰め替えループを廃止。読み込んだ配列をそのまま貼付。
'===========================================================
Private Sub PasteFromFileFast(ByVal filePath As String, _
                               ByVal wsDest As Worksheet, _
                               ByVal destStartRow As Long, _
                               ByRef destLastRow As Long)
    destLastRow = 0

    Const SNAP_COLS As Long = 7

    Dim srcWb As Workbook
    On Error Resume Next
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If srcWb Is Nothing Then
        MsgBox "ファイルを開けませんでした。" & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = srcWb.Sheets("単価")
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "取込ファイルに「単価」シートが見つかりませんでした。" & vbCrLf & filePath, vbExclamation
        srcWb.Close False
        Exit Sub
    End If

    Dim snapRows As Long
    snapRows = srcWs.Cells(srcWs.rows.Count, "A").End(xlUp).Row
    If snapRows < 1 Then
        srcWb.Close False
        Exit Sub
    End If

    ' 改修（#4）：必要範囲だけを直接読み込む。
    ' destStartRow=1 のときは先頭のヘッダー（行 1～4）も含めて読み込み。
    ' destStartRow>1 のときは 6 行目以降のデータ部分のみ読み込み。
    Dim srcStartRow As Long
    If destStartRow = 1 Then
        srcStartRow = 1
    Else
        srcStartRow = 6
    End If

    If snapRows < srcStartRow Then
        srcWb.Close False
        Exit Sub
    End If

    Dim dataArr As Variant
    dataArr = srcWs.Range(srcWs.Cells(srcStartRow, 1), srcWs.Cells(snapRows, SNAP_COLS)).value

    ' 1 セルしか読まなかった場合は配列化されないため Variant 配列に揃える
    If Not IsArray(dataArr) Then
        Dim singleVal As Variant
        singleVal = dataArr
        ReDim dataArr(1 To 1, 1 To SNAP_COLS)
        dataArr(1, 1) = singleVal
    End If

    Dim colWidths(1 To SNAP_COLS) As Double
    If destStartRow = 1 Then
        Dim c As Long
        For c = 1 To SNAP_COLS
            colWidths(c) = srcWs.Columns(c).ColumnWidth
        Next c
    End If

    srcWb.Close False
    Set srcWs = Nothing
    Set srcWb = Nothing

    Dim rowCount As Long
    rowCount = UBound(dataArr, 1)

    If destStartRow = 1 Then
        For c = 1 To SNAP_COLS
            wsDest.Columns(c).ColumnWidth = colWidths(c)
        Next c
    End If

    wsDest.Range(wsDest.Cells(destStartRow, 1), _
                 wsDest.Cells(destStartRow + rowCount - 1, SNAP_COLS)).value = dataArr

    wsDest.Range(wsDest.Cells(destStartRow, 1), _
                 wsDest.Cells(destStartRow + rowCount - 1, 1)).NumberFormat = "0"

    wsDest.Range(wsDest.Cells(destStartRow, 6), _
                 wsDest.Cells(destStartRow + rowCount - 1, 6)).NumberFormat = "#,##0"

    destLastRow = destStartRow + rowCount - 1
End Sub

'===========================================================
' 指定範囲をテーブルとして再設定
'===========================================================
Private Sub ResetAsTable(ByVal ws As Worksheet, _
                          ByVal headerRow As Long, _
                          ByVal lastRow As Long, _
                          ByVal tableStyle As String)
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    Dim newTable As ListObject
    Set newTable = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    On Error Resume Next
    newTable.tableStyle = tableStyle
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'  本モジュール専用の日本語パス・名称（ChrW 構築でキャッシュ）
'    線路出張所用_注文書_請求書アクセスサイト - ドキュメント
'    単価マスタ
'    参照データ
'    出張所別_単価適用線区.xlsx
'    帳票7_支払金額計算シート - 帳票7_支払金額計算シート
'--------------------------------------------------------------------------

Private Function OrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H7528) & _
                 "_" & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & "_" & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & _
                 ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & ChrW$(&H30B9) & _
                 ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & _
                 " - " & ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    OrderInvoiceDocumentFolderText = cached
End Function

Private Function UnitPriceMasterFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5358) & ChrW$(&H4FA1) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) ' 単価マスタ
    End If
    UnitPriceMasterFolderText = cached
End Function

Private Function UnitPriceMasterReferenceFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H53C2) & ChrW$(&H7167) & ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF) ' 参照データ
    End If
    UnitPriceMasterReferenceFolderText = cached
End Function

Private Function UnitPriceMasterFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H5225) & "_" & _
                 ChrW$(&H5358) & ChrW$(&H4FA1) & ChrW$(&H9069) & ChrW$(&H7528) & _
                 ChrW$(&H7DDA) & ChrW$(&H533A) & ".xlsx" ' 出張所別_単価適用線区.xlsx
    End If
    UnitPriceMasterFileNameText = cached
End Function

Private Function UnitPriceFolderText() As String
    Static cached As String
    If cached = "" Then
        ' 帳票7_支払金額計算シート - 帳票7_支払金額計算シート
        Dim part As String
        part = ChrW$(&H5E33) & ChrW$(&H7968) & "7_" & _
               ChrW$(&H652F) & ChrW$(&H6255) & ChrW$(&H91D1) & ChrW$(&H984D) & _
               ChrW$(&H8A08) & ChrW$(&H7B97) & ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8)
        cached = part & " - " & part
    End If
    UnitPriceFolderText = cached
End Function
