Option Explicit

'==========================================================================
'  施工指示書・施工通知書 取込みモジュール  Construction_Order_Import
'  改修履歴: CHANGELOG.md 参照
'
'  概要:
'   ・ルートフォルダから施工指示書／施工通知書ブックを選択して取り込む
'   ・取込ブック A3 の値をシート名にして ThisWorkbook へ工事側シートを新規作成
'   ・工種分類が「購入充当」の行は別シート(購入充当指示／購入充当通知)へ分離
'   ・施工指示書は基本情報 B6(支店)/C6(出張所) に対応する管理室で行を絞り込む
'   ・取込後にソート(工事側／購入充当側でルールが異なる)と列幅自動調整を行う
'
'  出力ヘッダー(13列, Word 仕様画像どおり):
'   A 施工業者 / B 整理番号 / C 工事種類 / D 昼夜別 / E 単位 / F 数量 /
'   G 契約線区名 / H 管理室 / I JR単価 / J JR金額 / K 外注単価 /
'   L 外注金額 / M 工種分類
'   ※ 取込元に対応データが無い列(施工業者・外注単価・外注金額、通知書の管理室)は
'      空白のまま出力する(後工程で使用)。取込元の単価→JR単価、金額→JR金額。
'
'  エントリポイント: ImportConstructionDocument
'==========================================================================

' 取込種別
Private Const DOC_ORDER As Long = 1    ' 施工指示書
Private Const DOC_NOTICE As Long = 2   ' 施工通知書

' データ開始行
Private Const DATA_START_ROW As Long = 22

' 出力列インデックス(A=1)
Private Const COL_VENDOR As Long = 1      ' 施工業者(空白)
Private Const COL_SEIRI As Long = 2       ' 整理番号
Private Const COL_TYPE As Long = 3        ' 工事種類
Private Const COL_DAYNIGHT As Long = 4    ' 昼夜別
Private Const COL_UNIT As Long = 5        ' 単位
Private Const COL_QTY As Long = 6         ' 数量
Private Const COL_LINE As Long = 7        ' 契約線区名
Private Const COL_MGR As Long = 8         ' 管理室
Private Const COL_JR_PRICE As Long = 9    ' JR単価
Private Const COL_JR_AMOUNT As Long = 10  ' JR金額
Private Const COL_OUT_PRICE As Long = 11  ' 外注単価(空白)
Private Const COL_OUT_AMOUNT As Long = 12 ' 外注金額(空白)
Private Const COL_KIND As Long = 13       ' 工種分類
Private Const COL_AUTO_PRICE As Long = 15 ' 参照単価(O列)
Private Const COL_AUTO_AMOUNT As Long = 16 ' 参照単価金額(P列)
Private Const COL_PRICE_COMPARE As Long = 17 ' 単価比較(Q列)
Private Const COL_FLAG_SIDE As Long = 27  ' 補助:側線フラグ(AA列)
Private Const COL_FLAG_WELD As Long = 28  ' 補助:レール溶接フラグ(AB列)
Private Const OUTPUT_COL_COUNT As Long = 13

' 管理室マスタ(出張所別_単価適用線区.xlsx)
Private Const MGR_MASTER_SHEET As String = "JR管理室対応出張所"
Private Const MGR_MASTER_BRANCH_COL As Long = 2  ' B列 支店
Private Const MGR_MASTER_OFFICE_COL As Long = 3  ' C列 出張所
Private Const MGR_MASTER_ROOM_COL As Long = 6    ' F列 管理室名
Private Const MGR_MASTER_START_ROW As Long = 2

' 工種分類キーワード
Private Const PURCHASE_KEYWORD As String = "購入充当"
Private Const SIDELINE_KEYWORD As String = "側線"
Private Const WELDING_KEYWORD As String = "レール溶接"

' 基本情報セル
Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_PUBLIC_CELL As String = "B4"
Private Const BASIC_INFO_AMOUNT_CELL As String = "C22"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"

' 工事件名別マスタ(F列=積算線区、G列=施工指示書記載線区名)
Private Const PROJECT_MASTER_START_ROW As Long = 2
Private Const PROJECT_MASTER_UNIT_PRICE_LINE_COL As Long = 6
Private Const PROJECT_MASTER_SOURCE_LINE_COL As Long = 7
Private Const PROJECT_MASTER_FOLDER As String = "工事件名別マスタ"
Private Const UNIT_PRICE_DATA_START_ROW As Long = 7

' 取込ブックのシート名取得元セル
Private Const SOURCE_SHEET_NAME_CELL As String = "A3"

'==========================================================================
'  エントリポイント
'==========================================================================
Public Sub ImportConstructionDocument()
    Dim srcWb As Workbook, masterWb As Workbook
    Dim srcOpenedHere As Boolean, masterOpenedHere As Boolean
    Dim scrn As Boolean, calc As XlCalculation, evt As Boolean, alerts As Boolean

    ' 元のアプリ状態を保存
    scrn = Application.screenUpdating
    calc = Application.Calculation
    evt = Application.EnableEvents
    alerts = Application.DisplayAlerts

    On Error GoTo Cleanup

    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '--- 取込ファイルの選択 -------------------------------------------------
    Dim srcPath As String
    srcPath = PickSourceFile()
    If srcPath = "" Then GoTo Cleanup           ' キャンセル
    LogCI "srcPath=[" & srcPath & "]"

    '--- 取込種別の判定 -----------------------------------------------------
    Dim docType As Long
    docType = DetermineDocType(srcPath)
    If docType = 0 Then GoTo Cleanup            ' キャンセル
    LogCI "docType=" & docType

    '--- 取込ブックを開く ---------------------------------------------------
    Set srcWb = OpenWorkbookReadOnly(srcPath, srcOpenedHere)
    If srcWb Is Nothing Then
        MsgBox "取込対象ブックを開けませんでした。" & vbCrLf & srcPath, vbExclamation
        GoTo Cleanup
    End If

    Dim srcWs As Worksheet
    Set srcWs = srcWb.ActiveSheet               ' ※ 取込ブックのアクティブシートを対象とする

    '--- 新規シート名(A3) ---------------------------------------------------
    Dim baseSheetName As String
    baseSheetName = SanitizeSheetName(CommonNzText(srcWs.Range(SOURCE_SHEET_NAME_CELL).value))
    If baseSheetName = "" Then
        MsgBox "取込ブックの " & SOURCE_SHEET_NAME_CELL & " が空のため、シート名を決定できません。", vbExclamation
        GoTo Cleanup
    End If
    LogCI "baseSheetName=[" & baseSheetName & "]"

    '--- データ最終行(取込元の整理番号=A列) ---------------------------------
    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.rows.Count, "A").End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "取込対象データ(" & DATA_START_ROW & "行目以降)が見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    '--- 取込列マップ -------------------------------------------------------
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

    '--- 管理室フィルタ集合(施工指示書のみ) ---------------------------------
    Dim mgrSet As Object   ' Scripting.Dictionary
    If docType = DOC_ORDER Then
        Set mgrSet = BuildManagerRoomSet(masterWb, masterOpenedHere)
        If mgrSet Is Nothing Then GoTo Cleanup
        If mgrSet.Count = 0 Then
            MsgBox "基本情報 " & BASIC_INFO_BRANCH_CELL & "/" & BASIC_INFO_OFFICE_CELL & _
                   " に対応する管理室が管理室マスタに見つかりませんでした。" & vbCrLf & _
                   "支店名・出張所名をご確認ください。", vbExclamation
            GoTo Cleanup
        End If
    End If

    '--- 取込元の列ベクトルを取得 -------------------------------------------
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

    '--- 行ごとに分類(工事側 / 購入充当側) -----------------------------------
    Dim worksRows As Collection, purchRows As Collection
    Set worksRows = New Collection
    Set purchRows = New Collection

    ' 診断用: 取込データに含まれる管理室の一覧(施工指示書のみ)
    Dim seenMgr As Object
    Set seenMgr = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim seiriText As String, kindText As String, mgrOut As String, mraw As String
    For i = 1 To n
        seiriText = Trim$(CommonNzText(vSeiri(i)))
        If seiriText <> "" Then                          ' 空行スキップ(整理番号が空)
            Dim keepRow As Boolean
            keepRow = True

            If docType = DOC_ORDER Then
                mgrOut = CommonNzText(vMgr(i))
                mraw = Trim$(mgrOut)
                If mraw <> "" Then
                    If Not seenMgr.Exists(mraw) Then seenMgr.Add mraw, True
                End If
                ' 管理室フィルタ(購入充当側も同様に適用)
                If Not mgrSet.Exists(CommonRemoveAllSpaces(mgrOut)) Then keepRow = False
            Else
                mgrOut = ""                              ' 通知書は管理室なし(空白)
            End If

            If keepRow Then
                Dim rowArr(0 To 12) As Variant
                rowArr(COL_VENDOR - 1) = ""              ' 施工業者(空白)
                rowArr(COL_SEIRI - 1) = vSeiri(i)        ' 整理番号
                rowArr(COL_TYPE - 1) = vType(i)          ' 工事種類
                rowArr(COL_DAYNIGHT - 1) = vDN(i)        ' 昼夜別
                rowArr(COL_UNIT - 1) = vUnit(i)          ' 単位
                rowArr(COL_QTY - 1) = vQty(i)            ' 数量
                rowArr(COL_LINE - 1) = vLine(i)          ' 契約線区名
                rowArr(COL_MGR - 1) = mgrOut             ' 管理室
                rowArr(COL_JR_PRICE - 1) = vPrice(i)     ' JR単価(取込元の単価)
                rowArr(COL_JR_AMOUNT - 1) = vAmount(i)   ' JR金額(取込元の金額)
                rowArr(COL_OUT_PRICE - 1) = ""           ' 外注単価(空白)
                rowArr(COL_OUT_AMOUNT - 1) = ""          ' 外注金額(空白)
                rowArr(COL_KIND - 1) = vKind(i)          ' 工種分類

                kindText = CommonRemoveAllSpaces(CommonNzText(vKind(i)))
                If InStr(1, kindText, PURCHASE_KEYWORD) > 0 Then
                    purchRows.Add rowArr
                Else
                    worksRows.Add rowArr
                End If
            End If
        End If
    Next i
    LogCI "工事側=" & worksRows.Count & " / 購入充当側=" & purchRows.Count

    '--- 管理室フィルタで全件除外された場合の診断表示 -----------------------
    Dim filteredOutAll As Boolean
    filteredOutAll = (docType = DOC_ORDER) And (worksRows.Count = 0) And (purchRows.Count = 0)
    If filteredOutAll Then
        MsgBox "管理室フィルタで全ての行が除外されました(取込0件)。" & vbCrLf & vbCrLf & _
               "■基本情報(支店/出張所)に対応する管理室:" & vbCrLf & JoinKeys(mgrSet) & vbCrLf & vbCrLf & _
               "■取込データに含まれる管理室:" & vbCrLf & JoinKeys(seenMgr) & vbCrLf & vbCrLf & _
               "両者が違う地域の場合は、基本情報シートの支店・出張所と取込データの地域が" & _
               "一致しているかご確認ください。" & vbCrLf & _
               "「取込データに含まれる管理室」が(なし)の場合は、取込元の管理室セル位置(BP列)を" & _
               "ご確認ください。", vbExclamation, "取込み診断"
    End If

    '--- 取込ブック・マスタを閉じる(以降は出力処理のみ) ---------------------
    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Set srcWb = Nothing: srcOpenedHere = False
    If masterOpenedHere And Not masterWb Is Nothing Then masterWb.Close SaveChanges:=False
    Set masterWb = Nothing: masterOpenedHere = False

    '--- 工事側シート作成・書込み・ソート -----------------------------------
    Dim wsWorks As Worksheet
    Set wsWorks = CreateOrReplaceSheet(baseSheetName)
    If wsWorks Is Nothing Then GoTo Cleanup
    wsWorks.Tab.Color = RGB(255, 255, 0)
    WriteRecordsToSheet wsWorks, worksRows
    SortWorksSheet wsWorks
    WriteAdditionalHeaders wsWorks
    FillReferenceUnitPrices wsWorks
    FormatSheet wsWorks

    '--- 購入充当側シート作成・書込み・ソート(該当行がある場合のみ) ---------
    If purchRows.Count > 0 Then
        Dim purchName As String
        If docType = DOC_ORDER Then
            purchName = "購入充当指示"
        Else
            purchName = "購入充当通知"
        End If
        Dim wsPurch As Worksheet
        Set wsPurch = CreateOrReplaceSheet(purchName)
        If Not wsPurch Is Nothing Then
            If docType = DOC_NOTICE Then wsPurch.Tab.Color = RGB(255, 255, 0)
            WriteRecordsToSheet wsPurch, purchRows
            SortPurchaseSheet wsPurch
            If docType = DOC_NOTICE Then WriteAdditionalHeaders wsPurch, False
            FormatSheet wsPurch
        End If
    End If

    wsWorks.Activate
    wsWorks.Range("A1").Select

    Application.screenUpdating = scrn
    Application.Calculation = calc
    Application.EnableEvents = evt
    Application.DisplayAlerts = alerts

    If Not filteredOutAll Then
        MsgBox "取込みが完了しました。" & vbCrLf & _
               "工事側: " & worksRows.Count & " 件 (" & baseSheetName & ")" & vbCrLf & _
               "購入充当側: " & purchRows.Count & " 件", vbInformation
    End If
    Exit Sub

Cleanup:
    Dim errNo As Long, errDesc As String
    errNo = Err.Number: errDesc = Err.Description

    On Error Resume Next
    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If masterOpenedHere And Not masterWb Is Nothing Then masterWb.Close SaveChanges:=False
    On Error GoTo 0

    Application.screenUpdating = scrn
    Application.Calculation = calc
    Application.EnableEvents = evt
    Application.DisplayAlerts = alerts

    If errNo <> 0 Then
        MsgBox "取込み処理でエラーが発生しました。" & vbCrLf & _
               "Err " & errNo & ": " & errDesc, vbExclamation
    End If
End Sub

'==========================================================================
'  取込ファイル選択
'==========================================================================
Private Function PickSourceFile() As String
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

'==========================================================================
'  取込ルートフォルダの解決
'==========================================================================
Private Function GetImportRootFolder() As String
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
    ' 仕様書記載の絶対パス(フォールバック)
    candidates.Add "C:\Users\n-morimoto\" & CommonCompanyNameText() & "\" & docFolder & "\" & subPath

    Dim p As Variant
    For Each p In candidates
        If fso.FolderExists(CStr(p)) Then
            GetImportRootFolder = CStr(p)
            Exit Function
        End If
    Next p

    ' 見つからなければ先頭候補を返す(ダイアログ初期位置用)
    If candidates.Count > 0 Then GetImportRootFolder = CStr(candidates(1))
End Function

'==========================================================================
'  取込種別の判定(ファイル名 -> 不明ならユーザーに確認)
'==========================================================================
Private Function DetermineDocType(ByVal filePath As String) As Long
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

'==========================================================================
'  管理室集合の構築(基本情報 B6/C6 に対応する F列管理室名)
'==========================================================================
Private Function BuildManagerRoomSet(ByRef masterWb As Workbook, _
                                     ByRef masterOpenedHere As Boolean) As Object
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim branchName As String, officeName As String
    branchName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    officeName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    If branchName = "" Or officeName = "" Then
        MsgBox "基本情報シート " & BASIC_INFO_BRANCH_CELL & "(支店) または " & _
               BASIC_INFO_OFFICE_CELL & "(出張所) が空です。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim masterPath As String
    masterPath = ResolveMasterFilePath()
    If masterPath = "" Then
        MsgBox "出張所別_単価適用線区.xlsx が見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Set masterWb = OpenWorkbookReadOnly(masterPath, masterOpenedHere)
    If masterWb Is Nothing Then
        MsgBox "出張所別_単価適用線区.xlsx を開けませんでした。" & vbCrLf & masterPath, vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim wsMaster As Worksheet
    On Error Resume Next
    Set wsMaster = masterWb.worksheets(MGR_MASTER_SHEET)
    On Error GoTo 0
    If wsMaster Is Nothing Then
        MsgBox "管理室マスタに「" & MGR_MASTER_SHEET & "」シートが見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim mLast As Long, r As Long
    mLast = wsMaster.Cells(wsMaster.rows.Count, MGR_MASTER_BRANCH_COL).End(xlUp).Row

    Dim b As String, c As String, room As String
    For r = MGR_MASTER_START_ROW To mLast
        b = CommonRemoveAllSpaces(CommonNzText(wsMaster.Cells(r, MGR_MASTER_BRANCH_COL).value))
        c = CommonRemoveAllSpaces(CommonNzText(wsMaster.Cells(r, MGR_MASTER_OFFICE_COL).value))
        If b = branchName And c = officeName Then
            room = CommonRemoveAllSpaces(CommonNzText(wsMaster.Cells(r, MGR_MASTER_ROOM_COL).value))
            If room <> "" Then
                If Not dict.Exists(room) Then dict.Add room, True
            End If
        End If
    Next r

    LogCI "管理室集合 件数=" & dict.Count & " (支店=" & branchName & " 出張所=" & officeName & ")"
    Set BuildManagerRoomSet = dict
End Function

'==========================================================================
'  管理室マスタのパス解決(既存 mod_MaterialPriceImport を優先利用)
'==========================================================================
Private Function ResolveMasterFilePath() As String
    Dim p As String
    On Error Resume Next
    p = mod_MaterialPriceImport.GetMasterFilePath()
    On Error GoTo 0
    If p <> "" Then
        ResolveMasterFilePath = p
        Exit Function
    End If

    ' フォールバック: 環境変数から組み立て
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

'==========================================================================
'  ワークブックを読み取り専用で開く(既に開いていれば再利用)
'==========================================================================
Private Function OpenWorkbookReadOnly(ByVal filePath As String, _
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

'==========================================================================
'  列ベクトルの取得(1始まり 1次元配列で返す)
'==========================================================================
Private Function ReadColumnValues(ByVal ws As Worksheet, ByVal col As String, _
                                  ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim out() As Variant, v As Variant, i As Long, cnt As Long, colIdx As Long
    cnt = r2 - r1 + 1
    ReDim out(1 To cnt)

    ' 列をインデックス化し、シートの列数範囲内かを確認する。
    ' (.xls は 256 列まで。範囲外の列は参照すると Err1004 になるため空白扱い)
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

'==========================================================================
'  列文字(A,B,...,AA,...)を列番号(1始まり)に変換する。
'  ワークシートに依存しない純粋計算のため、範囲外列でも安全に判定できる。
'==========================================================================
Private Function ColLetterToNum(ByVal col As String) As Long
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

'==========================================================================
'  シート作成(同名が存在すれば確認の上で置換)
'==========================================================================
Private Function CreateOrReplaceSheet(ByVal sheetName As String) As Worksheet
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = ThisWorkbook.worksheets(sheetName)
    On Error GoTo 0

    If Not existing Is Nothing Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("シート「" & sheetName & "」は既に存在します。" & vbCrLf & _
                     "削除して作り直しますか？", vbYesNo + vbQuestion, "シートの上書き確認")
        If ans <> vbYes Then
            Set CreateOrReplaceSheet = Nothing
            Exit Function
        End If
        existing.Delete   ' DisplayAlerts は呼び出し元で False 設定済み
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.worksheets.Add(After:=ThisWorkbook.worksheets(ThisWorkbook.worksheets.Count))
    ws.Name = sheetName
    Set CreateOrReplaceSheet = ws
End Function

'==========================================================================
'  レコード書込み(ヘッダー + データ)
'==========================================================================
Private Sub WriteRecordsToSheet(ByVal ws As Worksheet, ByVal rows As Collection)
    ' ヘッダー(13列, Word 仕様画像どおり)
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

'==========================================================================
'  追加ヘッダー(O列・P列・Q列)
'==========================================================================
Private Sub WriteAdditionalHeaders(ByVal ws As Worksheet, _
                                   Optional ByVal includePriceComparison As Boolean = True)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つからないため、追加ヘッダーを設定できませんでした。", vbExclamation
        Exit Sub
    End If

    Dim amountName As String
    amountName = CommonNzText(wsInfo.Range(BASIC_INFO_AMOUNT_CELL).value)
    ws.Range("O1").value = CommonNzText(wsInfo.Range(BASIC_INFO_PUBLIC_CELL).value) & "公開" & amountName
    ws.Range("P1").value = amountName & "金額"
    If includePriceComparison Then ws.Range("Q1").value = "単価比較"
End Sub

'==========================================================================
'  出力表の最終列(単価比較がある場合はQ列)
'==========================================================================
Private Function GetOutputLastColumn(ByVal ws As Worksheet) As Long
    If CommonNzText(ws.Range("Q1").value) <> "" Then
        GetOutputLastColumn = ws.Range("Q1").Column
    ElseIf CommonNzText(ws.Range("P1").value) <> "" Then
        GetOutputLastColumn = ws.Range("P1").Column
    Else
        GetOutputLastColumn = COL_KIND
    End If
End Function

'==========================================================================
'  工事側シートの参照単価・金額・比較結果を設定
'==========================================================================
Private Sub FillReferenceUnitPrices(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim lineSheetMap As Object
    Set lineSheetMap = BuildConstructionLineSheetMap()

    Dim sheetPriceCaches As Object
    Set sheetPriceCaches = CreateObject("Scripting.Dictionary")
    sheetPriceCaches.CompareMode = vbTextCompare

    Dim matchedCount As Long, unresolvedLineCount As Long, missingRecordCount As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim unitPriceSheetName As String
        unitPriceSheetName = ResolveUnitPriceSheetName(lineSheetMap, CommonNzText(ws.Cells(r, COL_LINE).value))

        Dim referencePrice As Variant
        referencePrice = Empty

        If unitPriceSheetName = "" Then
            unresolvedLineCount = unresolvedLineCount + 1
        Else
            Dim priceRows As Object
            Set priceRows = GetUnitPriceRows(unitPriceSheetName, sheetPriceCaches)

            Dim recordKey As String
            recordKey = NormalizeRecordKey(ws.Cells(r, COL_SEIRI).value)

            If priceRows Is Nothing Or recordKey = "" Then
                missingRecordCount = missingRecordCount + 1
            ElseIf priceRows.Exists(recordKey) Then
                Dim dayNightPrices As Variant
                dayNightPrices = priceRows(recordKey)
                referencePrice = SelectDayNightPrice(CommonNzText(ws.Cells(r, COL_DAYNIGHT).value), dayNightPrices)
                If Not IsEmpty(referencePrice) Then matchedCount = matchedCount + 1
            Else
                missingRecordCount = missingRecordCount + 1
            End If
        End If

        If Not IsEmpty(referencePrice) Then ws.Cells(r, COL_AUTO_PRICE).value = referencePrice
        WritePriceComparison ws, r
    Next r

    ws.Range(ws.Cells(2, COL_AUTO_AMOUNT), ws.Cells(lastRow, COL_AUTO_AMOUNT)).FormulaR1C1 = _
        "=IF(OR(RC[-1]="""",RC[-10]=""""),"""",RC[-1]*RC[-10])"

    With ws.Range(ws.Cells(2, COL_AUTO_PRICE), ws.Cells(lastRow, COL_AUTO_AMOUNT))
        .NumberFormatLocal = ChrW$(&HA5) & "#,##0;[赤]-" & ChrW$(&HA5) & "#,##0"
    End With
    ws.Range(ws.Cells(1, COL_PRICE_COMPARE), ws.Cells(lastRow, COL_PRICE_COMPARE)).HorizontalAlignment = xlCenter

    LogCI "参照単価一致=" & matchedCount & _
          " / 線区未解決=" & unresolvedLineCount & _
          " / 整理番号未一致=" & missingRecordCount
End Sub

'==========================================================================
'  工事件名別マスタのF/G列から、施工指示書線区名→単価シート名を構築
'==========================================================================
Private Function BuildConstructionLineSheetMap() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    ' マスタが取得できない場合も、施工指示書線区名と単価シート名が同じなら参照できる。
    Dim targetSheet As Worksheet
    For Each targetSheet In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(targetSheet) Then
            AddLineSheetAliases result, targetSheet.Name, targetSheet.Name
        End If
    Next targetSheet

    Dim masterPath As String
    masterPath = ResolveProjectLineMasterPath()
    If masterPath = "" Then
        LogCI "工事件名別マスタ未検出 -> 単価シート名の直接照合のみ"
        Set BuildConstructionLineSheetMap = result
        Exit Function
    End If

    Dim masterBook As Workbook, openedHere As Boolean
    Set masterBook = OpenWorkbookReadOnly(masterPath, openedHere)
    If masterBook Is Nothing Then
        LogCI "工事件名別マスタを開けない path=[" & masterPath & "]"
        Set BuildConstructionLineSheetMap = result
        Exit Function
    End If

    On Error GoTo Cleanup

    Dim masterSheet As Worksheet
    For Each masterSheet In masterBook.worksheets
        Dim lastRow As Long
        lastRow = masterSheet.Cells(masterSheet.rows.Count, PROJECT_MASTER_UNIT_PRICE_LINE_COL).End(xlUp).Row

        Dim r As Long
        For r = PROJECT_MASTER_START_ROW To lastRow
            Dim unitPriceLineName As String
            Dim sourceLineName As String
            Dim actualSheetName As String

            unitPriceLineName = CommonNzText(masterSheet.Cells(r, PROJECT_MASTER_UNIT_PRICE_LINE_COL).value)
            If Trim$(unitPriceLineName) <> "" Then
                actualSheetName = FindImportedUnitPriceSheetName(unitPriceLineName)
                If actualSheetName <> "" Then
                    sourceLineName = CommonNzText(masterSheet.Cells(r, PROJECT_MASTER_SOURCE_LINE_COL).value)
                    If Trim$(sourceLineName) = "" Then sourceLineName = unitPriceLineName

                    AddLineSheetAliases result, sourceLineName, actualSheetName
                    AddLineSheetAliases result, unitPriceLineName, actualSheetName
                End If
            End If
        Next r
    Next masterSheet

Cleanup:
    If Err.Number <> 0 Then LogCI "線区名マスタ読込エラー Err " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If openedHere And Not masterBook Is Nothing Then masterBook.Close SaveChanges:=False
    On Error GoTo 0

    LogCI "線区名→単価シート対応数=" & result.Count & " master=[" & masterPath & "]"
    Set BuildConstructionLineSheetMap = result
End Function

Private Sub AddLineSheetAliases(ByVal lineSheetMap As Object, _
                                ByVal sourceLineName As String, _
                                ByVal unitPriceSheetName As String)
    AddLineSheetAlias lineSheetMap, "E|" & NormalizeLineLookupText(sourceLineName, False), unitPriceSheetName
    AddLineSheetAlias lineSheetMap, "S|" & NormalizeLineLookupText(sourceLineName, True), unitPriceSheetName
End Sub

Private Sub AddLineSheetAlias(ByVal lineSheetMap As Object, _
                              ByVal key As String, _
                              ByVal unitPriceSheetName As String)
    If lineSheetMap Is Nothing Or Len(key) <= 2 Then Exit Sub
    If Not lineSheetMap.Exists(key) Then
        lineSheetMap.Add key, unitPriceSheetName
    ElseIf StrComp(CStr(lineSheetMap(key)), unitPriceSheetName, vbTextCompare) <> 0 Then
        LogCI "線区名対応が重複 key=[" & key & "] first=[" & CStr(lineSheetMap(key)) & _
              "] ignored=[" & unitPriceSheetName & "]"
    End If
End Sub

Private Function ResolveUnitPriceSheetName(ByVal lineSheetMap As Object, _
                                           ByVal importedLineName As String) As String
    If lineSheetMap Is Nothing Then Exit Function

    Dim key As String
    key = "E|" & NormalizeLineLookupText(importedLineName, False)
    If Len(key) > 2 And lineSheetMap.Exists(key) Then
        ResolveUnitPriceSheetName = CStr(lineSheetMap(key))
        Exit Function
    End If

    key = "S|" & NormalizeLineLookupText(importedLineName, True)
    If Len(key) > 2 And lineSheetMap.Exists(key) Then
        ResolveUnitPriceSheetName = CStr(lineSheetMap(key))
    End If
End Function

Private Function NormalizeLineLookupText(ByVal sourceText As String, _
                                         ByVal removeParenthetical As Boolean) As String
    Dim result As String
    result = CommonNormalizeText(sourceText)
    result = Replace$(result, ChrW$(&HFF65), ChrW$(&H30FB))
    If removeParenthetical Then result = RemoveParentheticalText(result)
    NormalizeLineLookupText = CommonRemoveAllSpaces(result)
End Function

Private Function RemoveParentheticalText(ByVal sourceText As String) As String
    Dim result As String, depth As Long
    Dim i As Long, ch As String

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch = "(" Or ch = ChrW$(&HFF08) Then
            depth = depth + 1
        ElseIf ch = ")" Or ch = ChrW$(&HFF09) Then
            If depth > 0 Then
                depth = depth - 1
            Else
                result = result & ch
            End If
        ElseIf depth = 0 Then
            result = result & ch
        End If
    Next i

    RemoveParentheticalText = result
End Function

Private Function ResolveProjectLineMasterPath() As String
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Or Len(ThisWorkbook.Path) = 0 Then Exit Function

    Dim lineType As String, projectName As String
    lineType = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    projectName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value))
    If lineType = "" Or projectName = "" Then Exit Function

    If LCase$(Right$(projectName, 5)) <> ".xlsx" Then projectName = projectName & ".xlsx"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim documentRoot As String
    documentRoot = fso.GetParentFolderName(ThisWorkbook.Path)

    Dim candidates As Collection
    Set candidates = New Collection
    candidates.Add fso.BuildPath(documentRoot, "単価マスタ\" & PROJECT_MASTER_FOLDER & "\" & lineType & "\" & projectName)
    candidates.Add fso.BuildPath(documentRoot, "マスタデータ\" & lineType & "\" & projectName)

    Dim candidate As Variant
    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            ResolveProjectLineMasterPath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Function FindImportedUnitPriceSheetName(ByVal expectedSheetName As String) As String
    Dim normalizedExpected As String
    normalizedExpected = NormalizeLineLookupText(expectedSheetName, False)
    If normalizedExpected = "" Then Exit Function

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(ws) Then
            If StrComp(NormalizeLineLookupText(ws.Name, False), normalizedExpected, vbTextCompare) = 0 Then
                FindImportedUnitPriceSheetName = ws.Name
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function GetUnitPriceRows(ByVal unitPriceSheetName As String, _
                                  ByVal sheetPriceCaches As Object) As Object
    If sheetPriceCaches.Exists(unitPriceSheetName) Then
        Set GetUnitPriceRows = sheetPriceCaches(unitPriceSheetName)
        Exit Function
    End If

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(unitPriceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then Exit Function

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = priceSheet.Cells(priceSheet.rows.Count, COL_SEIRI).End(xlUp).Row

    Dim r As Long
    For r = UNIT_PRICE_DATA_START_ROW To lastRow
        Dim recordKey As String
        recordKey = NormalizeRecordKey(priceSheet.Cells(r, COL_SEIRI).value)
        If recordKey <> "" And Not result.Exists(recordKey) Then
            result.Add recordKey, Array(priceSheet.Cells(r, 5).value, priceSheet.Cells(r, 6).value)
        End If
    Next r

    sheetPriceCaches.Add unitPriceSheetName, result
    Set GetUnitPriceRows = result
End Function

Private Function NormalizeRecordKey(ByVal value As Variant) As String
    NormalizeRecordKey = CommonRemoveAllSpaces(CommonNzText(value))
End Function

Private Function SelectDayNightPrice(ByVal dayNightText As String, _
                                     ByVal dayNightPrices As Variant) As Variant
    Dim normalized As String
    normalized = CommonRemoveAllSpaces(CommonNormalizeText(dayNightText))

    If InStr(1, normalized, "昼", vbTextCompare) > 0 Then
        SelectDayNightPrice = dayNightPrices(0)
    ElseIf InStr(1, normalized, "夜", vbTextCompare) > 0 Then
        SelectDayNightPrice = dayNightPrices(1)
    Else
        SelectDayNightPrice = Empty
    End If
End Function

Private Sub WritePriceComparison(ByVal ws As Worksheet, ByVal rowIndex As Long)
    Dim priceMatches As Boolean
    priceMatches = UnitPriceValuesMatch(ws.Cells(rowIndex, COL_AUTO_PRICE).value, _
                                        ws.Cells(rowIndex, COL_JR_PRICE).value)

    With ws.Cells(rowIndex, COL_PRICE_COMPARE)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        If priceMatches Then
            .value = "単価一致"
            .Interior.Color = RGB(0, 255, 0)
            .Font.Color = RGB(0, 0, 255)
        Else
            .value = "単価不一致"
            .Interior.Color = RGB(255, 255, 0)
            .Font.Color = RGB(255, 0, 0)
        End If
    End With
End Sub

Private Function UnitPriceValuesMatch(ByVal leftValue As Variant, _
                                      ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then Exit Function
    If Len(Trim$(CommonNzText(leftValue))) = 0 Or Len(Trim$(CommonNzText(rightValue))) = 0 Then Exit Function
    If Not IsNumeric(leftValue) Or Not IsNumeric(rightValue) Then Exit Function

    UnitPriceValuesMatch = (Abs(CDbl(leftValue) - CDbl(rightValue)) < 0.0000001)
End Function

'==========================================================================
'  工事側シートのソート
'    1.契約線区名(側線は最後)  2.工種分類(レール溶接は最後)
'    3.昼夜別  4.整理番号
'==========================================================================
Private Sub SortWorksSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    ' 補助フラグ列(側線・レール溶接)
    Dim r As Long, lineName As String, kindName As String
    For r = 2 To lastRow
        lineName = CommonRemoveAllSpaces(CommonNzText(ws.Cells(r, COL_LINE).value))
        kindName = CommonRemoveAllSpaces(CommonNzText(ws.Cells(r, COL_KIND).value))
        ws.Cells(r, COL_FLAG_SIDE).value = IIf(InStr(1, lineName, SIDELINE_KEYWORD) > 0, 1, 0)
        ws.Cells(r, COL_FLAG_WELD).value = IIf(InStr(1, kindName, WELDING_KEYWORD) > 0, 1, 0)
    Next r

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_FLAG_SIDE), ws.Cells(lastRow, COL_FLAG_SIDE)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' 側線フラグ
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_LINE), ws.Cells(lastRow, COL_LINE)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' 契約線区名
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_FLAG_WELD), ws.Cells(lastRow, COL_FLAG_WELD)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' レール溶接フラグ
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_KIND), ws.Cells(lastRow, COL_KIND)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' 工種分類
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_DAYNIGHT), ws.Cells(lastRow, COL_DAYNIGHT)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' 昼夜別
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_SEIRI), ws.Cells(lastRow, COL_SEIRI)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal ' 整理番号
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_FLAG_WELD))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ws.Range(ws.Cells(1, COL_FLAG_SIDE), ws.Cells(lastRow, COL_FLAG_WELD)).ClearContents ' 補助列を消去
End Sub

'==========================================================================
'  購入充当側シートのソート(整理番号のみ)
'==========================================================================
Private Sub SortPurchaseSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(2, COL_SEIRI), ws.Cells(lastRow, COL_SEIRI)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_KIND))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

'==========================================================================
'  書式設定(ヘッダー装飾・罫線・数値書式・列幅自動調整)
'==========================================================================
Private Sub FormatSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 1 Then lastRow = 1

    ' ヘッダー装飾(黒地・白文字・中央揃え・太字)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, GetOutputLastColumn(ws)))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 数値書式(JR単価・JR金額・外注単価・外注金額)
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, COL_JR_PRICE), ws.Cells(lastRow, COL_OUT_AMOUNT)).NumberFormatLocal = "#,##0"
    End If

    ' 罫線(表全体)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, GetOutputLastColumn(ws))).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With

    ' 列幅自動調整(取込文字列にフィット)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, GetOutputLastColumn(ws))).EntireColumn.AutoFit

    ' 行の高さ(シート全体 18)
    ws.Cells.RowHeight = 18

    ' 中央揃え(B=整理番号, D=昼夜別, E=単位, G=契約線区名, H=管理室, M=工種分類)
    Dim centerCols As Variant, cv As Variant
    centerCols = Array(COL_SEIRI, COL_DAYNIGHT, COL_UNIT, COL_LINE, COL_MGR, COL_KIND)
    For Each cv In centerCols
        ws.Range(ws.Cells(1, CLng(cv)), ws.Cells(lastRow, CLng(cv))).HorizontalAlignment = xlCenter
    Next cv
End Sub

'==========================================================================
'  出力シートのデータ最終行(整理番号=B列で判定。A列=施工業者は空白のため)
'==========================================================================
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    GetLastDataRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).Row
End Function

'==========================================================================
'  Dictionary のキーをカンマ区切り文字列に(診断表示用)
'==========================================================================
Private Function JoinKeys(ByVal d As Object) As String
    Dim s As String
    If Not d Is Nothing Then
        Dim k As Variant
        For Each k In d.Keys
            s = s & IIf(s = "", "", ", ") & CStr(k)
        Next k
    End If
    If s = "" Then s = "(なし)"
    JoinKeys = s
End Function

'==========================================================================
'  シート名サニタイズ(禁則文字除去・31文字制限)
'==========================================================================
Private Function SanitizeSheetName(ByVal s As String) As String
    Dim t As String
    t = CommonNormalizeText(s)
    Dim bad As Variant, ch As Variant
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In bad
        t = Replace$(t, CStr(ch), "_")
    Next ch
    t = Trim$(t)
    If Len(t) > 31 Then t = Left$(t, 31)
    SanitizeSheetName = t
End Function

'==========================================================================
'  デバッグログ(イミディエイトウィンドウ出力)
'==========================================================================
Private Sub LogCI(ByVal msg As String)
    Debug.Print "[ConstructionImport] " & Format(Now, "hh:mm:ss") & "  " & msg
End Sub


