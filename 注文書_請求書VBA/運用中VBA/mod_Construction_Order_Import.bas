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
'   ※列番号は相対定義(直前の列+1)。列を挿入/削除する場合は、該当位置の Const を
'     1行追加/削除するだけで以降の列番号が自動的にずれる。併せて WriteRecordsToSheet の
'     headers 配列と rowArr への代入も同じ位置に追加/削除すること(順番厳守)。
'     絶対参照や数式オフセットはすべて定数から算出しているため、ここ以外の修正は不要。
Private Const COL_VENDOR         As Long = 1                        ' A 施工業者
Private Const COL_SEIRI          As Long = COL_VENDOR + 1           ' B 整理番号
Private Const COL_TYPE           As Long = COL_SEIRI + 1            ' C 工事種類
Private Const COL_DAYNIGHT       As Long = COL_TYPE + 1             ' D 昼夜別
Private Const COL_UNIT           As Long = COL_DAYNIGHT + 1         ' E 単位
Private Const COL_QTY            As Long = COL_UNIT + 1             ' F 数量
Private Const COL_LINE           As Long = COL_QTY + 1              ' G 契約線区名
Private Const COL_MGR            As Long = COL_LINE + 1             ' H 管理室
Private Const COL_JR_PRICE       As Long = COL_MGR + 1              ' I JR単価
Private Const COL_JR_AMOUNT      As Long = COL_JR_PRICE + 1         ' J JR金額
Private Const COL_OUT_PRICE      As Long = COL_JR_AMOUNT + 1        ' 外注単価(空白/作成後に削除)
Private Const COL_OUT_AMOUNT     As Long = COL_OUT_PRICE + 1        ' 外注金額(空白/作成後に削除)
Private Const COL_KIND           As Long = COL_OUT_AMOUNT + 1       ' 工種分類(データ最終列)
Private Const COL_GAP_AFTER_DATA As Long = COL_KIND + 1             ' 空白の区切り列
Private Const COL_AUTO_PRICE     As Long = COL_GAP_AFTER_DATA + 1   ' 参照単価
Private Const COL_AUTO_AMOUNT    As Long = COL_AUTO_PRICE + 1       ' 参照単価金額
Private Const COL_PRICE_COMPARE  As Long = COL_AUTO_AMOUNT + 1      ' 単価比較
Private Const COL_PRICE_GUIDANCE As Long = COL_PRICE_COMPARE + 1    ' 単価不一致時の案内
Private Const PRICE_GUIDANCE_COLUMN_WIDTH As Double = 59#           ' 案内表示時の列幅
Private Const COL_FLAG_SIDE      As Long = COL_PRICE_GUIDANCE + 9   ' 補助:側線フラグ(出力列より右の作業列)
Private Const COL_FLAG_WELD      As Long = COL_FLAG_SIDE + 1        ' 補助:レール溶接フラグ
Private Const OUTPUT_COL_COUNT   As Long = COL_KIND                 ' データ列数(=工種分類の列番号)

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

'--- 産廃処理行の制御 ------------------------------------------------------
Private Const SANPAI_KEYWORD As String = "産廃処理"
' 取込済み単価シートに塗りつぶしセルが無い場合の既定色(グレー)
Private Const SANPAI_FALLBACK_FILL_COLOR As Long = 14277081   ' RGB(217,217,217)

' 基本情報セル
Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_PUBLIC_CELL As String = "B4"
Private Const BASIC_INFO_AMOUNT_CELL As String = "C22"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"
Private Const PRICE_GUIDANCE_AMOUNT_TYPE_MESSAGE As String = _
    "基本情報シートのC22セル：単価適用区分(年初単価or設計変更単価)を確認して下さい。"

' 参照シート(取込元)H9 -> 基本情報 C12 転記
Private Const REF_VALUE_SOURCE_CELL As String = "H9"
Private Const BASIC_INFO_REF_VALUE_CELL As String = "C12"
Private Const BASIC_INFO_REF_FONT_NAME As String = "BIZ UDゴシック"

' 工事件名別マスタ(F列=積算線区、G列=施工指示書記載線区名)
Private Const PROJECT_MASTER_START_ROW As Long = 2
Private Const PROJECT_MASTER_UNIT_PRICE_LINE_COL As Long = 6
Private Const PROJECT_MASTER_SOURCE_LINE_COL As Long = 7
Private Const PROJECT_MASTER_FOLDER As String = "工事件名別マスタ"
Private Const UNIT_PRICE_DATA_START_ROW As Long = 7

' 購入充当単価 取込(単価適用線区シート + 名称_購入充当単価シート)
Private Const PRICE_LINE_SHEET As String = "単価適用線区"
Private Const PRICE_LINE_BRANCH_COL As Long = 2     ' B列 支店
Private Const PRICE_LINE_OFFICE_COL As Long = 3     ' C列 出張所
Private Const PRICE_LINE_NAME_COL As Long = 5       ' E列 名称
Private Const PRICE_LINE_START_ROW As Long = 2
Private Const PURCHASE_PRICE_SHEET_SUFFIX As String = "_購入充当単価"
Private Const PURCHASE_PRICE_KEY_COL As Long = 1    ' 単価シート A列(照合キー)
Private Const PURCHASE_PRICE_VALUE_COL As Long = 6  ' 単価シート F列(単価)
Private Const PURCHASE_PRICE_DATA_START_ROW As Long = 2
Private Const PURCHASE_NOTICE_SEIRI_COL As Long = 1     ' 購入充当通知 A列(整理番号)
Private Const PURCHASE_PRICE_LOOKUP_COL As Long = PURCHASE_NOTICE_SEIRI_COL
' 購入充当通知は施工業者列を削除するため、整理番号より右は共通レイアウトの1列左。
Private Const PURCHASE_NOTICE_TYPE_COL As Long = COL_TYPE - 1
Private Const PURCHASE_NOTICE_DAYNIGHT_COL As Long = COL_DAYNIGHT - 1
Private Const PURCHASE_NOTICE_UNIT_COL As Long = COL_UNIT - 1
Private Const PURCHASE_NOTICE_QTY_COL As Long = COL_QTY - 1
Private Const PURCHASE_NOTICE_LINE_COL As Long = COL_LINE - 1
Private Const PURCHASE_NOTICE_MGR_COL As Long = COL_MGR - 1
Private Const PURCHASE_NOTICE_JR_PRICE_COL As Long = COL_JR_PRICE - 1
Private Const PURCHASE_NOTICE_JR_AMOUNT_COL As Long = COL_JR_AMOUNT - 1
Private Const PURCHASE_NOTICE_OUT_PRICE_COL As Long = COL_OUT_PRICE - 1
Private Const PURCHASE_NOTICE_OUT_AMOUNT_COL As Long = COL_OUT_AMOUNT - 1
Private Const PURCHASE_NOTICE_KIND_COL As Long = COL_KIND - 1
Private Const PURCHASE_NOTICE_GAP_AFTER_DATA_COL As Long = COL_GAP_AFTER_DATA - 1
Private Const PURCHASE_NOTICE_AUTO_PRICE_COL As Long = COL_AUTO_PRICE - 1
Private Const PURCHASE_NOTICE_AUTO_AMOUNT_COL As Long = COL_AUTO_AMOUNT - 1
Private Const PURCHASE_NOTICE_PRICE_COMPARE_COL As Long = COL_PRICE_COMPARE - 1
Private Const PURCHASE_NOTICE_PRICE_GUIDANCE_COL As Long = COL_PRICE_GUIDANCE - 1

' 施工会社別単価・金額列(取込シートのJ列と工種分類列の間)
Private Const SUBCON_PRICE_FIRST_COL As Long = 11
Private Const UNIT_PRICE_VENDOR_NAME_ROW As Long = 5
Private Const UNIT_PRICE_VENDOR_FIRST_DAY_COL As Long = 7

' 取込ブックのシート名取得元セル
Private Const SOURCE_SHEET_NAME_CELL As String = "A3"

'==========================================================================
'  エントリポイント
'==========================================================================
Public Sub ImportConstructionDocument()
    Dim srcWb As Workbook
    Dim srcOpenedHere As Boolean
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

    '--- 取込前チェック: 基本情報の必須入力を確認 --------------------------
    If Not ValidateBasicInfoForImport(docType) Then GoTo Cleanup

    '--- 取込ブックを開く ---------------------------------------------------
    Set srcWb = OpenWorkbookReadOnly(srcPath, srcOpenedHere)
    If srcWb Is Nothing Then
        MsgBox "取込対象ブックを開けませんでした。" & vbCrLf & srcPath, vbExclamation
        GoTo Cleanup
    End If

    Dim srcWs As Worksheet
    Set srcWs = srcWb.ActiveSheet               ' ※ 取込ブックのアクティブシートを対象とする

    '--- 新規シート名(A3) ---------------------------------------------------
    '--- 参照シート H9 の値を取得(基本情報 C12 への転記用) -------------------
    Dim refValueH9 As Variant
    refValueH9 = srcWs.Range(REF_VALUE_SOURCE_CELL).value
    LogCI "参照シート " & REF_VALUE_SOURCE_CELL & "=[" & CommonNzText(refValueH9) & "]"

    Dim sourceA3Text As String
    sourceA3Text = CommonNzText(srcWs.Range(SOURCE_SHEET_NAME_CELL).value)

    Dim baseSheetName As String
    baseSheetName = SanitizeSheetName(sourceA3Text)
    If baseSheetName = "" Then
        MsgBox "取込ブックの " & SOURCE_SHEET_NAME_CELL & " が空のため、シート名を決定できません。", vbExclamation
        GoTo Cleanup
    End If
    LogCI "baseSheetName=[" & baseSheetName & "]"

    Dim guidanceDocumentName As String
    guidanceDocumentName = ResolveGuidanceDocumentName(sourceA3Text, docType)
    LogCI "再取込み案内文書名=[" & guidanceDocumentName & "] A3=[" & sourceA3Text & "]"

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
        Set mgrSet = BuildManagerRoomSet()
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
                Dim rowArr(0 To OUTPUT_COL_COUNT - 1) As Variant
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

    '--- 取込ブックを閉じる(以降は出力処理のみ) -----------------------------
    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Set srcWb = Nothing: srcOpenedHere = False

    '--- 工事側シート作成・書込み・ソート -----------------------------------
    '--- 参照シート H9 を 基本情報 C12 へ転記(中央揃え・BIZ UDゴシック) -------
    WriteReferenceValueToBasicInfo refValueH9

    Dim wsWorks As Worksheet
    Set wsWorks = CreateOrReplaceSheet(baseSheetName)
    If wsWorks Is Nothing Then GoTo Cleanup
    wsWorks.Tab.Color = RGB(255, 255, 0)
    WriteRecordsToSheet wsWorks, worksRows
    SortWorksSheet wsWorks
    WriteAdditionalHeaders wsWorks
    FillReferenceUnitPrices wsWorks, guidanceDocumentName
    FormatSheet wsWorks
    ApplyPriceGuidanceColumnLayout wsWorks
    mod_subcontractorselector.ApplySubcontractorDropdowns wsWorks   ' A列(施工業者)に施工会社ドロップダウンを付与
    ApplySanpaiRowRestrictions wsWorks   ' 産廃処理行のA列を塗りつぶし＋入力不可化

    ' A列(施工業者)を中央揃え
    wsWorks.Columns(COL_VENDOR).HorizontalAlignment = xlCenter

    ' 施工通知書取込のみ H列(管理室)を非表示(通知書は管理室データなし)
    If docType = DOC_NOTICE Then wsWorks.Columns(COL_MGR).Hidden = True

    ' K列(外注単価)・L列(外注金額)は不要 -> 列ごと削除し右列以降を左へ詰める
    '   ※全整形・単価転記が済んだ最後に実行。参照単価金額のR1C1相対数式は
    '     Excelが列削除に合わせて自動調整するため、数式オフセットの手修正は不要。
    wsWorks.Range(wsWorks.Cells(1, COL_OUT_PRICE), _
                  wsWorks.Cells(1, COL_OUT_AMOUNT)).EntireColumn.Delete Shift:=xlToLeft

    ' JR金額の合計行(I列:ラベル / J列:SUM)をB列最終行の直下に作成
    WriteJrTotalRow wsWorks

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
            If docType = DOC_NOTICE Then
                ApplyPurchaseNoticeLayout wsPurch
                SortPurchaseSheet wsPurch, PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_KIND_COL
                WritePurchaseNoticeAdditionalHeaders wsPurch
                FillPurchaseUnitPrices wsPurch
                FormatPurchaseNoticeSheet wsPurch
                ApplyPurchaseNoticeColumnExclusions wsPurch
                WritePurchaseNoticeJrTotalRow wsPurch   ' H列:ラベル / I列:SUM をA列最終行の直下に作成
            Else
                SortPurchaseSheet wsPurch
                FormatSheet wsPurch
                WriteJrTotalRow wsPurch   ' I列:ラベル / J列:SUM をB列最終行の直下に作成
            End If
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
'  取込前チェック: 基本情報シートの必須セルが入力済みか確認
'    未入力があれば対象セルを案内し、最初の未入力セルを選択して False を返す
'    共通必須      : C20(在来線・新幹線区分) / C21(工事種別)  ※単価・線区解決に必要
'    施工指示書のみ: B6(支店) / C6(出張所)                    ※管理室フィルタに必要
'==========================================================================
Private Function ValidateBasicInfoForImport(ByVal docType As Long) As Boolean
    ValidateBasicInfoForImport = False

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりません。", vbExclamation
        Exit Function
    End If

    ' チェック対象(セル, ラベル)。施工指示書(DOC_ORDER)のみ支店/出張所を追加
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
        ' 最初の未入力セルへ誘導(画面更新を一時的に戻して選択を見せる)
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

'==========================================================================
'  管理室集合の構築(基本情報 B6/C6 に対応する F列管理室名)
'==========================================================================
Private Function BuildManagerRoomSet() As Object
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
    Set connection = OpenUnitPriceMasterAdoConnection(masterPath)
    If connection Is Nothing Then
        MsgBox "出張所別_単価適用線区.xlsx が見つかりませんでした。", vbExclamation
        Set BuildManagerRoomSet = Nothing
        Exit Function
    End If

    On Error GoTo ErrorHandler

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(connection, MGR_MASTER_SHEET)
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
                   BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

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
'  購入充当通知専用レイアウト
'    A列(施工業者)を削除し、整理番号以降を1列左へ詰める。
'==========================================================================
Private Sub ApplyPurchaseNoticeLayout(ByVal ws As Worksheet)
    ws.Columns(COL_VENDOR).Delete Shift:=xlToLeft
    ws.Cells(1, PURCHASE_NOTICE_SEIRI_COL).value = "整理番号"
End Sub

'==========================================================================
'  追加ヘッダー(参照単価・参照金額・単価比較)
'==========================================================================
Private Sub WriteAdditionalHeaders(ByVal ws As Worksheet, _
                                   Optional ByVal includePriceComparison As Boolean = True)
    WriteAdditionalHeadersAtColumns _
        ws, COL_AUTO_PRICE, COL_AUTO_AMOUNT, COL_PRICE_COMPARE, includePriceComparison
End Sub

Private Sub WritePurchaseNoticeAdditionalHeaders(ByVal ws As Worksheet)
    WriteAdditionalHeadersAtColumns _
        ws, PURCHASE_NOTICE_AUTO_PRICE_COL, PURCHASE_NOTICE_AUTO_AMOUNT_COL, _
        PURCHASE_NOTICE_PRICE_COMPARE_COL, True
End Sub

Private Sub WriteAdditionalHeadersAtColumns( _
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

'==========================================================================
'  出力表の最終列(単価比較がある場合は案内欄)
'==========================================================================
Private Function GetOutputLastColumn( _
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

'==========================================================================
'  A列で使用中の施工会社ごとに、単価・金額の2列をJ列の右へ作成
'==========================================================================
Public Sub RefreshSubcontractorPriceColumns(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If CommonNzText(ws.Cells(1, COL_VENDOR).value) <> "施工業者" Then Exit Sub
    If CommonNzText(ws.Cells(1, COL_SEIRI).value) <> "整理番号" Then Exit Sub

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)

    Dim vendorNames As Collection
    Set vendorNames = CollectSelectedSubcontractors(ws, lastRow)

    Dim kindColumn As Long
    kindColumn = FindHeaderColumn(ws, "工種分類")
    If kindColumn = 0 Then Exit Sub

    If kindColumn > SUBCON_PRICE_FIRST_COL Then
        ws.Range(ws.Columns(SUBCON_PRICE_FIRST_COL), _
                 ws.Columns(kindColumn - 1)).Delete Shift:=xlToLeft
    End If
    If vendorNames.Count = 0 Or lastRow < 2 Then Exit Sub

    Dim insertedColumnCount As Long
    insertedColumnCount = vendorNames.Count * 2
    ws.Range(ws.Columns(SUBCON_PRICE_FIRST_COL), _
             ws.Columns(SUBCON_PRICE_FIRST_COL + insertedColumnCount - 1)).Insert Shift:=xlToRight

    Dim vendorColumnMap As Object
    Set vendorColumnMap = CreateObject("Scripting.Dictionary")
    vendorColumnMap.CompareMode = vbTextCompare

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorNames.Count
        Dim vendorName As String
        Dim priceColumn As Long
        vendorName = CStr(vendorNames(vendorIndex))
        priceColumn = SUBCON_PRICE_FIRST_COL + ((vendorIndex - 1) * 2)

        vendorColumnMap.Add NormalizeVendorPriceName(vendorName), priceColumn
        ws.Cells(1, priceColumn).value = vendorName & "単価"
        ws.Cells(1, priceColumn + 1).value = vendorName & "金額"
    Next vendorIndex

    Dim lineSheetMap As Object
    Set lineSheetMap = BuildConstructionLineSheetMap()

    Dim vendorPriceCaches As Object
    Set vendorPriceCaches = CreateObject("Scripting.Dictionary")
    vendorPriceCaches.CompareMode = vbTextCompare

    Dim matchedCount As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim rowVendorKey As String
        rowVendorKey = NormalizeVendorPriceName(CommonNzText(ws.Cells(r, COL_VENDOR).value))
        If rowVendorKey <> "" And vendorColumnMap.Exists(rowVendorKey) Then
            priceColumn = CLng(vendorColumnMap(rowVendorKey))

            Dim unitPriceSheetName As String
            unitPriceSheetName = ResolveUnitPriceSheetName( _
                lineSheetMap, CommonNzText(ws.Cells(r, COL_LINE).value))

            Dim vendorPriceRows As Object
            Set vendorPriceRows = GetVendorUnitPriceRows( _
                unitPriceSheetName, CommonNzText(ws.Cells(r, COL_VENDOR).value), vendorPriceCaches)

            Dim recordKey As String
            recordKey = NormalizeRecordKey(ws.Cells(r, COL_SEIRI).value)
            If Not vendorPriceRows Is Nothing And recordKey <> "" Then
                If vendorPriceRows.Exists(recordKey) Then
                    Dim dayNightPrices As Variant
                    Dim vendorPrice As Variant
                    dayNightPrices = vendorPriceRows(recordKey)
                    vendorPrice = SelectDayNightPrice( _
                        CommonNzText(ws.Cells(r, COL_DAYNIGHT).value), dayNightPrices)
                    If Not IsEmpty(vendorPrice) And Not IsError(vendorPrice) Then
                        ws.Cells(r, priceColumn).value = vendorPrice
                        matchedCount = matchedCount + 1
                    End If
                End If
            End If

            ws.Cells(r, priceColumn + 1).FormulaR1C1 = _
                "=IF(OR(RC[-1]="""",RC" & COL_QTY & "=""""),"""",RC[-1]*RC" & COL_QTY & ")"
        End If
    Next r

    ' 施工会社ごとの合計(単価列:「会社名+合計」/ 金額列:SUM)をB列最終行の直下に作成
    '   ※列は毎回削除→再挿入されるため、合計セルも実行のたびに再作成される
    For vendorIndex = 1 To vendorNames.Count
        priceColumn = SUBCON_PRICE_FIRST_COL + ((vendorIndex - 1) * 2)
        WriteTotalCells ws, lastRow + 1, _
                        priceColumn, CStr(vendorNames(vendorIndex)) & "合計", _
                        priceColumn + 1, lastRow
    Next vendorIndex

    FormatSubcontractorPriceColumns ws, lastRow, insertedColumnCount
    LogCI "施工会社別単価列: 会社数=" & vendorNames.Count & _
          " / 単価一致=" & matchedCount
End Sub

Private Function CollectSelectedSubcontractors(ByVal ws As Worksheet, _
                                                ByVal lastRow As Long) As Collection
    Dim result As New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim r As Long
    For r = 2 To lastRow
        Dim vendorName As String
        Dim vendorKey As String
        vendorName = Trim$(CommonNzText(ws.Cells(r, COL_VENDOR).value))
        vendorKey = NormalizeVendorPriceName(vendorName)
        If vendorKey <> "" And Not seen.Exists(vendorKey) Then
            seen.Add vendorKey, True
            result.Add vendorName
        End If
    Next r

    Set CollectSelectedSubcontractors = result
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim hit As Range
    Set hit = ws.rows(1).Find(What:=headerText, After:=ws.Cells(1, 1), _
                              LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                              MatchCase:=False)
    If Not hit Is Nothing Then FindHeaderColumn = hit.Column
End Function

Private Function GetVendorUnitPriceRows(ByVal unitPriceSheetName As String, _
                                        ByVal vendorName As String, _
                                        ByVal vendorPriceCaches As Object) As Object
    If unitPriceSheetName = "" Then Exit Function

    Dim cacheKey As String
    cacheKey = unitPriceSheetName & "|" & NormalizeVendorPriceName(vendorName)
    If vendorPriceCaches.Exists(cacheKey) Then
        Set GetVendorUnitPriceRows = vendorPriceCaches(cacheKey)
        Exit Function
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(unitPriceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then
        vendorPriceCaches.Add cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim vendorDayColumn As Long
    vendorDayColumn = FindUnitPriceVendorDayColumn(priceSheet, vendorName)
    If vendorDayColumn = 0 Then
        vendorPriceCaches.Add cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim priceLastRow As Long
    priceLastRow = priceSheet.Cells(priceSheet.rows.Count, COL_SEIRI).End(xlUp).Row

    Dim r As Long
    For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
        Dim recordKey As String
        recordKey = NormalizeRecordKey(priceSheet.Cells(r, COL_SEIRI).value)
        If recordKey <> "" And Not result.Exists(recordKey) Then
            result.Add recordKey, Array(priceSheet.Cells(r, vendorDayColumn).value, _
                                        priceSheet.Cells(r, vendorDayColumn + 1).value)
        End If
    Next r

    vendorPriceCaches.Add cacheKey, result
    Set GetVendorUnitPriceRows = result
End Function

Private Function FindUnitPriceVendorDayColumn(ByVal priceSheet As Worksheet, _
                                              ByVal vendorName As String) As Long
    Dim vendorKey As String
    vendorKey = NormalizeVendorPriceName(vendorName)
    If vendorKey = "" Then Exit Function

    Dim lastColumn As Long
    lastColumn = priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, _
                                  priceSheet.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = UNIT_PRICE_VENDOR_FIRST_DAY_COL To lastColumn
        If StrComp(NormalizeVendorPriceName( _
                       CommonNzText(priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, c).value)), _
                   vendorKey, vbTextCompare) = 0 Then
            FindUnitPriceVendorDayColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function NormalizeVendorPriceName(ByVal vendorName As String) As String
    NormalizeVendorPriceName = CommonRemoveAllSpaces(CommonNormalizeText(vendorName))
End Function

Private Sub FormatSubcontractorPriceColumns(ByVal ws As Worksheet, _
                                            ByVal lastRow As Long, _
                                            ByVal columnCount As Long)
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = SUBCON_PRICE_FIRST_COL
    lastColumn = firstColumn + columnCount - 1

    With ws.Range(ws.Cells(1, firstColumn), ws.Cells(1, lastColumn))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With

    If lastRow >= 2 Then
        With ws.Range(ws.Cells(2, firstColumn), ws.Cells(lastRow, lastColumn))
            .NumberFormatLocal = "#,##0;[赤]-#,##0"
        End With
        With ws.Range(ws.Cells(1, firstColumn), ws.Cells(lastRow, lastColumn)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
    End If

    ' 産廃処理行: 追加された施工会社別単価・金額列を塗りつぶし(単価シートと同色)
    If lastRow >= 2 Then
        Dim sanpaiFillColor As Long
        Dim sanpaiRow As Long
        sanpaiFillColor = GetSanpaiFillColor()
        For sanpaiRow = 2 To lastRow
            If IsSanpaiRow(ws, sanpaiRow) Then
                ws.Range(ws.Cells(sanpaiRow, firstColumn), _
                         ws.Cells(sanpaiRow, lastColumn)).Interior.Color = sanpaiFillColor
            End If
        Next sanpaiRow
    End If

    ws.Range(ws.Columns(firstColumn), ws.Columns(lastColumn)).AutoFit
End Sub

'==========================================================================
'  産廃処理行: C列(工事種類)に「産廃処理」を含む行か判定
'==========================================================================
Private Function IsSanpaiRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    IsSanpaiRow = (InStr(1, CommonRemoveAllSpaces(CommonNzText(ws.Cells(rowIndex, COL_TYPE).value)), _
                         SANPAI_KEYWORD, vbTextCompare) > 0)
End Function

'==========================================================================
'  産廃処理行の塗りつぶし色を取得
'    取込済み単価シートのE/F列(昼・夜単価)で最初に見つかった
'    塗りつぶしセルの色を採用。見つからない場合は既定のグレー。
'==========================================================================
Private Function GetSanpaiFillColor() As Long
    GetSanpaiFillColor = SANPAI_FALLBACK_FILL_COLOR

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(ws) Then
            Dim priceLastRow As Long
            priceLastRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).Row

            Dim r As Long, c As Long
            For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
                For c = 5 To 6                      ' E:昼単価 / F:夜単価
                    If ws.Cells(r, c).Interior.Pattern <> xlPatternNone Then
                        GetSanpaiFillColor = ws.Cells(r, c).Interior.Color
                        Exit Function
                    End If
                Next c
            Next r
        End If
    Next ws
End Function

'==========================================================================
'  産廃処理行: A列(施工業者)を塗りつぶし＋入力不可化
'    施工業者ドロップダウン付与(ApplySubcontractorDropdowns)の後に
'    呼び出すこと(A列の入力規則を上書きするため)。
'    入力不可はユーザー設定の入力規則(=FALSE)で実現。
'    ※Excelの仕様上、コピー＆貼り付けは防げない点に注意。
'==========================================================================
Public Sub ApplySanpaiRowRestrictions(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim fillColor As Long
    fillColor = GetSanpaiFillColor()

    Dim restrictedCount As Long
    Dim r As Long
    For r = 2 To lastRow
        If IsSanpaiRow(ws, r) Then
            With ws.Cells(r, COL_VENDOR)
                .ClearContents
                .Interior.Color = fillColor
                With .Validation
                    .Delete
                    .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                         Formula1:="=FALSE"
                    .IgnoreBlank = True
                    .InCellDropdown = False
                    .ShowInput = False
                    .ErrorTitle = "入力不可"
                    .ErrorMessage = "産廃処理の行は施工業者を入力できません。"
                    .ShowError = True
                End With
            End With
            restrictedCount = restrictedCount + 1
        End If
    Next r

    LogCI "産廃処理行: A列塗りつぶし・入力不可=" & restrictedCount & " 行"
End Sub

'==========================================================================
'  合計セルの共通書込み
'    ラベルセル: 黒のxlThin枠で囲む
'    SUMセル   : 小数点以下切り捨て(ROUNDDOWN)・桁区切り表示・
'                赤のxlMedium枠で囲む
'==========================================================================
Private Sub WriteTotalCells(ByVal ws As Worksheet, ByVal totalRow As Long, _
                            ByVal labelColumn As Long, ByVal labelText As String, _
                            ByVal sumColumn As Long, ByVal sumLastRow As Long)
    ' ラベルセル
    With ws.Cells(totalRow, labelColumn)
        .value = labelText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(0, 0, 0)
    End With

    ' SUMセル(小数点以下切り捨て・桁区切り・赤の中太枠)
    With ws.Cells(totalRow, sumColumn)
        .FormulaR1C1 = "=ROUNDDOWN(SUM(R2C:R" & sumLastRow & "C),0)"
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=RGB(255, 0, 0)
    End With
End Sub

'==========================================================================
'  工事側シート: JR金額の合計行を作成
'    B列(整理番号)の最終行の直下に、I列(JR単価列)へ「JR合計」、
'    J列(JR金額列)へ J2:J最終行 のSUMを設定。
'==========================================================================
Private Sub WriteJrTotalRow(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    WriteTotalCells ws, lastRow + 1, _
                    COL_JR_PRICE, "JR合計", _
                    COL_JR_AMOUNT, lastRow
End Sub

'==========================================================================
'  購入充当通知シート: JR金額の合計行を作成
'    A列(整理番号)の最終行の直下に、H列(JR単価列)へ「JR合計」、
'    I列(JR金額列)へ I2:I最終行 のSUMを設定。
'==========================================================================
Private Sub WritePurchaseNoticeJrTotalRow(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, PURCHASE_NOTICE_SEIRI_COL)
    If lastRow < 2 Then Exit Sub

    WriteTotalCells ws, lastRow + 1, _
                    PURCHASE_NOTICE_JR_PRICE_COL, "JR合計", _
                    PURCHASE_NOTICE_JR_AMOUNT_COL, lastRow
End Sub

'==========================================================================
'  工事側シートの参照単価・金額・比較結果を設定
'==========================================================================
Private Sub FillReferenceUnitPrices(ByVal ws As Worksheet, _
                                    ByVal guidanceDocumentName As String)
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
        WritePriceComparison ws, r, unitPriceSheetName, True, guidanceDocumentName
    Next r

    ws.Range(ws.Cells(2, COL_AUTO_AMOUNT), ws.Cells(lastRow, COL_AUTO_AMOUNT)).FormulaR1C1 = _
        "=IF(OR(RC[" & (COL_AUTO_PRICE - COL_AUTO_AMOUNT) & "]="""",RC[" & (COL_QTY - COL_AUTO_AMOUNT) & "]=""""),"""",RC[" & (COL_AUTO_PRICE - COL_AUTO_AMOUNT) & "]*RC[" & (COL_QTY - COL_AUTO_AMOUNT) & "])"

    With ws.Range(ws.Cells(2, COL_AUTO_PRICE), ws.Cells(lastRow, COL_AUTO_AMOUNT))
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
    End With
    ws.Range(ws.Cells(1, COL_PRICE_COMPARE), ws.Cells(lastRow, COL_PRICE_COMPARE)).HorizontalAlignment = xlCenter

    LogCI "参照単価一致=" & matchedCount & _
          " / 線区未解決=" & unresolvedLineCount & _
          " / 整理番号未一致=" & missingRecordCount
End Sub

'==========================================================================
'  適用積算線区単価シートの追記・単価変更を施工指示書/施工通知書へ反映
'==========================================================================
Public Sub RefreshConstructionReferencePricesForUnitPriceChange( _
    ByVal wsUnitPrice As Worksheet, ByVal changedRange As Range)

    If wsUnitPrice Is Nothing Or changedRange Is Nothing Then Exit Sub
    If Not mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then Exit Sub

    Dim monitored As Range
    Set monitored = Application.Intersect(changedRange, _
        Application.Union(wsUnitPrice.Columns(COL_SEIRI), _
                          wsUnitPrice.Columns(5), _
                          wsUnitPrice.Columns(6)))
    If monitored Is Nothing Then Exit Sub

    Dim changedPriceRows As Object
    Set changedPriceRows = CreateObject("Scripting.Dictionary")
    changedPriceRows.CompareMode = vbTextCompare

    Dim changedCell As Range
    Dim recordKey As String
    For Each changedCell In monitored.Cells
        If changedCell.Row >= UNIT_PRICE_DATA_START_ROW Then
            recordKey = NormalizeRecordKey(wsUnitPrice.Cells(changedCell.Row, COL_SEIRI).value)
            If recordKey <> "" Then
                changedPriceRows(recordKey) = Array( _
                    wsUnitPrice.Cells(changedCell.Row, 5).value, _
                    wsUnitPrice.Cells(changedCell.Row, 6).value)
            End If
        End If
    Next changedCell

    If changedPriceRows.Count = 0 Then Exit Sub

    Dim lineSheetMap As Object
    Set lineSheetMap = BuildConstructionLineSheetMap()

    Dim wsTarget As Worksheet
    For Each wsTarget In ThisWorkbook.worksheets
        If IsConstructionDocumentOutputSheet(wsTarget) Then
            RefreshConstructionReferencePricesOnSheet _
                wsTarget, wsUnitPrice.Name, changedPriceRows, lineSheetMap
        End If
    Next wsTarget
End Sub

Private Function IsConstructionDocumentOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Trim$(CommonNzText(ws.Cells(1, 1).value)) <> "施工業者" Then Exit Function
    If Trim$(CommonNzText(ws.Cells(1, 2).value)) <> "整理番号" Then Exit Function
    IsConstructionDocumentOutputSheet = (FindHeaderColumn(ws, "単価比較") > 0)
End Function

Private Sub RefreshConstructionReferencePricesOnSheet( _
    ByVal ws As Worksheet, _
    ByVal unitPriceSheetName As String, _
    ByVal changedPriceRows As Object, _
    ByVal lineSheetMap As Object)

    Dim seiriColumn As Long
    Dim dayNightColumn As Long
    Dim quantityColumn As Long
    Dim jrPriceColumn As Long
    Dim comparisonColumn As Long
    seiriColumn = FindHeaderColumn(ws, "整理番号")
    dayNightColumn = FindHeaderColumn(ws, "昼夜別")
    quantityColumn = FindHeaderColumn(ws, "数量")
    jrPriceColumn = FindHeaderColumn(ws, "JR単価")
    comparisonColumn = FindHeaderColumn(ws, "単価比較")

    If seiriColumn = 0 Or dayNightColumn = 0 Or quantityColumn = 0 Then Exit Sub
    If jrPriceColumn = 0 Or comparisonColumn < 3 Then Exit Sub

    Dim autoPriceColumn As Long
    Dim autoAmountColumn As Long
    Dim guidanceColumn As Long
    autoPriceColumn = comparisonColumn - 2
    autoAmountColumn = comparisonColumn - 1
    guidanceColumn = comparisonColumn + 1

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, seiriColumn).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim normalizedSourceSheet As String
    normalizedSourceSheet = NormalizeLineLookupText(unitPriceSheetName, False)

    Dim r As Long
    Dim recordKey As String
    Dim resolvedSheetName As String
    Dim referencePrice As Variant
    For r = 2 To lastRow
        recordKey = NormalizeRecordKey(ws.Cells(r, seiriColumn).value)
        If recordKey <> "" And changedPriceRows.Exists(recordKey) Then
            resolvedSheetName = ResolveUnitPriceSheetName( _
                lineSheetMap, CommonNzText(ws.Cells(r, COL_LINE).value))

            If NormalizeLineLookupText(resolvedSheetName, False) = normalizedSourceSheet Then
                referencePrice = SelectDayNightPrice( _
                    CommonNzText(ws.Cells(r, dayNightColumn).value), _
                    changedPriceRows(recordKey))

                If IsEmpty(referencePrice) Or IsError(referencePrice) Then
                    ws.Cells(r, autoPriceColumn).ClearContents
                Else
                    ws.Cells(r, autoPriceColumn).value = referencePrice
                End If

                ws.Cells(r, autoAmountColumn).FormulaR1C1 = _
                    "=IF(OR(RC[" & (autoPriceColumn - autoAmountColumn) & "]=""""," & _
                    "RC[" & (quantityColumn - autoAmountColumn) & "]=""""),""""," & _
                    "RC[" & (autoPriceColumn - autoAmountColumn) & "]*" & _
                    "RC[" & (quantityColumn - autoAmountColumn) & "])"

                WritePriceComparisonAtColumns _
                    ws, r, unitPriceSheetName, autoPriceColumn, jrPriceColumn, _
                    comparisonColumn, guidanceColumn, True
            End If
        End If
    Next r

    ws.Range(ws.Cells(2, autoPriceColumn), _
             ws.Cells(lastRow, autoAmountColumn)).NumberFormatLocal = "#,##0;[赤]-#,##0"
    ws.Range(ws.Cells(1, comparisonColumn), _
             ws.Cells(lastRow, comparisonColumn)).HorizontalAlignment = xlCenter
    ApplyPriceGuidanceColumnLayoutAtColumns ws, comparisonColumn, guidanceColumn
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

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then
        LogCI "工事件名別マスタへADO接続できない path=[" & masterPath & "]"
        Set BuildConstructionLineSheetMap = result
        Exit Function
    End If

    On Error GoTo Cleanup

    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim sheetName As Variant
    Dim recordset As Object
    For Each sheetName In sheetNames
        Set recordset = CreateObject("ADODB.Recordset")
        recordset.Open "SELECT [F6], [F7] FROM " & _
                       BuildAdoSheetTableName(CStr(sheetName)), connection, 0, 1, 1

        Dim rowNumber As Long
        rowNumber = 1
        Do Until recordset.EOF
            Dim unitPriceLineName As String
            Dim sourceLineName As String
            Dim actualSheetName As String

            unitPriceLineName = CommonNzText(recordset.Fields(0).value)
            If rowNumber >= PROJECT_MASTER_START_ROW And Trim$(unitPriceLineName) <> "" Then
                actualSheetName = FindImportedUnitPriceSheetName(unitPriceLineName)
                If actualSheetName <> "" Then
                    sourceLineName = CommonNzText(recordset.Fields(1).value)
                    If Trim$(sourceLineName) = "" Then sourceLineName = unitPriceLineName

                    AddLineSheetAliases result, sourceLineName, actualSheetName
                    AddLineSheetAliases result, unitPriceLineName, actualSheetName
                End If
            End If
            recordset.MoveNext
            rowNumber = rowNumber + 1
        Loop

        CommonCloseAdoRecordset recordset
        Set recordset = Nothing
    Next sheetName

Cleanup:
    If Err.Number <> 0 Then LogCI "線区名マスタ読込エラー Err " & Err.Number & ": " & Err.Description
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection

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
    If wsInfo Is Nothing Then
        LogCI "工事件名別マスタ解決: 基本情報シートなし"
        Exit Function
    End If

    Dim lineType As String, projectName As String
    lineType = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    projectName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value))
    LogCI "工事件名別マスタ解決 lineType=[" & lineType & "] projectName=[" & projectName & _
          "] workbookPath=[" & ThisWorkbook.Path & "]"
    If lineType = "" Or projectName = "" Then
        LogCI "工事件名別マスタ解決: C20またはC21が空"
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim documentRoots As Collection
    Set documentRoots = New Collection

    If Len(ThisWorkbook.Path) > 0 Then
        AddUniqueText documentRoots, ThisWorkbook.Path
        AddUniqueText documentRoots, fso.GetParentFolderName(ThisWorkbook.Path)
    End If

    Dim managerMasterPath As String
    managerMasterPath = ResolveMasterFilePath()
    If managerMasterPath <> "" Then
        AddUniqueText documentRoots, fso.GetParentFolderName(fso.GetParentFolderName(managerMasterPath))
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText documentRoots, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
                     "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
    End If

    Dim documentRoot As Variant
    For Each documentRoot In documentRoots
        Dim masterFolders As Collection
        Set masterFolders = New Collection
        AddUniqueText masterFolders, fso.BuildPath(CStr(documentRoot), _
                      "単価マスタ\" & PROJECT_MASTER_FOLDER & "\" & lineType)
        AddUniqueText masterFolders, fso.BuildPath(CStr(documentRoot), _
                      "マスタデータ\" & lineType)

        Dim masterFolder As Variant
        For Each masterFolder In masterFolders
            LogCI "工事件名別マスタ探索 folder=[" & CStr(masterFolder) & "]"
            ResolveProjectLineMasterPath = FindProjectMasterFile(CStr(masterFolder), projectName)
            If ResolveProjectLineMasterPath <> "" Then
                LogCI "工事件名別マスタ解決 path=[" & ResolveProjectLineMasterPath & "]"
                Exit Function
            End If
        Next masterFolder
    Next documentRoot

    LogCI "工事件名別マスタ解決失敗 lineType=[" & lineType & "] projectName=[" & projectName & "]"
End Function

Private Function FindProjectMasterFile(ByVal masterFolder As String, _
                                       ByVal projectName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(masterFolder) Then Exit Function

    Dim requestedFileName As String
    requestedFileName = projectName
    If LCase$(Right$(requestedFileName, 5)) <> ".xlsx" Then
        requestedFileName = requestedFileName & ".xlsx"
    End If

    Dim exactPath As String
    exactPath = fso.BuildPath(masterFolder, requestedFileName)
    If fso.FileExists(exactPath) Then
        FindProjectMasterFile = exactPath
        Exit Function
    End If

    Dim requestedKey As String
    requestedKey = NormalizeProjectMasterName(projectName)
    If requestedKey = "" Then Exit Function

    Dim sourceFile As Object
    For Each sourceFile In fso.GetFolder(masterFolder).Files
        If LCase$(fso.GetExtensionName(sourceFile.Name)) = "xlsx" Then
            If StrComp(NormalizeProjectMasterName(fso.GetBaseName(sourceFile.Name)), _
                       requestedKey, vbTextCompare) = 0 Then
                FindProjectMasterFile = sourceFile.Path
                Exit Function
            End If
        End If
    Next sourceFile
End Function

Private Function NormalizeProjectMasterName(ByVal sourceText As String) As String
    Dim result As String
    result = CommonRemoveAllSpaces(CommonNormalizeText(sourceText))
    If LCase$(Right$(result, 5)) = ".xlsx" Then result = Left$(result, Len(result) - 5)
    result = Replace$(result, ChrW$(&H30FB), "")
    result = Replace$(result, ChrW$(&HFF65), "")
    NormalizeProjectMasterName = result
End Function

Private Sub AddUniqueText(ByVal values As Collection, ByVal newValue As String)
    If values Is Nothing Or Len(Trim$(newValue)) = 0 Then Exit Sub

    Dim item As Variant
    For Each item In values
        If StrComp(CStr(item), newValue, vbTextCompare) = 0 Then Exit Sub
    Next item

    values.Add newValue
End Sub

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

Private Sub WritePriceComparison(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                                 ByVal unitPriceSheetName As String, _
                                 Optional ByVal includeGuidance As Boolean = True, _
                                 Optional ByVal guidanceDocumentName As String = "")
    WritePriceComparisonAtColumns _
        ws, rowIndex, unitPriceSheetName, COL_AUTO_PRICE, COL_JR_PRICE, _
        COL_PRICE_COMPARE, COL_PRICE_GUIDANCE, includeGuidance
End Sub

Private Sub WritePriceComparisonAtColumns( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal unitPriceSheetName As String, _
    ByVal autoPriceColumn As Long, _
    ByVal jrPriceColumn As Long, _
    ByVal comparisonColumn As Long, _
    ByVal guidanceColumn As Long, _
    ByVal includeGuidance As Boolean)

    Dim priceMatches As Boolean
    priceMatches = UnitPriceValuesMatch(ws.Cells(rowIndex, autoPriceColumn).value, _
                                        ws.Cells(rowIndex, jrPriceColumn).value)

    With ws.Cells(rowIndex, guidanceColumn)
        .ClearContents
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
        .WrapText = False
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(rowIndex, comparisonColumn)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        If priceMatches Then
            .value = "単価一致"
            .Interior.Color = RGB(0, 255, 0)
            .Font.Color = RGB(0, 0, 255)
        Else
            .value = "単価不一致"
            .Interior.Color = RGB(255, 255, 0)
            .Font.Color = RGB(255, 0, 0)

            If includeGuidance Then
                With ws.Cells(rowIndex, guidanceColumn)
                    If Len(Trim$(CommonNzText(ws.Cells(rowIndex, autoPriceColumn).value))) = 0 Then
                        Dim guidanceSheetName As String
                        guidanceSheetName = unitPriceSheetName
                        If guidanceSheetName = "" Then
                            guidanceSheetName = CommonNzText(ws.Cells(rowIndex, COL_LINE).value)
                        End If
                        .value = "独自工種の内容を" & guidanceSheetName & "シートに入力してください。"
                        .Font.Color = RGB(255, 0, 0)
                    Else
                        .value = PRICE_GUIDANCE_AMOUNT_TYPE_MESSAGE
                        .Font.Color = RGB(0, 0, 255)
                    End If
                    .Font.Bold = True
                    .WrapText = False
                    .VerticalAlignment = xlCenter
                End With
            End If
        End If
    End With
End Sub

'==========================================================================
'  取込元A3から再取込み案内に表示する文書名を決定する
'==========================================================================
Private Function ResolveGuidanceDocumentName(ByVal sourceA3Text As String, _
                                             ByVal docType As Long) As String
    Dim normalized As String
    normalized = CommonRemoveAllSpaces(CommonNormalizeText(sourceA3Text))

    If InStr(1, normalized, "通知", vbTextCompare) > 0 Then
        ResolveGuidanceDocumentName = "施工通知書"
    ElseIf InStr(1, normalized, "指示", vbTextCompare) > 0 Then
        ResolveGuidanceDocumentName = "施工指示書"
    ElseIf docType = DOC_ORDER Then
        ResolveGuidanceDocumentName = "施工指示書"
    Else
        ResolveGuidanceDocumentName = "施工通知書"
    End If
End Function

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
Private Sub SortPurchaseSheet(ByVal ws As Worksheet, _
                              Optional ByVal seiriColumn As Long = COL_SEIRI, _
                              Optional ByVal lastDataColumn As Long = COL_KIND)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, seiriColumn)
    If lastRow < 2 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(2, seiriColumn), ws.Cells(lastRow, seiriColumn)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastDataColumn))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

'==========================================================================
'  書式設定(ヘッダー装飾・罫線・数値書式・列幅自動調整)
'==========================================================================
Private Sub FormatSheet(ByVal ws As Worksheet, _
                        Optional ByVal dataKeyColumn As Long = COL_SEIRI)
    FormatSheetAtColumns _
        ws, dataKeyColumn, COL_DAYNIGHT, COL_UNIT, COL_LINE, COL_MGR, _
        COL_JR_PRICE, COL_OUT_AMOUNT, COL_KIND, COL_AUTO_PRICE, _
        COL_AUTO_AMOUNT, COL_PRICE_COMPARE, COL_PRICE_GUIDANCE
End Sub

Private Sub FormatPurchaseNoticeSheet(ByVal ws As Worksheet)
    FormatSheetAtColumns _
        ws, PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_DAYNIGHT_COL, _
        PURCHASE_NOTICE_UNIT_COL, PURCHASE_NOTICE_LINE_COL, PURCHASE_NOTICE_MGR_COL, _
        PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_OUT_AMOUNT_COL, _
        PURCHASE_NOTICE_KIND_COL, PURCHASE_NOTICE_AUTO_PRICE_COL, _
        PURCHASE_NOTICE_AUTO_AMOUNT_COL, PURCHASE_NOTICE_PRICE_COMPARE_COL, _
        PURCHASE_NOTICE_PRICE_GUIDANCE_COL
End Sub

Private Sub FormatSheetAtColumns( _
    ByVal ws As Worksheet, _
    ByVal dataKeyColumn As Long, _
    ByVal dayNightColumn As Long, _
    ByVal unitColumn As Long, _
    ByVal lineColumn As Long, _
    ByVal mgrColumn As Long, _
    ByVal jrPriceColumn As Long, _
    ByVal outAmountColumn As Long, _
    ByVal kindColumn As Long, _
    ByVal autoPriceColumn As Long, _
    ByVal autoAmountColumn As Long, _
    ByVal comparisonColumn As Long, _
    ByVal guidanceColumn As Long)

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, dataKeyColumn)
    If lastRow < 1 Then lastRow = 1

    Dim outputLastColumn As Long
    outputLastColumn = GetOutputLastColumn( _
        ws, kindColumn, autoAmountColumn, comparisonColumn, guidanceColumn)

    ' ヘッダー装飾(黒地・白文字・中央揃え・太字)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, outputLastColumn))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 数値書式(JR単価・JR金額・外注単価・外注金額)
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, jrPriceColumn), _
                 ws.Cells(lastRow, outAmountColumn)).NumberFormatLocal = "#,##0"
    End If

    ' データ列と参照単価列に罫線を設定。空白区切り列と案内欄は対象外。
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, kindColumn)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
    If CommonNzText(ws.Cells(1, autoAmountColumn).value) <> "" Then
        Dim borderLastColumn As Long
        borderLastColumn = autoAmountColumn
        If CommonNzText(ws.Cells(1, comparisonColumn).value) <> "" Then
            borderLastColumn = comparisonColumn
        End If
        With ws.Range(ws.Cells(1, autoPriceColumn), _
                      ws.Cells(lastRow, borderLastColumn)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
        If borderLastColumn = comparisonColumn Then
            ws.Range(ws.Cells(1, guidanceColumn), _
                     ws.Cells(lastRow, guidanceColumn)).Borders.LineStyle = xlNone
        End If
    End If

    ' 列幅自動調整(取込文字列にフィット)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, outputLastColumn)).EntireColumn.AutoFit

    ' 参照単価列は参照金額列と同じ列幅にし、ヘッダーだけ縮小表示する
    If CommonNzText(ws.Cells(1, autoAmountColumn).value) <> "" Then
        ws.Columns(autoPriceColumn).ShrinkToFit = False
        ws.Columns(autoPriceColumn).ColumnWidth = ws.Columns(autoAmountColumn).ColumnWidth
        ws.Cells(1, autoPriceColumn).ShrinkToFit = True
    End If

    ' 注意文は折り返さないため、単価不一致行も含めて行の高さは18で統一する
    ws.Cells.RowHeight = 18

    ' 中央揃え(整理番号, 昼夜別, 単位, 契約線区名, 管理室, 工種分類)
    Dim centerCols As Variant, cv As Variant
    centerCols = Array(dataKeyColumn, dayNightColumn, unitColumn, lineColumn, mgrColumn, kindColumn)
    For Each cv In centerCols
        ws.Range(ws.Cells(1, CLng(cv)), ws.Cells(lastRow, CLng(cv))).HorizontalAlignment = xlCenter
    Next cv
End Sub

'==========================================================================
'  工事側シートの案内列(R列)をメッセージ有無に応じて表示切替
'==========================================================================
Private Sub ApplyPriceGuidanceColumnLayout(ByVal ws As Worksheet)
    ApplyPriceGuidanceColumnLayoutAtColumns _
        ws, COL_PRICE_COMPARE, COL_PRICE_GUIDANCE
End Sub

Private Sub ApplyPriceGuidanceColumnLayoutAtColumns( _
    ByVal ws As Worksheet, _
    ByVal comparisonColumn As Long, _
    ByVal guidanceColumn As Long)

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)

    Dim hasGuidance As Boolean
    Dim r As Long
    For r = 2 To lastRow
        If CommonNzText(ws.Cells(r, comparisonColumn).value) = "単価不一致" And _
           CommonNzText(ws.Cells(r, guidanceColumn).value) <> "" Then
            hasGuidance = True
            Exit For
        End If
    Next r

    With ws.Columns(guidanceColumn)
        .Hidden = Not hasGuidance
        If hasGuidance Then .AutoFit
    End With
End Sub

'==========================================================================
'  購入充当通知では後工程で使用しない列を非表示にする
'==========================================================================
Private Sub ApplyPurchaseNoticeColumnExclusions(ByVal ws As Worksheet)
    Dim excludedColumns As Range
    Set excludedColumns = Application.Union(ws.Columns(PURCHASE_NOTICE_MGR_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_PRICE_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_AMOUNT_COL))
    excludedColumns.EntireColumn.Hidden = True
End Sub

'==========================================================================
'  出力シートのデータ最終行(指定された整理番号列で判定)
'==========================================================================
Private Function GetLastDataRow(ByVal ws As Worksheet, _
                                Optional ByVal dataKeyColumn As Long = COL_SEIRI) As Long
    GetLastDataRow = ws.Cells(ws.rows.Count, dataKeyColumn).End(xlUp).Row
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

'==========================================================================
'  参照シート H9 の値を 基本情報 C12 へ転記(中央揃え・BIZ UDゴシック)
'==========================================================================
'==========================================================================
'  購入充当通知シートへ単価(F列)を N列へ取込み
'    照合元: 購入充当通知 A列(整理番号)  <->  名称_購入充当単価 A列
'    取込値: 名称_購入充当単価 F列  ->  購入充当通知 N列
'    O列: N列(参照単価) * E列(数量)
'    P列: N列(参照単価) と H列(JR単価)の比較結果
'    行範囲: 2 ~ A列(整理番号)の最終行
'==========================================================================
Private Sub FillPurchaseUnitPrices(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, PURCHASE_NOTICE_SEIRI_COL).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' 基本情報 B6/C6 -> 単価適用線区マスタ -> 名称_購入充当単価 シート名
    Dim priceSheetName As String
    priceSheetName = ResolvePurchasePriceSheetName()
    If priceSheetName = "" Then
        LogCI "購入充当単価: シート名の解決に失敗(基本情報B6/C6 または 単価適用線区マスタ未一致)"
        Exit Sub
    End If

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(priceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then
        LogCI "購入充当単価: シート「" & priceSheetName & "」が見つかりません"
        MsgBox "単価シート「" & priceSheetName & "」が見つかりませんでした。" & vbCrLf & _
               "購入充当単価の取込みをスキップします。", vbExclamation
        Exit Sub
    End If

    ' 名称_購入充当単価 シートの A列(キー) -> F列(単価) 辞書を構築
    Dim priceMap As Object
    Set priceMap = CreateObject("Scripting.Dictionary")
    priceMap.CompareMode = vbTextCompare

    Dim pLast As Long, pr As Long, key As String
    pLast = priceSheet.Cells(priceSheet.rows.Count, PURCHASE_PRICE_KEY_COL).End(xlUp).Row
    For pr = PURCHASE_PRICE_DATA_START_ROW To pLast
        key = CommonRemoveAllSpaces(CommonNzText(priceSheet.Cells(pr, PURCHASE_PRICE_KEY_COL).value))
        If key <> "" And Not priceMap.Exists(key) Then
            priceMap.Add key, priceSheet.Cells(pr, PURCHASE_PRICE_VALUE_COL).value
        End If
    Next pr

    ' 購入充当通知 A列(整理番号) と照合し N列へ転記
    Dim matched As Long, lookupCount As Long, unmatched As Long
    Dim r As Long, lookupKey As String
    For r = 2 To lastRow
        lookupKey = NormalizeRecordKey(ws.Cells(r, PURCHASE_PRICE_LOOKUP_COL).value)
        If lookupKey <> "" Then
            lookupCount = lookupCount + 1
            If priceMap.Exists(lookupKey) Then
                ws.Cells(r, PURCHASE_NOTICE_AUTO_PRICE_COL).value = priceMap(lookupKey)
                matched = matched + 1
            Else
                unmatched = unmatched + 1
            End If
        End If
    Next r

    ' O列へ N列(参照単価) * E列(数量) の数式を設定
    ws.Range(ws.Cells(2, PURCHASE_NOTICE_AUTO_AMOUNT_COL), _
             ws.Cells(lastRow, PURCHASE_NOTICE_AUTO_AMOUNT_COL)).FormulaR1C1 = _
        "=RC[" & (PURCHASE_NOTICE_AUTO_PRICE_COL - PURCHASE_NOTICE_AUTO_AMOUNT_COL) & _
        "]*RC[" & (PURCHASE_NOTICE_QTY_COL - PURCHASE_NOTICE_AUTO_AMOUNT_COL) & "]"

    For r = 2 To lastRow
        WritePriceComparisonAtColumns _
            ws, r, priceSheetName, PURCHASE_NOTICE_AUTO_PRICE_COL, _
            PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_PRICE_COMPARE_COL, _
            PURCHASE_NOTICE_PRICE_GUIDANCE_COL, False
    Next r

    With ws.Range(ws.Cells(2, PURCHASE_NOTICE_AUTO_PRICE_COL), _
                  ws.Cells(lastRow, PURCHASE_NOTICE_AUTO_AMOUNT_COL))
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
    End With
    ws.Range(ws.Cells(1, PURCHASE_NOTICE_PRICE_COMPARE_COL), _
             ws.Cells(lastRow, PURCHASE_NOTICE_PRICE_COMPARE_COL)).HorizontalAlignment = xlCenter

    LogCI "購入充当単価転記: " & matched & " 件 / 照合対象=" & lookupCount & _
          " / 未一致=" & unmatched & " (単価シート=" & priceSheetName & ")"
End Sub

'==========================================================================
'  基本情報 B6(支店)/C6(出張所) -> 単価適用線区マスタ E列(名称)
'    -> 「名称_購入充当単価」シート名を返す
'==========================================================================
Private Function ResolvePurchasePriceSheetName() As String
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim BranchName As String, officeName As String
    BranchName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    officeName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    If BranchName = "" Or officeName = "" Then Exit Function

    Dim masterPath As String
    Dim connection As Object
    Set connection = OpenUnitPriceMasterAdoConnection(masterPath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(connection, PRICE_LINE_SHEET)

    Dim resultName As String
    If actualSheetName = "" Then
        LogCI "購入充当単価: マスタに「" & PRICE_LINE_SHEET & "」シートがありません"
    Else
        Dim recordset As Object
        Set recordset = CreateObject("ADODB.Recordset")
        recordset.Open "SELECT [F2], [F3], [F5] FROM " & _
                       BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

        Dim b As String, c As String, nameText As String
        Do Until recordset.EOF
            b = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(0).value))
            c = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(1).value))
            If b = BranchName And c = officeName Then
                nameText = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(2).value))
                If nameText <> "" Then
                    resultName = nameText & PURCHASE_PRICE_SHEET_SUFFIX
                    Exit Do
                End If
            End If
            recordset.MoveNext
        Loop
    End If

Cleanup:
    If Err.Number <> 0 Then
        LogCI "購入充当単価マスタADO読込エラー Err " & Err.Number & ": " & Err.Description
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    ResolvePurchasePriceSheetName = resultName
End Function

'==========================================================================
'  ADO共通ヘルパー
'==========================================================================
Private Function OpenUnitPriceMasterAdoConnection(ByRef resolvedPath As String) As Object
    Dim candidates As Collection
    Set candidates = New Collection

    AddUniqueText candidates, ResolveMasterFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim documentRoot As String
    If Len(ThisWorkbook.Path) > 0 Then
        documentRoot = fso.GetParentFolderName(ThisWorkbook.Path)
        AddUniqueText candidates, fso.BuildPath(documentRoot, _
            "単価マスタ\" & PROJECT_MASTER_FOLDER & "\出張所別_単価適用線区.xlsx")
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText candidates, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
            "線路出張所用_注文書_請求書アクセスサイト - ドキュメント\単価マスタ\" & _
            PROJECT_MASTER_FOLDER & "\出張所別_単価適用線区.xlsx"
    End If

    Dim candidate As Variant
    Dim connection As Object
    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            Set connection = CommonOpenExcelAdoConnection(CStr(candidate))
            If Not connection Is Nothing Then
                resolvedPath = CStr(candidate)
                Set OpenUnitPriceMasterAdoConnection = connection
                Exit Function
            End If
        End If
    Next candidate
End Function

Private Function FindAdoWorksheetName(ByVal connection As Object, _
                                      ByVal expectedSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim expectedKey As String
    expectedKey = CommonRemoveAllSpaces(CommonNormalizeText(expectedSheetName))

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(CommonRemoveAllSpaces(CommonNormalizeText(CStr(sheetName))), _
                   expectedKey, vbTextCompare) = 0 Then
            FindAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

Private Function BuildAdoSheetTableName(ByVal sheetName As String) As String
    BuildAdoSheetTableName = "[" & Replace$(sheetName, "]", "]]") & "$]"
End Function

Private Sub WriteReferenceValueToBasicInfo(ByVal refValue As Variant)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        LogCI "基本情報シートが見つからないため " & BASIC_INFO_REF_VALUE_CELL & " への転記をスキップ"
        Exit Sub
    End If

    With wsInfo.Range(BASIC_INFO_REF_VALUE_CELL)
        .value = refValue
        .HorizontalAlignment = xlCenter
        .Font.Name = BASIC_INFO_REF_FONT_NAME
    End With
    LogCI BASIC_INFO_REF_VALUE_CELL & " に参照シート " & REF_VALUE_SOURCE_CELL & " を転記"
End Sub
