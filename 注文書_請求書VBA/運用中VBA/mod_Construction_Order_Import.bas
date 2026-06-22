Option Explicit

Private Const DOC_ORDER As Long = 1
Private Const DOC_NOTICE As Long = 2

Private Const DATA_START_ROW As Long = 22

Private Const COL_VENDOR         As Long = 1
Private Const COL_SEIRI          As Long = COL_VENDOR + 1
Private Const COL_TYPE           As Long = COL_SEIRI + 1
Private Const COL_DAYNIGHT       As Long = COL_TYPE + 1
Private Const COL_UNIT           As Long = COL_DAYNIGHT + 1
Private Const COL_QTY            As Long = COL_UNIT + 1
Private Const COL_LINE           As Long = COL_QTY + 1
Private Const COL_MGR            As Long = COL_LINE + 1
Private Const COL_JR_PRICE       As Long = COL_MGR + 1
Private Const COL_JR_AMOUNT      As Long = COL_JR_PRICE + 1
Private Const COL_OUT_PRICE      As Long = COL_JR_AMOUNT + 1
Private Const COL_OUT_AMOUNT     As Long = COL_OUT_PRICE + 1
Private Const COL_KIND           As Long = COL_OUT_AMOUNT + 1
Private Const COL_GAP_AFTER_DATA As Long = COL_KIND + 1
Private Const COL_AUTO_PRICE     As Long = COL_GAP_AFTER_DATA + 1
Private Const COL_AUTO_AMOUNT    As Long = COL_AUTO_PRICE + 1
Private Const COL_PRICE_COMPARE  As Long = COL_AUTO_AMOUNT + 1
Private Const COL_PRICE_GUIDANCE As Long = COL_PRICE_COMPARE + 1
Private Const PRICE_GUIDANCE_COLUMN_WIDTH As Double = 59#
Private Const COL_FLAG_SIDE      As Long = COL_PRICE_GUIDANCE + 9
Private Const COL_FLAG_WELD      As Long = COL_FLAG_SIDE + 1
Private Const OUTPUT_COL_COUNT   As Long = COL_KIND
Private Const WELDING_OUTPUT_COL_OFFSET As Long = 1
Private Const WELDING_OUTPUT_COL_COUNT As Long = OUTPUT_COL_COUNT + WELDING_OUTPUT_COL_OFFSET
Private Const WELD_COL_WELDING_VENDOR As Long = 1
Private Const WELD_COL_TRACK_VENDOR As Long = 2
Private Const WELDING_VENDOR_HEADER As String = "溶接会社"
Private Const TRACK_VENDOR_HEADER As String = "軌道手元会社"
Private Const WELDING_WORK_TYPE_KEYWORD As String = "溶接工事"
Private Const TRACK_WORK_TYPE_KEYWORD As String = "軌道工事"
Private Const BASIC_INFO_VENDOR_WORK_TYPE_ROW As Long = 10

Private Const MGR_MASTER_SHEET As String = "JR管理室対応出張所"
Private Const MGR_MASTER_BRANCH_COL As Long = 2
Private Const MGR_MASTER_OFFICE_COL As Long = 3
Private Const MGR_MASTER_ROOM_COL As Long = 6
Private Const MGR_MASTER_START_ROW As Long = 2

Private Const PURCHASE_KEYWORD As String = "購入充当"
Private Const SIDELINE_KEYWORD As String = "側線"
Private Const WELDING_KEYWORD As String = "レール溶接"
Private Const CONSTRUCTION_SHEET_SUFFIX_WELDING As String = "(溶接)"
Private Const CONSTRUCTION_SHEET_SUFFIX_WORKS As String = "(工事)"
Private Const WELDING_SEIRI_OFFSET As Long = 20000

Private Const SANPAI_KEYWORD As String = "産廃処理"
Private Const SANPAI_FALLBACK_FILL_COLOR As Long = 14277081   ' RGB(217,217,217)

Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_PUBLIC_CELL As String = "B4"
Private Const BASIC_INFO_AMOUNT_CELL As String = "C22"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"
Private Const BASIC_INFO_WORKS_TOTAL_CELL As String = "C31"
Private Const BASIC_INFO_PURCHASE_TOTAL_CELL As String = "C32"
Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_TOTAL_ROW As Long = 33
Private Const BASIC_INFO_VENDOR_FIRST_COL As Long = 6
Private Const BASIC_INFO_VENDOR_STEP_COLS As Long = 3
Private Const BASIC_INFO_VENDOR_MAX_BLOCKS As Long = 20
Private Const BASIC_INFO_VENDOR_COUNT_CELL As String = "F9"
Private Const BASIC_INFO_TOTAL_NUMBER_FORMAT As String = "#,##0;[赤]-#,##0"
Private Const BASIC_INFO_SUBTOTAL_CELL As String = "C33"
Private Const BASIC_INFO_TAX_CELL As String = "C34"
Private Const BASIC_INFO_GRAND_TOTAL_CELL As String = "C35"
Private Const BASIC_INFO_TAX_LABEL_CELL As String = "B34"
Private Const BASIC_INFO_TAX_RATE_DEFAULT As Double = 0.1
Private Const BASIC_INFO_YEN_TOTAL_RANGE As String = "C31:C35"
Private Const PRICE_GUIDANCE_AMOUNT_TYPE_MESSAGE As String = _
    "基本情報シートのC22セル：単価適用区分(年初単価or設計変更単価)を確認して下さい。"

Private Const REF_VALUE_SOURCE_CELL_NOTICE As String = "H9"
Private Const REF_VALUE_SOURCE_CELL_ORDER As String = "F9"
Private Const BASIC_INFO_REF_VALUE_CELL As String = "C12"
Private Const BASIC_INFO_REF_FONT_NAME As String = "BIZ UDゴシック"

Private Const PROJECT_MASTER_START_ROW As Long = 2
Private Const PROJECT_MASTER_UNIT_PRICE_LINE_COL As Long = 6
Private Const PROJECT_MASTER_SOURCE_LINE_COL As Long = 7
Private Const MASTER_DATA_FOLDER As String = "マスタデータ"
Private Const UNIT_PRICE_LINE_MASTER_FILE As String = "出張所別_単価適用線区.xlsx"
Private Const VENDOR_MASTER_FILE As String = "業者マスタ(全社版).xlsx"
Private Const VENDOR_MASTER_ABBREV_COL As Long = 1     ' A列 業者名(略称) … 帳票・シートヘッダー表記
Private Const VENDOR_MASTER_OFFICIAL_COL As Long = 2   ' B列 請求者氏名(正規名) … 基本情報F11表記
Private Const UNIT_PRICE_DATA_START_ROW As Long = 7

Private Const PRICE_LINE_SHEET As String = "単価適用線区"
Private Const PRICE_LINE_BRANCH_COL As Long = 2
Private Const PRICE_LINE_OFFICE_COL As Long = 3
Private Const PRICE_LINE_NAME_COL As Long = 5
Private Const PRICE_LINE_START_ROW As Long = 2
Private Const PURCHASE_PRICE_SHEET_SUFFIX As String = "_購入充当単価"

Private mVendorAliasMapCache As Object
Private Const WELDING_PRICE_SHEET_SUFFIX As String = "_レール溶接単価"
Private Const PURCHASE_ORDER_SHEET_NAME As String = "購入充当指示"
Private Const PURCHASE_NOTICE_SHEET_NAME As String = "購入充当通知"
Private Const PURCHASE_PRICE_KEY_COL As Long = 1
Private Const PURCHASE_PRICE_VALUE_COL As Long = 6
Private Const PURCHASE_PRICE_DATA_START_ROW As Long = 2
Private Const PURCHASE_NOTICE_SEIRI_COL As Long = 1
Private Const PURCHASE_PRICE_LOOKUP_COL As Long = PURCHASE_NOTICE_SEIRI_COL
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

Private Const SUBCON_PRICE_FIRST_COL As Long = 11
Private Const WELDING_SUBCON_PRICE_FIRST_COL As Long = SUBCON_PRICE_FIRST_COL + WELDING_OUTPUT_COL_OFFSET
Private Const WELDING_PRICE_SEIRI_COL As Long = 2
Private Const WELDING_PRICE_DATA_START_ROW As Long = 7
Private Const WELDING_PRICE_WELDING_DAY_COL As Long = 7
Private Const WELDING_PRICE_FIRST_RAIL_DAY_COL As Long = 9
Private Const WELDING_PRICE_VENDOR_NAME_ROW As Long = 5
Private Const UNIT_PRICE_VENDOR_NAME_ROW As Long = 5
Private Const UNIT_PRICE_VENDOR_FIRST_DAY_COL As Long = 7
Private Const UNIT_PRICE_WORK_TYPE_COL As Long = 3
Private Const UNIT_PRICE_REF_DAY_COL As Long = 5
Private Const UNIT_PRICE_REF_NIGHT_COL As Long = 6

Private Const SOURCE_SHEET_NAME_CELL As String = "A3"

Private mSuppressOverwritePrompt As Boolean
Private mLastCreatedImportSheet As Worksheet
Private mSanpaiFillColorCached As Boolean
Private mSanpaiFillColorCache As Long

Public Sub ImportConstructionDocument()
    Dim scrn As Boolean, evt As Boolean, alerts As Boolean

    scrn = Application.screenUpdating
    evt = Application.EnableEvents
    alerts = Application.DisplayAlerts

    mSuppressOverwritePrompt = False
    Set mLastCreatedImportSheet = Nothing

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

    RefreshBasicInfoConstructionTotals

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

Private Sub RestoreAutomaticCalculation()
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
End Sub

Private Sub ImportOneConstructionDocument(ByVal srcPath As String, _
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
    baseSheetName = SanitizeSheetName(sourceA3Text)
    If baseSheetName = "" Then
        MsgBox "取込ブックの " & SOURCE_SHEET_NAME_CELL & " が空のため、シート名を決定できません。", vbExclamation
        GoTo Cleanup
    End If
    LogCI "baseSheetName=[" & baseSheetName & "]"

    Dim guidanceDocumentName As String
    guidanceDocumentName = ResolveGuidanceDocumentName(sourceA3Text, docType)
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
                    weldArr(WeldingRowArrayIndex(COL_VENDOR)) = ""
                    weldArr(WELD_COL_TRACK_VENDOR - 1) = ""
                    weldArr(WeldingRowArrayIndex(COL_SEIRI)) = vSeiri(i)
                    weldArr(WeldingRowArrayIndex(COL_TYPE)) = vType(i)
                    weldArr(WeldingRowArrayIndex(COL_DAYNIGHT)) = vDN(i)
                    weldArr(WeldingRowArrayIndex(COL_UNIT)) = vUnit(i)
                    weldArr(WeldingRowArrayIndex(COL_QTY)) = vQty(i)
                    weldArr(WeldingRowArrayIndex(COL_LINE)) = vLine(i)
                    weldArr(WeldingRowArrayIndex(COL_MGR)) = mgrOut
                    weldArr(WeldingRowArrayIndex(COL_JR_PRICE)) = vPrice(i)
                    weldArr(WeldingRowArrayIndex(COL_JR_AMOUNT)) = vAmount(i)
                    weldArr(WeldingRowArrayIndex(COL_OUT_PRICE)) = ""
                    weldArr(WeldingRowArrayIndex(COL_OUT_AMOUNT)) = ""
                    weldArr(WeldingRowArrayIndex(COL_KIND)) = vKind(i)
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
               "■基本情報(支店/出張所)に対応する管理室:" & vbCrLf & JoinKeys(mgrSet) & vbCrLf & vbCrLf & _
               "■取込データに含まれる管理室:" & vbCrLf & JoinKeys(seenMgr) & vbCrLf & vbCrLf & _
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
            BuildConstructionSheetName(sourceA3Text, False), _
            worksRows, docType, guidanceDocumentName, False)
        If wsActivate Is Nothing Then Set wsActivate = wsWorks
    End If

    If weldRows.Count > 0 Then
        Set wsWeld = BuildConstructionOutputSheet( _
            BuildConstructionSheetName(sourceA3Text, True), _
            weldRows, docType, guidanceDocumentName, True)
        If wsActivate Is Nothing Then Set wsActivate = wsWeld
    End If

    If purchRows.Count > 0 Then
        Dim purchName As String
        If docType = DOC_ORDER Then
            purchName = PURCHASE_ORDER_SHEET_NAME
        Else
            purchName = PURCHASE_NOTICE_SHEET_NAME
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

Private Function BuildConstructionOutputSheet(ByVal sheetName As String, _
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
            ws, OutputSheetCol(ws, COL_AUTO_PRICE), OutputSheetCol(ws, COL_AUTO_AMOUNT), _
            OutputSheetCol(ws, COL_PRICE_COMPARE), True
    Else
        WriteAdditionalHeaders ws
    End If
    FillReferenceUnitPrices ws, guidanceDocumentName, isWelding
    If isWelding Then
        FormatSheet ws, OutputSheetCol(ws, COL_SEIRI)
        ApplyPriceGuidanceColumnLayoutAtColumns _
            ws, OutputSheetCol(ws, COL_PRICE_COMPARE), OutputSheetCol(ws, COL_PRICE_GUIDANCE)
        mod_subcontractorselector.ApplySubcontractorDropdowns ws
    Else
        FormatSheet ws
        ApplyPriceGuidanceColumnLayout ws
        mod_subcontractorselector.ApplySubcontractorDropdowns ws
    End If
    ApplySanpaiRowRestrictions ws
    RefreshOutputSheetVendorColumnColors ws, GetLastDataRow(ws)

    If Not isWelding Then
        ws.Columns(COL_VENDOR).HorizontalAlignment = xlCenter
    End If

    If docType = DOC_NOTICE Then ws.Columns(OutputSheetCol(ws, COL_MGR)).Hidden = True

    ws.Range(ws.Cells(1, OutputSheetCol(ws, COL_OUT_PRICE)), _
             ws.Cells(1, OutputSheetCol(ws, COL_OUT_AMOUNT))).EntireColumn.Delete Shift:=xlToLeft

    WriteJrTotalRow ws
    RedrawOutputSheetDataBorders ws

    If isWelding Then
        ApplyWeldingOutputSheetColumnAlignment ws
    End If

    If isWelding Then
        ApplyOutputSheetHeaderAutoFilter ws, _
            OutputSheetCol(ws, COL_SEIRI), OutputSheetCol(ws, COL_KIND), _
            OutputSheetCol(ws, COL_AUTO_AMOUNT), OutputSheetCol(ws, COL_PRICE_COMPARE), _
            OutputSheetCol(ws, COL_PRICE_GUIDANCE)
    Else
        ApplyOutputSheetHeaderAutoFilter ws
    End If

    ApplyOutputSheetHeaderFreezePanes ws

    Set BuildConstructionOutputSheet = ws
End Function

Private Sub PrepareExistingOutputSheets(ByRef cancelImport As Boolean)
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

Private Sub DeleteSheetByName(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.worksheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
    End If
    On Error GoTo 0
End Sub

Private Function CollectExistingManagedOutputSheetNames() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsManagedImportOutputSheet(ws) Then result.Add ws.Name
    Next ws

    Set CollectExistingManagedOutputSheetNames = result
End Function

Private Function IsManagedImportOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    If IsPurchaseOutputSheet(ws) Then
        IsManagedImportOutputSheet = True
        Exit Function
    End If

    If Not SheetNameHasConstructionOutputSuffix(ws.Name) Then Exit Function

    If IsConstructionOutputSheet(ws) Then
        IsManagedImportOutputSheet = True
        Exit Function
    End If

    IsManagedImportOutputSheet = _
        (FindHeaderColumn(ws, "整理番号") > 0) And _
        (FindHeaderColumn(ws, "JR金額") > 0) And _
        (FindHeaderColumn(ws, "工種分類") > 0)
End Function

Private Function SheetNameHasConstructionOutputSuffix(ByVal sheetName As String) As Boolean
    SheetNameHasConstructionOutputSuffix = _
        SheetNameEndsWithSuffixText(sheetName, CONSTRUCTION_SHEET_SUFFIX_WORKS) Or _
        SheetNameEndsWithSuffixText(sheetName, CONSTRUCTION_SHEET_SUFFIX_WELDING)
End Function

Private Function SheetNameEndsWithSuffixText(ByVal sheetName As String, ByVal suffixText As String) As Boolean
    Dim normalizedName As String
    Dim normalizedSuffix As String
    normalizedName = NormalizeSheetNameParentheses(sheetName)
    normalizedSuffix = NormalizeSheetNameParentheses(suffixText)

    If Len(normalizedName) < Len(normalizedSuffix) Then Exit Function
    SheetNameEndsWithSuffixText = _
        (StrComp(Right$(normalizedName, Len(normalizedSuffix)), normalizedSuffix, vbTextCompare) = 0)
End Function

Private Function NormalizeSheetNameParentheses(ByVal text As String) As String
    Dim t As String
    t = text
    t = Replace$(t, ChrW$(&HFF08), "(")
    t = Replace$(t, ChrW$(&HFF09), ")")
    NormalizeSheetNameParentheses = t
End Function

Private Function PickSourceFiles() As Collection
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

Private Function ValidateBasicInfoForImport(ByVal docType As Long) As Boolean
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

Private Function ResolveMasterFilePath() As String
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

Private Function ReadColumnValues(ByVal ws As Worksheet, ByVal col As String, _
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

Private Function GetImportSheetAnchorSheet() As Worksheet
    On Error Resume Next
    Set GetImportSheetAnchorSheet = Sheet1
    On Error GoTo 0

    If GetImportSheetAnchorSheet Is Nothing Then
        Set GetImportSheetAnchorSheet = CommonGetBasicInfoWorksheet(ThisWorkbook)
    End If
End Function

Private Function FindRightmostManagedImportSheetAfterAnchor( _
        ByVal anchorSheet As Worksheet) As Worksheet
    If anchorSheet Is Nothing Then Exit Function

    Dim ws As Worksheet
    Dim candidate As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Index > anchorSheet.Index Then
            If IsManagedImportOutputSheet(ws) Then
                If Not IsPurchaseOutputSheet(ws) Then
                    If candidate Is Nothing Or ws.Index > candidate.Index Then
                        Set candidate = ws
                    End If
                End If
            End If
        End If
    Next ws

    Set FindRightmostManagedImportSheetAfterAnchor = candidate
End Function

Private Function ResolveImportSheetInsertAfter( _
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

Private Function CreateOrReplaceSheet(ByVal sheetName As String) As Worksheet
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

Private Sub NormalizeManagedImportSheetOrder()
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

Private Sub AppendOutputSheetsBySuffix(ByVal target As Collection, ByVal suffixText As String)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsManagedImportOutputSheet(ws) Then
            If SheetNameEndsWithSuffixText(ws.Name, suffixText) Then
                target.Add ws
            End If
        End If
    Next ws
End Sub

Private Sub AppendPurchaseOutputSheets(ByVal target As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsPurchaseOutputSheet(ws) Then
            target.Add ws
        End If
    Next ws
End Sub

Private Sub WriteRecordsToSheet(ByVal ws As Worksheet, ByVal rows As Collection)
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

Private Sub WriteWeldingRecordsToSheet(ByVal ws As Worksheet, ByVal rows As Collection)
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

Private Sub ApplyPurchaseNoticeLayout(ByVal ws As Worksheet)
    ws.Columns(COL_VENDOR).Delete Shift:=xlToLeft
    ws.Cells(1, PURCHASE_NOTICE_SEIRI_COL).value = "整理番号"
End Sub

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

Public Sub RefreshSubcontractorPriceColumns(ByVal ws As Worksheet, _
                                            Optional ByVal changedRows As Collection = Nothing)
    If ws Is Nothing Then Exit Sub
    If Not IsConstructionVendorOutputSheet(ws) Then Exit Sub
    If FindHeaderColumn(ws, "整理番号") = 0 Then Exit Sub

    Dim scrn As Boolean
    Dim calcMode As XlCalculation
    Dim evt As Boolean
    scrn = Application.screenUpdating
    calcMode = Application.Calculation
    evt = Application.EnableEvents
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim refreshErrNo As Long
    Dim refreshErrDesc As String
    Dim refreshStep As String
    Dim coreErrNo As Long
    Dim coreErrDesc As String
    Dim matchedCount As Long
    matchedCount = 0

    Dim lastRow As Long
    Dim vendorNames As Collection
    Dim kindColumn As Long
    Dim subconFirstCol As Long
    Dim layoutMatches As Boolean
    Dim insertedColumnCount As Long
    Dim vendorColumnMap As Object
    Dim lineSheetMap As Object
    Dim vendorPriceCaches As Object
    Dim seiriColumn As Long
    Dim dayNightColumn As Long
    Dim lineColumn As Long
    Dim qtyColumn As Long
    Dim isWeldingSheet As Boolean
    Dim vendorColumns As Collection
    Dim partialUpdate As Boolean

    Set vendorNames = New Collection

    refreshStep = "CollectVendors"
    On Error Resume Next
    lastRow = GetLastDataRow(ws)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Set vendorNames = CollectSelectedSubcontractors(ws, lastRow)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Err.Clear

    refreshStep = "FindKindColumn"
    kindColumn = FindHeaderColumn(ws, "工種分類")
    If Err.Number <> 0 Then GoTo RefreshSetupError
    If kindColumn = 0 Then GoTo RefreshExit
    Err.Clear

    refreshStep = "ResolveSubconFirstCol"
    subconFirstCol = OutputSheetSubconPriceFirstCol(ws)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Err.Clear

    refreshStep = "LayoutCheck"
    layoutMatches = SubconColumnLayoutMatches(ws, vendorNames, subconFirstCol, kindColumn)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Err.Clear

    If Not layoutMatches Then
        refreshStep = "LayoutDelete"
        If kindColumn > subconFirstCol Then
            ws.Range(ws.Columns(subconFirstCol), _
                     ws.Columns(kindColumn - 1)).Delete Shift:=xlToLeft
            If Err.Number <> 0 Then GoTo RefreshSetupError
            kindColumn = FindHeaderColumn(ws, "工種分類")
            If Err.Number <> 0 Then GoTo RefreshSetupError
            If kindColumn = 0 Then GoTo RefreshExit
        End If
        Err.Clear
    End If
    On Error GoTo 0

    If vendorNames.Count = 0 Or lastRow < 2 Then
        lastRow = GetLastDataRow(ws)
        If lastRow >= 2 Then
            Dim emptyVendors As New Collection
            WriteOutputTotalRows ws, emptyVendors, 0, 0
            RefreshOutputSheetVendorColumnColors ws, lastRow
        End If
        RefreshBasicInfoConstructionTotals
        GoTo RefreshExit
    End If

    insertedColumnCount = vendorNames.Count * 2

    If Not layoutMatches Then
        refreshStep = "LayoutInsert"
        On Error Resume Next
        ws.Range(ws.Columns(subconFirstCol), _
                 ws.Columns(subconFirstCol + insertedColumnCount - 1)).Insert Shift:=xlToRight
        If Err.Number <> 0 Then GoTo RefreshSetupError

        lastRow = GetLastDataRow(ws)
        If Err.Number <> 0 Then GoTo RefreshSetupError
        If lastRow < 2 Then
            On Error GoTo 0
            RefreshBasicInfoConstructionTotals
            GoTo RefreshExit
        End If

        Dim vendorIndex As Long
        For vendorIndex = 1 To vendorNames.Count
            Dim vendorName As String
            Dim priceColumn As Long
            vendorName = CStr(vendorNames(vendorIndex))
            priceColumn = subconFirstCol + ((vendorIndex - 1) * 2)
            ws.Cells(1, priceColumn).value = vendorName & "単価"
            If Err.Number <> 0 Then GoTo RefreshSetupError
            ws.Cells(1, priceColumn + 1).value = vendorName & "金額"
            If Err.Number <> 0 Then GoTo RefreshSetupError
        Next vendorIndex

        kindColumn = FindHeaderColumn(ws, "工種分類")
        If Err.Number <> 0 Then GoTo RefreshSetupError
        Err.Clear
        On Error GoTo 0
    End If

    If kindColumn > subconFirstCol Then
        insertedColumnCount = kindColumn - subconFirstCol
    End If

    refreshStep = "BuildVendorColumnMap"
    On Error Resume Next
    Set vendorColumnMap = BuildVendorPriceColumnMap(vendorNames, subconFirstCol)
    If Err.Number <> 0 Then GoTo RefreshSetupError

    refreshStep = "BuildLineMap"
    Set lineSheetMap = BuildConstructionLineSheetMap()
    Err.Clear
    If lineSheetMap Is Nothing Then
        Set lineSheetMap = CreateObject("Scripting.Dictionary")
        lineSheetMap.CompareMode = vbTextCompare
    End If

    Set vendorPriceCaches = CreateObject("Scripting.Dictionary")
    vendorPriceCaches.CompareMode = vbTextCompare

    refreshStep = "PrepareApply"
    seiriColumn = OutputSheetSeiriColumn(ws)
    dayNightColumn = OutputSheetCol(ws, COL_DAYNIGHT)
    lineColumn = OutputSheetCol(ws, COL_LINE)
    qtyColumn = OutputSheetCol(ws, COL_QTY)
    isWeldingSheet = IsWeldingOutputSheet(ws)
    Set vendorColumns = OutputSheetVendorColumns(ws)
    partialUpdate = False
    If Not changedRows Is Nothing Then partialUpdate = (changedRows.Count > 0)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Err.Clear
    On Error GoTo 0

    '  以降は段階毎にエラーを捕捉し、どの段階で失敗したかを refreshStep として
    '  LogCI に記録する。単価適用以外の装飾・合計行描画は、失敗しても処理全体を
    '  中断させず、エラーダイアログも出さない。単価適用が一件も成功しなかった
    '  場合のみ利用者へ通知する。

    '  (1) 単価適用(中核処理)
    refreshStep = "ApplyPrices"
    On Error Resume Next
    If partialUpdate And layoutMatches Then
        ApplySubcontractorPricesPartial ws, changedRows, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
    Else
        ApplySubcontractorPricesBatch ws, lastRow, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
    End If
    If Err.Number <> 0 Then
        coreErrNo = Err.Number
        coreErrDesc = Err.Description
        LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
        matchedCount = 0
        lastRow = GetLastDataRow(ws)
        ApplySubcontractorPricesBatch ws, lastRow, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
        If Err.Number <> 0 Then
            If coreErrNo = 0 Then
                coreErrNo = Err.Number
                coreErrDesc = Err.Description
            End If
            LogCI "RefreshSubcontractorPriceColumns step=ApplyPricesRetry Err " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
    ElseIf partialUpdate And layoutMatches And matchedCount = 0 Then
        lastRow = GetLastDataRow(ws)
        ApplySubcontractorPricesBatch ws, lastRow, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
    End If
    On Error GoTo 0

    '  (2) 金額数式
    refreshStep = "AmountFormulas"
    On Error Resume Next
    lastRow = GetLastDataRow(ws)
    ApplySubcontractorAmountFormulas ws, lastRow, vendorColumnMap, qtyColumn
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (3) 列内塗りつぶし(単価有無の背景色)
    refreshStep = "ColumnInteriors"
    On Error Resume Next
    RefreshSubcontractorColumnInteriors ws, lastRow, subconFirstCol, insertedColumnCount
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (3b) 施工会社列の業者情報色
    refreshStep = "VendorColumnColors"
    On Error Resume Next
    lastRow = GetLastDataRow(ws)
    RefreshOutputSheetVendorColumnColors ws, lastRow
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (4) 列書式設定(レイアウト変更時のみ)
    If Not layoutMatches Then
        refreshStep = "FormatColumns"
        On Error Resume Next
        FormatSubcontractorPriceColumns ws, lastRow, insertedColumnCount, subconFirstCol
        RedrawOutputSheetDataBorders ws
        If Err.Number <> 0 Then
            LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    '  (5) 合計行
    refreshStep = "WriteTotals"
    On Error Resume Next
    WriteOutputTotalRows ws, vendorNames, subconFirstCol, insertedColumnCount
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumns step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (6) 基本情報合計(内部に独自ハンドラあり)以降は再度 RefreshError で保護
    On Error GoTo RefreshError
    refreshStep = "BasicInfoTotals"
    RefreshBasicInfoConstructionTotals

    refreshStep = "Log"
    LogCI "施工会社別単価列: 会社数=" & vendorNames.Count & _
          " / 単価一致=" & matchedCount & _
          IIf(partialUpdate And layoutMatches, " / 部分更新", "")
    GoTo RefreshExit

RefreshSetupError:
    refreshErrNo = Err.Number
    refreshErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    LogCI "RefreshSubcontractorPriceColumns Setup Err " & refreshErrNo & ": " & _
          refreshErrDesc & " (step=" & refreshStep & ")"
    GoTo RefreshExit

RefreshError:
    refreshErrNo = Err.Number
    refreshErrDesc = Err.Description

RefreshExit:
    Application.screenUpdating = scrn
    Application.Calculation = calcMode
    If calcMode = xlCalculationAutomatic Then
        On Error Resume Next
        ws.Calculate
        On Error GoTo 0
    End If
    Application.EnableEvents = evt
    If refreshErrNo <> 0 Then
        '  Setup段階や終盤など、致命的なエラー時のみダイアログ表示
        MsgBox "施工会社別の単価・金額列を更新できませんでした。" & vbCrLf & _
               refreshErrDesc, vbExclamation
        LogCI "RefreshSubcontractorPriceColumns Err " & refreshErrNo & ": " & refreshErrDesc & " (step=" & refreshStep & ")"
    ElseIf coreErrNo <> 0 And matchedCount = 0 Then
        '  単価適用が一件も成功しなかった場合のみ通知
        MsgBox "施工会社別の単価・金額列を更新できませんでした。" & vbCrLf & _
               coreErrDesc, vbExclamation
        LogCI "RefreshSubcontractorPriceColumns core Err " & coreErrNo & ": " & coreErrDesc & " (step=ApplyPrices)"
    ElseIf matchedCount = 0 Then
        If Not vendorNames Is Nothing Then
            LogCI "施工会社別単価: 一致0件 (会社数=" & vendorNames.Count & ")"
        Else
            LogCI "施工会社別単価: 一致0件"
        End If
    End If
End Sub

Private Function BuildVendorPriceColumnMap(ByVal vendorNames As Collection, _
                                           ByVal subconFirstCol As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    If vendorNames Is Nothing Then
        Set BuildVendorPriceColumnMap = result
        Exit Function
    End If

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorNames.Count
        Dim vendorKey As String
        vendorKey = NormalizeVendorPriceName(CStr(vendorNames(vendorIndex)))
        If vendorKey <> "" Then
            If Not result.Exists(vendorKey) Then
                result.Add vendorKey, subconFirstCol + ((vendorIndex - 1) * 2)
            End If
        End If
    Next vendorIndex

    Set BuildVendorPriceColumnMap = result
End Function

Private Function CollectExistingSubconColumnVendors(ByVal ws As Worksheet, _
                                                    ByVal subconFirstCol As Long, _
                                                    ByVal kindColumn As Long) As Collection
    Dim result As New Collection
    If kindColumn <= subconFirstCol Then
        Set CollectExistingSubconColumnVendors = result
        Exit Function
    End If

    Dim colIndex As Long
    For colIndex = subconFirstCol To kindColumn - 1 Step 2
        Dim headerText As String
        headerText = Trim$(CommonNzText(ws.Cells(1, colIndex).value))
        If Len(headerText) >= 2 And Right$(headerText, 2) = "単価" Then
            result.Add Left$(headerText, Len(headerText) - 2)
        End If
    Next colIndex

    Set CollectExistingSubconColumnVendors = result
End Function

Private Function SubconColumnLayoutMatches(ByVal ws As Worksheet, _
                                           ByVal vendorNames As Collection, _
                                           ByVal subconFirstCol As Long, _
                                           ByVal kindColumn As Long) As Boolean
    Dim existingNames As Collection
    Set existingNames = CollectExistingSubconColumnVendors(ws, subconFirstCol, kindColumn)
    If existingNames.Count <> vendorNames.Count Then Exit Function

    Dim i As Long
    For i = 1 To vendorNames.Count
        If NormalizeVendorPriceName(CStr(existingNames(i))) <> _
           NormalizeVendorPriceName(CStr(vendorNames(i))) Then Exit Function
    Next i

    SubconColumnLayoutMatches = True
End Function

Private Function LookupSubcontractorVendorPrice( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal vendorColumn As Long, _
    ByVal lineSheetMap As Object, _
    ByVal vendorPriceCaches As Object, _
    ByVal seiriColumn As Long, _
    ByVal dayNightColumn As Long, _
    ByVal lineColumn As Long, _
    ByVal isWeldingSheet As Boolean) As Variant

    LookupSubcontractorVendorPrice = Empty
    If IsSanpaiRow(ws, rowIndex) Then Exit Function

    Dim recordKey As String
    If isWeldingSheet Then
        recordKey = BuildWeldingLookupKey(ws.Cells(rowIndex, seiriColumn).value)
    Else
        recordKey = NormalizeRecordKey(ws.Cells(rowIndex, seiriColumn).value)
    End If
    If recordKey = "" Then Exit Function

    Dim dayNightText As String
    dayNightText = CommonNzText(ws.Cells(rowIndex, dayNightColumn).value)

    If isWeldingSheet Then
        Dim weldingPriceSheetName As String
        weldingPriceSheetName = ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then Exit Function
        LookupSubcontractorVendorPrice = LookupWeldingOutputVendorPrice( _
            weldingPriceSheetName, vendorPriceCaches, recordKey, _
            CommonNzText(ws.Cells(rowIndex, vendorColumn).value), _
            (vendorColumn = WELD_COL_WELDING_VENDOR), dayNightText)
    Else
        Dim unitPriceSheetName As String
        unitPriceSheetName = ResolveUnitPriceSheetName( _
            lineSheetMap, CommonNzText(ws.Cells(rowIndex, lineColumn).value))
        If unitPriceSheetName = "" Then Exit Function

        Dim vendorPriceRows As Object
        Set vendorPriceRows = GetVendorUnitPriceRows( _
            unitPriceSheetName, CommonNzText(ws.Cells(rowIndex, vendorColumn).value), vendorPriceCaches)
        If vendorPriceRows Is Nothing Then Exit Function
        If Not vendorPriceRows.Exists(recordKey) Then
            LogCI "施工会社単価: 整理番号未一致 sheet=[" & unitPriceSheetName & "] key=[" & recordKey & "]"
            Exit Function
        End If

        Dim dayNightPrices As Variant
        dayNightPrices = vendorPriceRows(recordKey)
        LookupSubcontractorVendorPrice = SelectUsableDayNightPrice(dayNightText, dayNightPrices)
    End If
End Function

Private Sub ApplySubcontractorPricesPartial( _
    ByVal ws As Worksheet, _
    ByVal changedRows As Collection, _
    ByVal vendorColumnMap As Object, _
    ByVal vendorColumns As Collection, _
    ByVal lineSheetMap As Object, _
    ByVal vendorPriceCaches As Object, _
    ByVal seiriColumn As Long, _
    ByVal dayNightColumn As Long, _
    ByVal lineColumn As Long, _
    ByVal qtyColumn As Long, _
    ByVal isWeldingSheet As Boolean, _
    ByRef matchedCount As Long)

    If vendorColumnMap Is Nothing Then Exit Sub
    If vendorColumns Is Nothing Then Exit Sub
    If changedRows Is Nothing Then Exit Sub
    If vendorPriceCaches Is Nothing Then Exit Sub

    Dim vendorKey As Variant
    For Each vendorKey In vendorColumnMap.Keys
        Dim priceColumn As Long
        priceColumn = CLng(vendorColumnMap(vendorKey))

        Dim rowRef As Variant
        For Each rowRef In changedRows
            Dim rowIndex As Long
            rowIndex = CLng(rowRef)

            Dim vendorCol As Variant
            Dim rowVendorKey As String
            Dim appliesToRow As Boolean
            appliesToRow = False
            For Each vendorCol In vendorColumns
                rowVendorKey = NormalizeVendorPriceName( _
                    CommonNzText(ws.Cells(rowIndex, CLng(vendorCol)).value))
                If rowVendorKey = CStr(vendorKey) Then
                    appliesToRow = True
                    Dim vendorPrice As Variant
                    vendorPrice = LookupSubcontractorVendorPrice( _
                        ws, rowIndex, CLng(vendorCol), lineSheetMap, vendorPriceCaches, _
                        seiriColumn, dayNightColumn, lineColumn, isWeldingSheet)
                    If Not IsEmpty(vendorPrice) And Not IsError(vendorPrice) Then
                        ws.Cells(rowIndex, priceColumn).value = vendorPrice
                        matchedCount = matchedCount + 1
                    Else
                        ws.Cells(rowIndex, priceColumn).ClearContents
                    End If
                    Exit For
                End If
            Next vendorCol

            If Not appliesToRow Then
                ws.Cells(rowIndex, priceColumn).ClearContents
            End If
        Next rowRef
    Next vendorKey
End Sub

Private Sub ApplySubcontractorPricesBatch( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal vendorColumnMap As Object, _
    ByVal vendorColumns As Collection, _
    ByVal lineSheetMap As Object, _
    ByVal vendorPriceCaches As Object, _
    ByVal seiriColumn As Long, _
    ByVal dayNightColumn As Long, _
    ByVal lineColumn As Long, _
    ByVal qtyColumn As Long, _
    ByVal isWeldingSheet As Boolean, _
    ByRef matchedCount As Long)

    If lastRow < 2 Then Exit Sub
    If vendorColumnMap Is Nothing Then Exit Sub
    If vendorColumns Is Nothing Then Exit Sub
    If vendorPriceCaches Is Nothing Then Exit Sub

    Dim rowCount As Long
    rowCount = lastRow - 1

    Dim seiriData As Variant
    Dim dayNightData As Variant
    Dim lineData As Variant
    Dim typeData As Variant
    Dim typeColumn As Long
    typeColumn = OutputSheetCol(ws, COL_TYPE)
    seiriData = ws.Range(ws.Cells(2, seiriColumn), ws.Cells(lastRow, seiriColumn)).Value2
    dayNightData = ws.Range(ws.Cells(2, dayNightColumn), ws.Cells(lastRow, dayNightColumn)).Value2
    lineData = ws.Range(ws.Cells(2, lineColumn), ws.Cells(lastRow, lineColumn)).Value2
    typeData = ws.Range(ws.Cells(2, typeColumn), ws.Cells(lastRow, typeColumn)).Value2

    Dim vendorColData() As Variant
    ReDim vendorColData(1 To vendorColumns.Count)
    Dim vendorColIndex As Long
    vendorColIndex = 0
    Dim vendorCol As Variant
    For Each vendorCol In vendorColumns
        vendorColIndex = vendorColIndex + 1
        vendorColData(vendorColIndex) = ws.Range( _
            ws.Cells(2, CLng(vendorCol)), ws.Cells(lastRow, CLng(vendorCol))).Value2
    Next vendorCol

    Dim amountFormula As String
    Dim rowVendorKey As String
    Dim vendorPrice As Variant

    Dim vendorKey As Variant
    For Each vendorKey In vendorColumnMap.Keys
        Dim priceColumn As Long
        priceColumn = CLng(vendorColumnMap(vendorKey))

        Dim priceValues() As Variant
        ReDim priceValues(1 To rowCount, 1 To 1)

        Dim r As Long
        For r = 1 To rowCount
            If IsSanpaiTypeText(CommonNzText(GetArrayCellValue(typeData, r, 1))) Then GoTo NextBatchRow

            vendorColIndex = 0
            For Each vendorCol In vendorColumns
                vendorColIndex = vendorColIndex + 1
                rowVendorKey = NormalizeVendorPriceName( _
                    CommonNzText(GetArrayCellValue(vendorColData(vendorColIndex), r, 1)))
                If rowVendorKey = CStr(vendorKey) Then
                    vendorPrice = LookupSubcontractorVendorPriceFromArrays( _
                        seiriData, dayNightData, lineData, typeData, vendorColData(vendorColIndex), _
                        r, CLng(vendorCol), lineSheetMap, vendorPriceCaches, isWeldingSheet)
                    If Not IsEmpty(vendorPrice) And Not IsError(vendorPrice) Then
                        priceValues(r, 1) = vendorPrice
                        matchedCount = matchedCount + 1
                    End If
                    Exit For
                End If
            Next vendorCol
NextBatchRow:
        Next r

        ws.Range(ws.Cells(2, priceColumn), ws.Cells(lastRow, priceColumn)).Value2 = priceValues
    Next vendorKey
End Sub

Private Sub ApplySubcontractorAmountFormulas(ByVal ws As Worksheet, _
                                               ByVal lastRow As Long, _
                                               ByVal vendorColumnMap As Object, _
                                               ByVal qtyColumn As Long)
    If lastRow < 2 Then Exit Sub
    If vendorColumnMap Is Nothing Then Exit Sub
    If vendorColumnMap.Count = 0 Then Exit Sub

    Dim amountFormula As String
    amountFormula = "=IF(OR(RC[-1]="""",RC" & qtyColumn & "=""""),"""",RC[-1]*RC" & qtyColumn & ")"

    Dim vendorKey As Variant
    For Each vendorKey In vendorColumnMap.Keys
        Dim amountColumn As Long
        amountColumn = CLng(vendorColumnMap(vendorKey)) + 1
        ws.Range(ws.Cells(2, amountColumn), ws.Cells(lastRow, amountColumn)).FormulaR1C1 = amountFormula
    Next vendorKey
End Sub

Private Sub RefreshSubcontractorColumnInteriors(ByVal ws As Worksheet, _
                                                  ByVal lastRow As Long, _
                                                  ByVal firstColumn As Long, _
                                                  ByVal columnCount As Long)
    If lastRow < 2 Or columnCount <= 0 Then Exit Sub

    Dim lastColumn As Long
    Dim sanpaiFillColor As Long
    Dim r As Long
    Dim priceCol As Long
    lastColumn = firstColumn + columnCount - 1
    sanpaiFillColor = GetSanpaiFillColor()

    For r = 2 To lastRow
        If IsSanpaiRow(ws, r) Then
            ws.Range(ws.Cells(r, firstColumn), ws.Cells(r, lastColumn)).Interior.Color = sanpaiFillColor
        Else
            For priceCol = firstColumn To lastColumn Step 2
                If Len(CommonNzText(ws.Cells(r, priceCol).value)) > 0 Then
                    ws.Cells(r, priceCol).Interior.Pattern = xlNone
                    ws.Cells(r, priceCol + 1).Interior.Pattern = xlNone
                Else
                    ws.Cells(r, priceCol).Interior.Color = sanpaiFillColor
                    ws.Cells(r, priceCol + 1).Interior.Color = sanpaiFillColor
                End If
            Next priceCol
        End If
    Next r
End Sub

Private Sub RefreshOutputSheetVendorColumnColors(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws Is Nothing Then Exit Sub
    If lastRow < 2 Then Exit Sub
    If Not IsConstructionVendorOutputSheet(ws) Then Exit Sub

    Dim vendorColumns As Collection
    Set vendorColumns = OutputSheetVendorColumns(ws)

    Dim r As Long
    Dim vendorCol As Variant
    For r = 2 To lastRow
        If IsSanpaiRow(ws, r) Then GoTo NextColorRow

        For Each vendorCol In vendorColumns
            mod_VendorInfoColors.ApplyOutputSheetVendorCellColor _
                ws, r, CLng(vendorCol), ResolveVendorColumnWorkTypeKeyword(ws, CLng(vendorCol))
        Next vendorCol
NextColorRow:
    Next r
End Sub

Private Function ResolveVendorColumnWorkTypeKeyword(ByVal ws As Worksheet, _
                                                    ByVal vendorColumn As Long) As String
    If Not IsWeldingOutputSheet(ws) Then Exit Function

    If vendorColumn = WELD_COL_WELDING_VENDOR Then
        ResolveVendorColumnWorkTypeKeyword = WELDING_WORK_TYPE_KEYWORD
    ElseIf vendorColumn = WELD_COL_TRACK_VENDOR Then
        ResolveVendorColumnWorkTypeKeyword = TRACK_WORK_TYPE_KEYWORD
    End If
End Function

Private Function GetArrayCellValue(ByVal arr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If IsArray(arr) Then
        GetArrayCellValue = arr(rowIndex, colIndex)
    Else
        If rowIndex = 1 And colIndex = 1 Then GetArrayCellValue = arr
    End If
End Function

Private Function LookupSubcontractorVendorPriceFromArrays( _
    ByVal seiriData As Variant, _
    ByVal dayNightData As Variant, _
    ByVal lineData As Variant, _
    ByVal typeData As Variant, _
    ByVal vendorData As Variant, _
    ByVal rowIndex As Long, _
    ByVal vendorColumn As Long, _
    ByVal lineSheetMap As Object, _
    ByVal vendorPriceCaches As Object, _
    ByVal isWeldingSheet As Boolean) As Variant

    LookupSubcontractorVendorPriceFromArrays = Empty
    If IsSanpaiTypeText(CommonNzText(GetArrayCellValue(typeData, rowIndex, 1))) Then Exit Function

    Dim recordKey As String
    If isWeldingSheet Then
        recordKey = BuildWeldingLookupKey(GetArrayCellValue(seiriData, rowIndex, 1))
    Else
        recordKey = NormalizeRecordKey(GetArrayCellValue(seiriData, rowIndex, 1))
    End If
    If recordKey = "" Then Exit Function

    Dim dayNightText As String
    dayNightText = CommonNzText(GetArrayCellValue(dayNightData, rowIndex, 1))

    Dim vendorName As String
    vendorName = CommonNzText(GetArrayCellValue(vendorData, rowIndex, 1))

    If isWeldingSheet Then
        Dim weldingPriceSheetName As String
        weldingPriceSheetName = ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then Exit Function
        LookupSubcontractorVendorPriceFromArrays = LookupWeldingOutputVendorPrice( _
            weldingPriceSheetName, vendorPriceCaches, recordKey, vendorName, _
            (vendorColumn = WELD_COL_WELDING_VENDOR), dayNightText)
    Else
        Dim unitPriceSheetName As String
        unitPriceSheetName = ResolveUnitPriceSheetName( _
            lineSheetMap, CommonNzText(GetArrayCellValue(lineData, rowIndex, 1)))
        If unitPriceSheetName = "" Then Exit Function

        Dim vendorPriceRows As Object
        Set vendorPriceRows = GetVendorUnitPriceRows( _
            unitPriceSheetName, vendorName, vendorPriceCaches)
        If vendorPriceRows Is Nothing Then Exit Function
        If Not vendorPriceRows.Exists(recordKey) Then Exit Function

        Dim dayNightPrices As Variant
        dayNightPrices = vendorPriceRows(recordKey)
        LookupSubcontractorVendorPriceFromArrays = SelectUsableDayNightPrice(dayNightText, dayNightPrices)
    End If
End Function

Public Sub RefreshBasicInfoConstructionTotals(Optional ByVal changedVendorIndex As Long = 0)
    On Error GoTo ErrorHandler

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim fullRefresh As Boolean
    fullRefresh = (changedVendorIndex <= 0)

    If Not fullRefresh Then
        If changedVendorIndex < 1 Or changedVendorIndex > vendorCount Then Exit Sub
    End If

    Dim vendorNames() As String
    Dim vendorTotals() As Double
    ReDim vendorNames(1 To BASIC_INFO_VENDOR_MAX_BLOCKS)
    ReDim vendorTotals(1 To BASIC_INFO_VENDOR_MAX_BLOCKS)

    Dim vendorStart As Long
    Dim vendorEnd As Long
    If fullRefresh Then
        vendorStart = 1
        vendorEnd = vendorCount
    Else
        vendorStart = changedVendorIndex
        vendorEnd = changedVendorIndex
    End If

    Dim i As Long
    For i = vendorStart To vendorEnd
        vendorNames(i) = GetBasicInfoCellText(wsInfo, _
            wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, BasicInfoVendorColumn(i)).Address)
    Next i

    If fullRefresh Then
        Dim vendorNameLog As String
        For i = 1 To vendorCount
            vendorNameLog = vendorNameLog & " [" & i & ":" & vendorNames(i) & "]"
        Next i
        LogCI "基本情報業者 F9件数=" & vendorCount & vendorNameLog
    End If

    Dim BranchName As String
    BranchName = GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL)
    Dim vendorAliasMap As Object
    Set vendorAliasMap = GetVendorAliasMap(BranchName)

    Dim worksTotal As Double
    Dim purchaseTotal As Double

    Dim ws As Worksheet
    Dim columnMap As Object
    Dim vendorKey As String
    For Each ws In ThisWorkbook.Worksheets
        If fullRefresh And IsPurchaseOutputSheet(ws) Then
            purchaseTotal = purchaseTotal + SumOutputJrAmount(ws)
        ElseIf IsConstructionOutputSheet(ws) Then
            If fullRefresh Then
                worksTotal = worksTotal + SumOutputJrAmount(ws)
            End If
            Set columnMap = BuildSheetVendorAmountColumnMap(ws, vendorAliasMap)
            For i = vendorStart To vendorEnd
                If vendorNames(i) <> "" Then
                    vendorKey = ResolveVendorCanonicalKey(vendorNames(i), vendorAliasMap)
                    If vendorKey <> "" Then
                        If Not columnMap Is Nothing Then
                            If columnMap.Exists(vendorKey) Then
                                vendorTotals(i) = vendorTotals(i) + _
                                    SumVendorAmountByColumn(ws, CLng(columnMap(vendorKey)))
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next ws

    If fullRefresh Then
        WriteBasicInfoAmount wsInfo, BASIC_INFO_WORKS_TOTAL_CELL, worksTotal
        WriteBasicInfoAmount wsInfo, BASIC_INFO_PURCHASE_TOTAL_CELL, purchaseTotal
        UpdateBasicInfoTaxTotals wsInfo
    End If

    Dim totalCellAddress As String
    Dim totalCell As Range
    Dim writeStart As Long
    Dim writeEnd As Long
    If fullRefresh Then
        writeStart = 1
        writeEnd = BASIC_INFO_VENDOR_MAX_BLOCKS
    Else
        writeStart = changedVendorIndex
        writeEnd = changedVendorIndex
    End If

    For i = writeStart To writeEnd
        totalCellAddress = wsInfo.Cells(BASIC_INFO_VENDOR_TOTAL_ROW, _
                                        BasicInfoVendorColumn(i)).Address
        If i <= vendorCount And vendorNames(i) <> "" Then
            WriteBasicInfoAmount wsInfo, totalCellAddress, vendorTotals(i), True
        Else
            WriteBasicInfoAmount wsInfo, totalCellAddress, 0, False
        End If

        Set totalCell = wsInfo.Range(totalCellAddress)
        If totalCell.MergeCells Then Set totalCell = totalCell.MergeArea.Cells(1, 1)
        totalCell.NumberFormatLocal = BasicInfoYenNumberFormat()
    Next i
    Exit Sub

ErrorHandler:
    LogCI "基本情報合計金額更新エラー Err " & Err.Number & ": " & Err.Description
    Err.Clear
End Sub

Public Sub ClearVendorAliasMapCache()
    Set mVendorAliasMapCache = Nothing
End Sub

Private Function GetVendorAliasMap(ByVal branchName As String) As Object
    Dim cacheKey As String
    cacheKey = CommonNormalizeText(branchName)

    If mVendorAliasMapCache Is Nothing Then
        Set mVendorAliasMapCache = CreateObject("Scripting.Dictionary")
        mVendorAliasMapCache.CompareMode = vbTextCompare
    End If
    If cacheKey <> "" Then
        If mVendorAliasMapCache.Exists(cacheKey) Then
            Dim cachedAliasMap As Object
            Set cachedAliasMap = mVendorAliasMapCache(cacheKey)
            If Not cachedAliasMap Is Nothing Then
                Set GetVendorAliasMap = cachedAliasMap
                Exit Function
            End If
        End If
    End If

    Dim aliasMap As Object
    Set aliasMap = BuildVendorAliasMap(branchName)
    If cacheKey <> "" Then
        If mVendorAliasMapCache.Exists(cacheKey) Then
            Set mVendorAliasMapCache(cacheKey) = aliasMap
        Else
            mVendorAliasMapCache.Add cacheKey, aliasMap
        End If
    End If
    Set GetVendorAliasMap = aliasMap
End Function

Private Function BuildSheetVendorAmountColumnMap(ByVal ws As Worksheet, _
                                                 ByVal aliasMap As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim kindColumn As Long
    kindColumn = FindHeaderColumn(ws, "工種分類")
    If kindColumn <= OutputSheetSubconPriceFirstCol(ws) Then
        Set BuildSheetVendorAmountColumnMap = result
        Exit Function
    End If

    Dim subconFirstCol As Long
    subconFirstCol = OutputSheetSubconPriceFirstCol(ws)

    Dim c As Long
    For c = subconFirstCol To kindColumn - 1
        Dim headerText As String
        headerText = CommonNzText(ws.Cells(1, c).value)
        If Len(headerText) > Len("金額") Then
            If Right$(headerText, Len("金額")) = "金額" Then
                Dim vendorKey As String
                vendorKey = ResolveVendorCanonicalKey(Left$(headerText, Len(headerText) - Len("金額")), aliasMap)
                If vendorKey <> "" Then
                    If Not result.Exists(vendorKey) Then
                        result.Add vendorKey, c
                    End If
                End If
            End If
        End If
    Next c

    Set BuildSheetVendorAmountColumnMap = result
End Function

Private Function SumVendorAmountByColumn(ByVal ws As Worksheet, _
                                         ByVal amountColumn As Long) As Double
    Dim seiriColumn As Long
    seiriColumn = FindHeaderColumn(ws, "整理番号")
    If seiriColumn = 0 Or amountColumn = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, seiriColumn)
    If lastRow < 2 Then Exit Function

    SumVendorAmountByColumn = RoundDownAmount(SumNumericColumn(ws, amountColumn, lastRow))
End Function

Private Function BasicInfoVendorColumn(ByVal vendorIndex As Long) As Long
    BasicInfoVendorColumn = BASIC_INFO_VENDOR_FIRST_COL + _
                            ((vendorIndex - 1) * BASIC_INFO_VENDOR_STEP_COLS)
End Function

Private Function GetBasicInfoVendorBlockCount(ByVal wsInfo As Worksheet) As Long
    Dim countValue As Long
    countValue = CLng(Val(StrConv(CommonNzText( _
        wsInfo.Range(BASIC_INFO_VENDOR_COUNT_CELL).value), vbNarrow)))
    If countValue < 1 Then countValue = 1
    If countValue > BASIC_INFO_VENDOR_MAX_BLOCKS Then countValue = BASIC_INFO_VENDOR_MAX_BLOCKS
    GetBasicInfoVendorBlockCount = countValue
End Function

Private Function IsPurchaseOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    IsPurchaseOutputSheet = _
        (StrComp(ws.Name, PURCHASE_ORDER_SHEET_NAME, vbTextCompare) = 0) Or _
        (StrComp(ws.Name, PURCHASE_NOTICE_SHEET_NAME, vbTextCompare) = 0)
End Function

Private Function IsConstructionOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If IsPurchaseOutputSheet(ws) Then Exit Function

    IsConstructionOutputSheet = _
        ((FindHeaderColumn(ws, "施工業者") > 0) Or IsWeldingOutputSheet(ws)) And _
        (FindHeaderColumn(ws, "整理番号") > 0) And _
        (FindHeaderColumn(ws, "JR金額") > 0) And _
        (FindHeaderColumn(ws, "工種分類") > 0)
End Function

Private Function SumOutputJrAmount(ByVal ws As Worksheet) As Double
    Dim seiriColumn As Long
    Dim amountColumn As Long
    seiriColumn = FindHeaderColumn(ws, "整理番号")
    amountColumn = FindHeaderColumn(ws, "JR金額")
    If seiriColumn = 0 Or amountColumn = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, seiriColumn)
    If lastRow < 2 Then Exit Function

    SumOutputJrAmount = RoundDownAmount(SumNumericColumn(ws, amountColumn, lastRow))
End Function

Private Function SumVendorAmountOnSheet(ByVal ws As Worksheet, _
                                        ByVal vendorName As String, _
                                        ByVal aliasMap As Object) As Double
    Dim vendorKey As String
    vendorKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
    If vendorKey = "" Then Exit Function

    Dim columnMap As Object
    Set columnMap = BuildSheetVendorAmountColumnMap(ws, aliasMap)
    If columnMap.Exists(vendorKey) Then
        SumVendorAmountOnSheet = SumVendorAmountByColumn(ws, CLng(columnMap(vendorKey)))
    End If
End Function

Private Function SumNumericColumn(ByVal ws As Worksheet, _
                                  ByVal targetColumn As Long, _
                                  ByVal lastRow As Long) As Double
    Dim totalAmount As Double
    Dim r As Long
    For r = 2 To lastRow
        Dim cellValue As Variant
        cellValue = ws.Cells(r, targetColumn).value
        If Not IsError(cellValue) Then
            If IsNumeric(cellValue) Then totalAmount = totalAmount + CDbl(cellValue)
        End If
    Next r

    SumNumericColumn = totalAmount
End Function

Private Function RoundDownAmount(ByVal amount As Double) As Double
    RoundDownAmount = Fix(amount)
End Function

Private Function GetBasicInfoCellText(ByVal wsInfo As Worksheet, _
                                      ByVal cellAddress As String) As String
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
    GetBasicInfoCellText = Trim$(CommonNzText(targetCell.value))
End Function

Private Sub WriteBasicInfoAmount(ByVal wsInfo As Worksheet, _
                                 ByVal cellAddress As String, _
                                 ByVal amount As Double, _
                                 Optional ByVal hasValue As Boolean = True)
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)

    targetCell.NumberFormatLocal = BASIC_INFO_TOTAL_NUMBER_FORMAT
    If hasValue Then
        targetCell.value = RoundDownAmount(amount)
    Else
        targetCell.ClearContents
    End If
End Sub

'  UpdateBasicInfoTaxTotals
'  基本情報シートの C33(小計)・C34(消費税)・C35(税込合計) を更新し、
'  C31:C35 に「\＋桁区切り」の表示形式を適用する。
'  C33 = C31 + C32
'  C34 = C33 × 税率(B34の表記から取得。取得できない場合は10%) ※小数点以下切り捨て
'  C35 = C33 + C34
Public Sub UpdateBasicInfoTaxTotals(Optional ByVal wsInfo As Worksheet)
    On Error GoTo ErrorHandler

    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim subtotal As Double
    subtotal = GetBasicInfoCellAmount(wsInfo, BASIC_INFO_WORKS_TOTAL_CELL) + _
               GetBasicInfoCellAmount(wsInfo, BASIC_INFO_PURCHASE_TOTAL_CELL)

    Dim taxAmount As Double
    taxAmount = Fix(subtotal * ResolveBasicInfoTaxRate(wsInfo))

    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_SUBTOTAL_CELL, subtotal
    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_TAX_CELL, taxAmount
    WriteBasicInfoPlainValue wsInfo, BASIC_INFO_GRAND_TOTAL_CELL, subtotal + taxAmount

    ApplyBasicInfoYenTotalFormat wsInfo

    LogCI "税込合計更新: 小計=" & subtotal & " / 消費税=" & taxAmount & _
          " / 税込合計=" & (subtotal + taxAmount)
    Exit Sub

ErrorHandler:
    LogCI "基本情報税込合計更新エラー Err " & Err.Number & ": " & Err.Description
End Sub

'  GetBasicInfoCellAmount
'  指定セルの数値を取得する(結合セル対応。空欄・非数値・エラー値は0)。
Private Function GetBasicInfoCellAmount(ByVal wsInfo As Worksheet, _
                                        ByVal cellAddress As String) As Double
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)

    Dim cellValue As Variant
    cellValue = targetCell.value
    If Not IsError(cellValue) Then
        If IsNumeric(cellValue) Then GetBasicInfoCellAmount = CDbl(cellValue)
    End If
End Function

'  WriteBasicInfoPlainValue
'  指定セルへ値のみを書き込む(結合セル対応。表示形式は変更しない)。
Private Sub WriteBasicInfoPlainValue(ByVal wsInfo As Worksheet, _
                                     ByVal cellAddress As String, _
                                     ByVal amount As Double)
    Dim targetCell As Range
    Set targetCell = wsInfo.Range(cellAddress)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
    targetCell.value = amount
End Sub

'  ResolveBasicInfoTaxRate
'  B34 のラベル(例:「消費税(10%)」)から税率を抽出する。
'  「%」直前の数値を税率として解釈し、取得できない場合は既定の10%を返す。
Private Function ResolveBasicInfoTaxRate(ByVal wsInfo As Worksheet) As Double
    ResolveBasicInfoTaxRate = BASIC_INFO_TAX_RATE_DEFAULT

    Dim labelText As String
    On Error Resume Next
    labelText = StrConv(CommonNzText(wsInfo.Range(BASIC_INFO_TAX_LABEL_CELL).value), vbNarrow)
    On Error GoTo 0
    If labelText = "" Then Exit Function

    Dim numText As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(labelText)
        ch = Mid$(labelText, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            numText = numText & ch
        ElseIf ch = "%" Then
            If numText <> "" And IsNumeric(numText) Then
                ResolveBasicInfoTaxRate = CDbl(numText) / 100
            End If
            Exit Function
        Else
            numText = ""
        End If
    Next i
End Function

'  ApplyBasicInfoYenTotalFormat
'  C31:C35 に「\＋桁区切り」(負数は赤字)の表示形式を適用する。
Private Sub ApplyBasicInfoYenTotalFormat(ByVal wsInfo As Worksheet)
    wsInfo.Range(BASIC_INFO_YEN_TOTAL_RANGE).NumberFormatLocal = BasicInfoYenNumberFormat()
End Sub

'  BasicInfoYenNumberFormat
'  「\＋桁区切り」(負数は赤字)の表示形式文字列を返す。
'  \記号はCP932での文字化けを避けるため ChrW$ で生成する。
Private Function BasicInfoYenNumberFormat() As String
    Dim yenMark As String
    yenMark = ChrW$(&HA5)   ' \

    BasicInfoYenNumberFormat = yenMark & "#,##0;[赤]-" & yenMark & "#,##0"
End Function

Private Function CollectSelectedSubcontractors(ByVal ws As Worksheet, _
                                                ByVal lastRow As Long) As Collection
    Dim result As New Collection
    If lastRow < 2 Then
        Set CollectSelectedSubcontractors = result
        Exit Function
    End If

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)

    Dim aliasMap As Object
    Set aliasMap = Nothing
    If Not wsInfo Is Nothing Then
        Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))
    End If

    Dim sheetByCanonical As Object
    Set sheetByCanonical = CreateObject("Scripting.Dictionary")
    sheetByCanonical.CompareMode = vbTextCompare

    Dim sheetOrderKeys As New Collection

    Dim vendorColumns As Collection
    Set vendorColumns = OutputSheetVendorColumns(ws)

    Dim r As Long
    Dim vendorCol As Variant
    For r = 2 To lastRow
        For Each vendorCol In vendorColumns
            Dim vendorName As String
            Dim canonicalKey As String
            vendorName = Trim$(CommonNzText(ws.Cells(r, CLng(vendorCol)).value))
            If vendorName = "" Then GoTo NextVendorCol

            canonicalKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
            If canonicalKey = "" Then GoTo NextVendorCol

            If Not sheetByCanonical.Exists(canonicalKey) Then
                sheetByCanonical.Add canonicalKey, vendorName
                sheetOrderKeys.Add canonicalKey
            End If
NextVendorCol:
        Next vendorCol
    Next r

    If sheetByCanonical.Count = 0 Then
        Set CollectSelectedSubcontractors = result
        Exit Function
    End If

    Dim usedCanonical As Object
    Set usedCanonical = CreateObject("Scripting.Dictionary")
    usedCanonical.CompareMode = vbTextCompare

    If Not wsInfo Is Nothing Then
        If IsWeldingOutputSheet(ws) Then
            AppendOrderedBasicInfoVendorsByWorkType wsInfo, aliasMap, sheetByCanonical, usedCanonical, result, WELDING_WORK_TYPE_KEYWORD
            AppendOrderedBasicInfoVendorsByWorkType wsInfo, aliasMap, sheetByCanonical, usedCanonical, result, TRACK_WORK_TYPE_KEYWORD
        Else
            Dim vendorCount As Long
            Dim i As Long
            vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

            For i = 1 To vendorCount
                Dim basicInfoName As String
                Dim basicInfoKey As String
                basicInfoName = GetBasicInfoCellText(wsInfo, _
                    wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, BasicInfoVendorColumn(i)).Address)
                If basicInfoName = "" Then GoTo NextBasicInfoVendor

                basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
                If basicInfoKey <> "" And sheetByCanonical.Exists(basicInfoKey) Then
                    If Not usedCanonical.Exists(basicInfoKey) Then
                        usedCanonical.Add basicInfoKey, True
                        result.Add CStr(sheetByCanonical(basicInfoKey))
                    End If
                End If
NextBasicInfoVendor:
            Next i
        End If
    End If

    Dim fallbackKey As Variant
    For Each fallbackKey In sheetOrderKeys
        If Not usedCanonical.Exists(CStr(fallbackKey)) Then
            result.Add CStr(sheetByCanonical(CStr(fallbackKey)))
        End If
    Next fallbackKey

    Set CollectSelectedSubcontractors = result
End Function

Private Sub AppendOrderedBasicInfoVendorsByWorkType( _
    ByVal wsInfo As Worksheet, _
    ByVal aliasMap As Object, _
    ByVal sheetByCanonical As Object, _
    ByVal usedCanonical As Object, _
    ByVal result As Collection, _
    ByVal workTypeKeyword As String)

    Dim vendorNameMap As Object
    Set vendorNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)
    If vendorNameMap Is Nothing Then Exit Sub

    Dim blockIndex As Long
    For blockIndex = 1 To BASIC_INFO_VENDOR_MAX_BLOCKS
        Dim valueCol As Long
        valueCol = BasicInfoVendorColumn(blockIndex)
        If Not BasicInfoBlockMatchesWorkType(wsInfo, valueCol, workTypeKeyword) Then GoTo NextWorkTypeBlock

        Dim basicInfoName As String
        Dim basicInfoKey As String
        Dim mappedName As String
        basicInfoName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).Address)
        If basicInfoName = "" Then GoTo NextWorkTypeBlock

        basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
        If basicInfoKey = "" Then GoTo NextWorkTypeBlock
        If Not sheetByCanonical.Exists(basicInfoKey) Then GoTo NextWorkTypeBlock
        If usedCanonical.Exists(basicInfoKey) Then GoTo NextWorkTypeBlock

        mappedName = Trim$(CommonNzText(sheetByCanonical(basicInfoKey)))
        If mappedName = "" Then GoTo NextWorkTypeBlock

        usedCanonical.Add basicInfoKey, True
        result.Add mappedName
NextWorkTypeBlock:
    Next blockIndex
End Sub

Private Function BasicInfoBlockMatchesWorkType(ByVal wsInfo As Worksheet, _
                                               ByVal valueCol As Long, _
                                               ByVal workTypeKeyword As String) As Boolean
    Dim workTypeText As String
    workTypeText = CommonRemoveAllSpaces(CommonNormalizeText( _
        CommonNzText(wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueCol).value)))
    BasicInfoBlockMatchesWorkType = (workTypeText <> "") And _
        (InStr(1, workTypeText, workTypeKeyword, vbTextCompare) > 0)
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
    If vendorPriceCaches Is Nothing Then Exit Function

    Dim cacheKey As String
    cacheKey = unitPriceSheetName & "|" & NormalizeVendorPriceName(vendorName)
    If vendorPriceCaches.Exists(cacheKey) Then
        Dim cachedRows As Object
        Set cachedRows = vendorPriceCaches(cacheKey)
        If Not cachedRows Is Nothing Then
            Set GetVendorUnitPriceRows = cachedRows
            Exit Function
        End If
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(unitPriceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then
        LogCI "施工会社単価: シート未検出 [" & unitPriceSheetName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If
    If Not mod_MaterialPriceImport.IsConstructionUnitPriceSheet(priceSheet) Then
        LogCI "施工会社単価: 工事単価シートではない [" & unitPriceSheetName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim aliasMap As Object
    Set aliasMap = Nothing
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If Not wsInfo Is Nothing Then
        Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))
    End If

    Dim vendorDayColumn As Long
    vendorDayColumn = FindUnitPriceVendorDayColumn(priceSheet, vendorName, aliasMap)
    If vendorDayColumn = 0 Then
        LogCI "施工会社単価: 業者列未検出 sheet=[" & unitPriceSheetName & "] vendor=[" & vendorName & "]"
        StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
        Set GetVendorUnitPriceRows = result
        Exit Function
    End If

    Dim priceLastRow As Long
    priceLastRow = GetUnitPriceSheetLastDataRow(priceSheet)

    Dim r As Long
    For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
        Dim recordKey As String
        recordKey = NormalizeRecordKey(priceSheet.Cells(r, COL_SEIRI).value)
        If recordKey = "" Then GoTo NextVendorPriceRow
        If Not IsUnitPriceVendorRowPriceEligible(priceSheet, r, vendorDayColumn) Then GoTo NextVendorPriceRow

        If Not result.Exists(recordKey) Then
            result.Add recordKey, Array(priceSheet.Cells(r, vendorDayColumn).Value2, _
                                        priceSheet.Cells(r, vendorDayColumn + 1).Value2)
        End If
NextVendorPriceRow:
    Next r

    LogCI "施工会社単価: sheet=[" & unitPriceSheetName & "] vendor=[" & vendorName & _
          "] col=" & vendorDayColumn & " keys=" & result.Count

    StoreVendorUnitPriceCache vendorPriceCaches, cacheKey, result
    Set GetVendorUnitPriceRows = result
End Function

Private Sub StoreVendorUnitPriceCache(ByVal vendorPriceCaches As Object, _
                                      ByVal cacheKey As String, _
                                      ByVal priceRows As Object)
    If vendorPriceCaches Is Nothing Then Exit Sub
    If vendorPriceCaches.Exists(cacheKey) Then
        Set vendorPriceCaches(cacheKey) = priceRows
    Else
        vendorPriceCaches.Add cacheKey, priceRows
    End If
End Sub

Private Function FindUnitPriceVendorDayColumn(ByVal priceSheet As Worksheet, _
                                              ByVal vendorName As String, _
                                              Optional ByVal aliasMap As Object = Nothing) As Long
    Dim vendorKey As String
    vendorKey = ResolveVendorCanonicalKey(vendorName, aliasMap)
    If vendorKey = "" Then Exit Function

    Dim lastColumn As Long
    lastColumn = priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, _
                                  priceSheet.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = UNIT_PRICE_VENDOR_FIRST_DAY_COL To lastColumn Step 2
        Dim headerCell As Range
        Dim headerKey As String
        Set headerCell = priceSheet.Cells(UNIT_PRICE_VENDOR_NAME_ROW, c)
        On Error Resume Next
        If headerCell.MergeCells Then
            Dim mergedHeader As Range
            Set mergedHeader = headerCell.MergeArea.Cells(1, 1)
            If Not mergedHeader Is Nothing Then Set headerCell = mergedHeader
        End If
        On Error GoTo 0
        headerKey = ResolveVendorCanonicalKey(CommonNzText(headerCell.value), aliasMap)
        If headerKey <> "" Then
            If StrComp(headerKey, vendorKey, vbTextCompare) = 0 Then
                FindUnitPriceVendorDayColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function GetUnitPriceSheetLastDataRow(ByVal priceSheet As Worksheet) As Long
    Dim scanStartRow As Long
    scanStartRow = UNIT_PRICE_DATA_START_ROW
    If Not priceSheet.UsedRange Is Nothing Then
        Dim usedLastRow As Long
        usedLastRow = priceSheet.UsedRange.Row + priceSheet.UsedRange.Rows.Count - 1
        If usedLastRow > scanStartRow Then scanStartRow = usedLastRow
    End If

    Dim rowIndex As Long
    For rowIndex = scanStartRow To UNIT_PRICE_DATA_START_ROW Step -1
        If Trim$(CommonNzText(priceSheet.Cells(rowIndex, COL_SEIRI).value)) <> "" Then
            GetUnitPriceSheetLastDataRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    GetUnitPriceSheetLastDataRow = UNIT_PRICE_DATA_START_ROW - 1
End Function

Private Function IsSanpaiTypeText(ByVal typeText As String) As Boolean
    IsSanpaiTypeText = (InStr(1, CommonRemoveAllSpaces(CommonNormalizeText(typeText)), _
                                SANPAI_KEYWORD, vbTextCompare) > 0)
End Function

Private Function UnitPriceValueIsUsable(ByVal value As Variant) As Boolean
    If IsEmpty(value) Or IsError(value) Then Exit Function
    If Len(Trim$(CommonNzText(value))) = 0 Then Exit Function
    UnitPriceValueIsUsable = IsNumeric(value)
End Function

Private Function IsUnitPriceSheetSanpaiRow(ByVal priceSheet As Worksheet, _
                                           ByVal rowIndex As Long) As Boolean
    IsUnitPriceSheetSanpaiRow = IsSanpaiTypeText( _
        CommonNzText(priceSheet.Cells(rowIndex, UNIT_PRICE_WORK_TYPE_COL).value))
End Function

Private Function IsUnitPriceVendorRowPriceEligible(ByVal priceSheet As Worksheet, _
                                                   ByVal rowIndex As Long, _
                                                   ByVal vendorDayColumn As Long) As Boolean
    If IsUnitPriceSheetSanpaiRow(priceSheet, rowIndex) Then Exit Function

    Dim dayPrice As Variant
    Dim nightPrice As Variant
    dayPrice = priceSheet.Cells(rowIndex, vendorDayColumn).Value2
    nightPrice = priceSheet.Cells(rowIndex, vendorDayColumn + 1).Value2

    '  独自工種など JR 参照(E/F)が空でも、施工会社列へ手入力された単価は採用する
    IsUnitPriceVendorRowPriceEligible = UnitPriceValueIsUsable(dayPrice) Or _
                                        UnitPriceValueIsUsable(nightPrice)
End Function

Private Function SelectUsableDayNightPrice(ByVal dayNightText As String, _
                                             ByVal dayNightPrices As Variant) As Variant
    Dim selectedPrice As Variant
    selectedPrice = SelectDayNightPrice(dayNightText, dayNightPrices)
    If UnitPriceValueIsUsable(selectedPrice) Then
        SelectUsableDayNightPrice = selectedPrice
    End If
End Function

Private Function NormalizeVendorPriceName(ByVal vendorName As String) As String
    NormalizeVendorPriceName = CommonRemoveAllSpaces(CommonNormalizeText(vendorName))
End Function

'  ResolveVendorCanonicalKey
'  業者マスタ(別名表)を正として、正規名・略称名のどちらの表記でも
'  同一の正規名キーへ解決する。マスタに無い名称は正規化文字列をそのまま返す
'  ためフォールバックされ、参照エラーにはならない。
Private Function ResolveVendorCanonicalKey(ByVal vendorName As String, _
                                           ByVal aliasMap As Object) As String
    Dim normalizedKey As String
    normalizedKey = NormalizeVendorPriceName(vendorName)
    If normalizedKey = "" Then Exit Function

    If Not aliasMap Is Nothing Then
        If aliasMap.Exists(normalizedKey) Then
            ResolveVendorCanonicalKey = CStr(aliasMap(normalizedKey))
            Exit Function
        End If
    End If

    ResolveVendorCanonicalKey = normalizedKey
End Function

Public Function ResolveBasicInfoVendorInfoIndex(ByVal vendorDisplayName As String, _
                                                Optional ByVal workTypeKeyword As String = "") As Long
    ResolveBasicInfoVendorInfoIndex = 0

    Dim normalizedDisplay As String
    normalizedDisplay = NormalizeVendorPriceName(vendorDisplayName)
    If normalizedDisplay = "" Then Exit Function

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim aliasMap As Object
    Set aliasMap = GetVendorAliasMap(GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL))

    Dim canonicalKey As String
    canonicalKey = ResolveVendorCanonicalKey(vendorDisplayName, aliasMap)

    Dim vendorNameMap As Object
    Set vendorNameMap = mod_VendorMaster.BuildVendorUnitPriceNameMap(wsInfo)

    Dim vendorCount As Long
    vendorCount = GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueCol As Long
        valueCol = BasicInfoVendorColumn(vendorIndex)

        If workTypeKeyword <> "" Then
            If Not BasicInfoBlockMatchesWorkType(wsInfo, valueCol, workTypeKeyword) Then GoTo NextVendorIndex
        End If

        Dim basicInfoName As String
        basicInfoName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol).Address)
        If basicInfoName = "" Then GoTo NextVendorIndex

        Dim basicInfoKey As String
        basicInfoKey = ResolveVendorCanonicalKey(basicInfoName, aliasMap)
        If canonicalKey <> "" And basicInfoKey = canonicalKey Then
            ResolveBasicInfoVendorInfoIndex = vendorIndex
            Exit Function
        End If

        If NormalizeVendorPriceName(basicInfoName) = normalizedDisplay Then
            ResolveBasicInfoVendorInfoIndex = vendorIndex
            Exit Function
        End If

        If Not vendorNameMap Is Nothing Then
            Dim nameKey As String
            nameKey = CommonNormalizeText(basicInfoName)
            If vendorNameMap.Exists(nameKey) Then
                Dim mappedDisplayName As String
                mappedDisplayName = Trim$(CommonNzText(vendorNameMap(nameKey)))
                If NormalizeVendorPriceName(mappedDisplayName) = normalizedDisplay Then
                    ResolveBasicInfoVendorInfoIndex = vendorIndex
                    Exit Function
                End If
            End If
        End If
NextVendorIndex:
    Next vendorIndex
End Function

'  BuildVendorAliasMap
'  業者マスタ(全社版).xlsx の「支店名(基本情報B6)」シートを開き、
'  A列=業者名(略称) / B列=請求者氏名(正規名) を読み込んで、
'  正規化(略称)・正規化(正規名) の双方を 正規化(正規名) へ対応付けた辞書を返す。
'  (1行目は見出し行だが、実業者名と一致しない無害なエントリになるだけ)
'  マスタ未検出・シート未検出・読込失敗時は空辞書を返す(突合は正規化のみで継続)。
Private Function BuildVendorAliasMap(ByVal branchName As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim connection As Object
    Dim recordset As Object

    On Error GoTo Cleanup

    If Trim$(branchName) = "" Then
        LogCI "業者マスタ別名: 基本情報B6(支店名)が空のため名寄せなし"
        GoTo Cleanup
    End If

    Dim masterPath As String
    masterPath = ResolveVendorMasterPath()
    If masterPath = "" Then
        LogCI "業者マスタ未検出 -> 名寄せなし(正規化のみで突合)"
        GoTo Cleanup
    End If

    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then
        LogCI "業者マスタADO接続不可 path=[" & masterPath & "]"
        GoTo Cleanup
    End If

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(connection, branchName)
    If actualSheetName = "" Then
        LogCI "業者マスタに支店シート[" & branchName & "]が見つかりません -> 名寄せなし"
        GoTo Cleanup
    End If

    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT [F" & VENDOR_MASTER_OFFICIAL_COL & "], [F" & VENDOR_MASTER_ABBREV_COL & "] FROM " & _
                   BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

    Dim official As String, abbrev As String, canonicalKey As String
    Do Until recordset.EOF
        official = CommonNzText(recordset.Fields(0).value)
        abbrev = CommonNzText(recordset.Fields(1).value)
        canonicalKey = NormalizeVendorPriceName(official)
        If canonicalKey <> "" Then
            AddVendorAlias result, official, canonicalKey
            AddVendorAlias result, abbrev, canonicalKey
        End If
        recordset.MoveNext
    Loop

    LogCI "業者マスタ別名 件数=" & result.Count & " 支店=[" & branchName & _
          "] sheet=[" & actualSheetName & "]"

Cleanup:
    If Err.Number <> 0 Then
        LogCI "業者マスタ読込エラー Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    Set BuildVendorAliasMap = result
End Function

'  AddVendorAlias
'  正規化した別名を正規名キーへ登録する(空・重複は無視)。
Private Sub AddVendorAlias(ByVal aliasMap As Object, _
                           ByVal aliasName As String, _
                           ByVal canonicalKey As String)
    Dim normalizedAlias As String
    normalizedAlias = NormalizeVendorPriceName(aliasName)
    If normalizedAlias = "" Then Exit Sub
    If Not aliasMap.Exists(normalizedAlias) Then aliasMap.Add normalizedAlias, canonicalKey
End Sub

'  ResolveVendorMasterPath
'  業者マスタ(全社版).xlsx のパスを複数候補から解決する。
Private Function ResolveVendorMasterPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim candidates As Collection
    Set candidates = New Collection

    Dim unitMasterPath As String
    unitMasterPath = ResolveMasterFilePath()
    If unitMasterPath <> "" Then
        AddUniqueText candidates, _
            fso.BuildPath(fso.GetParentFolderName(unitMasterPath), VENDOR_MASTER_FILE)
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        AddUniqueText candidates, _
            fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                          MASTER_DATA_FOLDER & "\" & VENDOR_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText candidates, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
            "線路出張所用_注文書_請求書アクセスサイト - ドキュメント\" & _
            MASTER_DATA_FOLDER & "\" & VENDOR_MASTER_FILE
    End If

    Dim candidate As Variant
    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            ResolveVendorMasterPath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Sub ApplyWeldingOutputSheetColumnAlignment(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not IsWeldingOutputSheet(ws) Then Exit Sub

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 1 Then lastRow = 1

    ws.Columns(WELD_COL_WELDING_VENDOR).HorizontalAlignment = xlCenter
    ws.Columns(WELD_COL_TRACK_VENDOR).HorizontalAlignment = xlCenter
    ws.Columns("F").HorizontalAlignment = xlCenter
    ws.Columns("I").HorizontalAlignment = xlCenter
    ws.Columns("R").HorizontalAlignment = xlCenter
End Sub

Private Sub FormatSubcontractorPriceColumns(ByVal ws As Worksheet, _
                                            ByVal lastRow As Long, _
                                            ByVal columnCount As Long, _
                                            Optional ByVal firstColumn As Long = 0)
    Dim lastColumn As Long
    If firstColumn = 0 Then firstColumn = OutputSheetSubconPriceFirstCol(ws)
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

    ws.Range(ws.Columns(firstColumn), ws.Columns(lastColumn)).AutoFit
End Sub

Private Function IsSanpaiRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    IsSanpaiRow = IsSanpaiTypeText( _
        CommonNzText(ws.Cells(rowIndex, OutputSheetCol(ws, COL_TYPE)).value))
End Function

Private Function GetSanpaiFillColor() As Long
    If mSanpaiFillColorCached Then
        GetSanpaiFillColor = mSanpaiFillColorCache
        Exit Function
    End If

    mSanpaiFillColorCache = SANPAI_FALLBACK_FILL_COLOR

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(ws) Then
            Dim priceLastRow As Long
            priceLastRow = ws.Cells(ws.rows.Count, COL_SEIRI).End(xlUp).Row

            Dim r As Long, c As Long
            For r = UNIT_PRICE_DATA_START_ROW To priceLastRow
                For c = 5 To 6
                    If ws.Cells(r, c).Interior.Pattern <> xlPatternNone Then
                        mSanpaiFillColorCache = ws.Cells(r, c).Interior.Color
                        GoTo SanpaiFillColorDone
                    End If
                Next c
            Next r
        End If
    Next ws

SanpaiFillColorDone:
    mSanpaiFillColorCached = True
    GetSanpaiFillColor = mSanpaiFillColorCache
End Function

Public Sub ApplySanpaiRowRestrictions(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim fillColor As Long
    fillColor = GetSanpaiFillColor()

    Dim restrictedCount As Long
    Dim r As Long
    Dim vendorCol As Variant
    Dim vendorColumns As Collection
    Set vendorColumns = OutputSheetVendorColumns(ws)
    For r = 2 To lastRow
        If IsSanpaiRow(ws, r) Then
            For Each vendorCol In vendorColumns
                With ws.Cells(r, CLng(vendorCol))
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
                        .ErrorMessage = "産廃処理の行は施工会社を入力できません。"
                        .ShowError = True
                    End With
                End With
            Next vendorCol
            restrictedCount = restrictedCount + 1
        End If
    Next r

    LogCI "産廃処理行: A列塗りつぶし・入力不可=" & restrictedCount & " 行"
End Sub

Private Sub WriteTotalCells(ByVal ws As Worksheet, ByVal totalRow As Long, _
                            ByVal labelColumn As Long, ByVal labelText As String, _
                            ByVal sumColumn As Long, ByVal sumLastRow As Long)
    With ws.Cells(totalRow, labelColumn)
        .value = labelText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With

    With ws.Cells(totalRow, sumColumn)
        .FormulaR1C1 = "=ROUNDDOWN(SUM(R2C:R" & sumLastRow & "C),0)"
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
    End With

    DrawDoubleBorder ws.Cells(totalRow, labelColumn), RGB(0, 0, 0)
    DrawDoubleBorder ws.Cells(totalRow, sumColumn), RGB(255, 0, 0)
End Sub

Private Function IsOutputTotalLabelText(ByVal text As String) As Boolean
    Dim normalized As String
    normalized = Trim$(CommonNormalizeText(CommonNzText(text)))
    IsOutputTotalLabelText = (normalized = "JR合計") Or _
        (Len(normalized) >= 2 And Right$(normalized, 2) = "合計")
End Function

Private Sub ClearDoubleBorder(ByVal target As Range)
    Dim edgeId As Variant
    For Each edgeId In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With target.Borders(edgeId)
            .LineStyle = xlLineStyleNone
        End With
    Next edgeId
End Sub

Private Sub ApplyThinGridBorder(ByVal target As Range)
    With target.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With
End Sub

Private Sub ClearOutputTotalPairFormatting(ByVal ws As Worksheet, _
                                           ByVal rowIndex As Long, _
                                           ByVal labelColumn As Long, _
                                           ByVal sumColumn As Long, _
                                           ByVal currentTotalRow As Long, _
                                           ByVal lastDataRow As Long, _
                                           ByVal seiriColumn As Long)
    If rowIndex = currentTotalRow Then Exit Sub
    If Not IsOutputTotalLabelText(CommonNzText(ws.Cells(rowIndex, labelColumn).value)) Then Exit Sub

    ClearDoubleBorder ws.Cells(rowIndex, labelColumn)
    ClearDoubleBorder ws.Cells(rowIndex, sumColumn)
    ws.Cells(rowIndex, labelColumn).ClearContents
    ws.Cells(rowIndex, sumColumn).ClearContents
End Sub

Private Sub ClearStaleOutputTotalFormatting(ByVal ws As Worksheet, _
                                            ByVal lastDataRow As Long, _
                                            ByVal totalRow As Long, _
                                            ByVal subconFirstCol As Long, _
                                            ByVal subconColumnCount As Long, _
                                            Optional ByVal seiriColumn As Long = 0, _
                                            Optional ByVal jrLabelColumn As Long = 0, _
                                            Optional ByVal jrSumColumn As Long = 0)
    Dim scanLastRow As Long
    scanLastRow = lastDataRow + 5
    If Not ws.UsedRange Is Nothing Then
        Dim usedLastRow As Long
        usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If usedLastRow > scanLastRow Then scanLastRow = usedLastRow
    End If

    If seiriColumn = 0 Then seiriColumn = OutputSheetSeiriColumn(ws)
    If jrLabelColumn = 0 Then jrLabelColumn = OutputSheetCol(ws, COL_JR_PRICE)
    If jrSumColumn = 0 Then jrSumColumn = OutputSheetCol(ws, COL_JR_AMOUNT)

    Dim rowIndex As Long
    For rowIndex = 2 To scanLastRow
        ClearOutputTotalPairFormatting ws, rowIndex, jrLabelColumn, jrSumColumn, _
            totalRow, lastDataRow, seiriColumn

        If subconColumnCount > 0 Then
            Dim colIndex As Long
            For colIndex = subconFirstCol To subconFirstCol + subconColumnCount - 1 Step 2
                ClearOutputTotalPairFormatting ws, rowIndex, colIndex, colIndex + 1, _
                    totalRow, lastDataRow, seiriColumn
            Next colIndex
        End If
    Next rowIndex
End Sub

Private Sub DrawDoubleBorder(ByVal target As Range, ByVal lineColor As Long)
    Dim edgeId As Variant
    For Each edgeId In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With target.Borders(edgeId)
            .LineStyle = xlDouble
            .Color = lineColor
        End With
    Next edgeId
End Sub

Private Sub RedrawTotalBorders(ByVal ws As Worksheet, ByVal totalRow As Long, _
                               ByVal labelColumn As Long, ByVal sumColumn As Long)
    DrawDoubleBorder ws.Cells(totalRow, labelColumn), RGB(0, 0, 0)
    DrawDoubleBorder ws.Cells(totalRow, sumColumn), RGB(255, 0, 0)
End Sub

Private Sub WriteOutputTotalRows( _
    ByVal ws As Worksheet, _
    ByVal vendorNames As Collection, _
    Optional ByVal subconFirstCol As Long = 0, _
    Optional ByVal subconColumnCount As Long = -1, _
    Optional ByVal dataKeyColumn As Long = 0, _
    Optional ByVal jrLabelColumn As Long = 0, _
    Optional ByVal jrSumColumn As Long = 0)

    If dataKeyColumn = 0 Then dataKeyColumn = OutputSheetSeiriColumn(ws)
    If jrLabelColumn = 0 Then jrLabelColumn = OutputSheetCol(ws, COL_JR_PRICE)
    If jrSumColumn = 0 Then jrSumColumn = OutputSheetCol(ws, COL_JR_AMOUNT)
    If subconFirstCol = 0 Then subconFirstCol = OutputSheetSubconPriceFirstCol(ws)
    If subconColumnCount < 0 Then
        If vendorNames Is Nothing Then
            subconColumnCount = 0
        Else
            subconColumnCount = vendorNames.Count * 2
        End If
    End If

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, dataKeyColumn)
    If lastRow < 2 Then Exit Sub

    ClearStaleOutputTotalFormatting ws, lastRow, lastRow + 1, subconFirstCol, _
        subconColumnCount, dataKeyColumn, jrLabelColumn, jrSumColumn

    WriteTotalCells ws, lastRow + 1, jrLabelColumn, "JR合計", jrSumColumn, lastRow

    If Not vendorNames Is Nothing Then
        Dim vendorIndex As Long
        Dim priceColumn As Long
        For vendorIndex = 1 To vendorNames.Count
            priceColumn = subconFirstCol + ((vendorIndex - 1) * 2)
            WriteTotalCells ws, lastRow + 1, _
                            priceColumn, CStr(vendorNames(vendorIndex)) & "合計", _
                            priceColumn + 1, lastRow
        Next vendorIndex
    End If
End Sub

Private Sub WriteJrTotalRow(ByVal ws As Worksheet)
    Dim emptyVendors As New Collection
    WriteOutputTotalRows ws, emptyVendors, 0, 0
End Sub

Private Sub WritePurchaseNoticeJrTotalRow(ByVal ws As Worksheet)
    Dim emptyVendors As New Collection
    WriteOutputTotalRows ws, emptyVendors, 0, 0, _
        PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_JR_AMOUNT_COL
End Sub

Private Sub FillReferenceUnitPrices(ByVal ws As Worksheet, _
                                    ByVal guidanceDocumentName As String, _
                                    Optional ByVal isWelding As Boolean = False)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim lineSheetMap As Object
    Set lineSheetMap = BuildConstructionLineSheetMap()

    Dim sheetPriceCaches As Object
    Set sheetPriceCaches = CreateObject("Scripting.Dictionary")
    sheetPriceCaches.CompareMode = vbTextCompare

    Dim weldingPriceSheetName As String
    If isWelding Then
        weldingPriceSheetName = ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then
            LogCI "レール溶接単価: シート名の解決に失敗(基本情報B6/C6 または 単価適用線区マスタ未一致)"
        ElseIf Not SheetExistsByName(weldingPriceSheetName) Then
            LogCI "レール溶接単価: シート「" & weldingPriceSheetName & "」が見つかりません"
        End If
    End If

    Dim matchedCount As Long, unresolvedLineCount As Long, missingRecordCount As Long
    Dim seiriColumn As Long
    Dim dayNightColumn As Long
    Dim qtyColumn As Long
    Dim autoPriceColumn As Long
    Dim autoAmountColumn As Long
    Dim compareColumn As Long
    seiriColumn = OutputSheetSeiriColumn(ws)
    dayNightColumn = OutputSheetCol(ws, COL_DAYNIGHT)
    qtyColumn = OutputSheetCol(ws, COL_QTY)
    autoPriceColumn = OutputSheetCol(ws, COL_AUTO_PRICE)
    autoAmountColumn = OutputSheetCol(ws, COL_AUTO_AMOUNT)
    compareColumn = OutputSheetCol(ws, COL_PRICE_COMPARE)

    Dim r As Long
    For r = 2 To lastRow
        Dim unitPriceSheetName As String
        Dim recordKey As String
        If isWelding Then
            unitPriceSheetName = weldingPriceSheetName
            recordKey = BuildWeldingLookupKey(ws.Cells(r, seiriColumn).value)
        Else
            unitPriceSheetName = ResolveUnitPriceSheetName(lineSheetMap, CommonNzText(ws.Cells(r, OutputSheetCol(ws, COL_LINE)).value))
            recordKey = NormalizeRecordKey(ws.Cells(r, seiriColumn).value)
        End If

        Dim referencePrice As Variant
        referencePrice = Empty

        If unitPriceSheetName = "" Then
            unresolvedLineCount = unresolvedLineCount + 1
        Else
            Dim priceRows As Object
            Set priceRows = GetUnitPriceRows(unitPriceSheetName, sheetPriceCaches)

            If priceRows Is Nothing Or recordKey = "" Then
                missingRecordCount = missingRecordCount + 1
            ElseIf priceRows.Exists(recordKey) Then
                Dim dayNightPrices As Variant
                dayNightPrices = priceRows(recordKey)
                referencePrice = SelectDayNightPrice(CommonNzText(ws.Cells(r, dayNightColumn).value), dayNightPrices)
                If Not IsEmpty(referencePrice) Then matchedCount = matchedCount + 1
            Else
                missingRecordCount = missingRecordCount + 1
            End If
        End If

        If IsEmpty(referencePrice) Or IsError(referencePrice) Then
            ws.Cells(r, autoPriceColumn).ClearContents
        Else
            ws.Cells(r, autoPriceColumn).value = referencePrice
        End If
        WritePriceComparison ws, r, unitPriceSheetName, True, guidanceDocumentName
    Next r

    ws.Range(ws.Cells(2, autoAmountColumn), ws.Cells(lastRow, autoAmountColumn)).FormulaR1C1 = _
        "=IF(OR(RC[" & (autoPriceColumn - autoAmountColumn) & "]="""",RC[" & (qtyColumn - autoAmountColumn) & "]=""""),"""",RC[" & (autoPriceColumn - autoAmountColumn) & "]*RC[" & (qtyColumn - autoAmountColumn) & "])"

    With ws.Range(ws.Cells(2, autoPriceColumn), ws.Cells(lastRow, autoAmountColumn))
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
    End With
    ws.Range(ws.Cells(1, compareColumn), ws.Cells(lastRow, compareColumn)).HorizontalAlignment = xlCenter

    LogCI "参照単価一致=" & matchedCount & _
          " / 線区未解決=" & unresolvedLineCount & _
          " / 整理番号未一致=" & missingRecordCount
End Sub

Public Sub RefreshConstructionReferenceUnitPricesOnExistingSheets()
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Dim refreshedCount As Long
    For Each ws In ThisWorkbook.Worksheets
        If IsConstructionDocumentOutputSheet(ws) Then
            Dim guidanceDocumentName As String
            guidanceDocumentName = ResolveGuidanceDocumentNameFromOutputSheet(ws)
            FillReferenceUnitPrices ws, guidanceDocumentName, IsWeldingOutputSheet(ws)
            If IsWeldingOutputSheet(ws) Then
                ApplyPriceGuidanceColumnLayoutAtColumns _
                    ws, OutputSheetCol(ws, COL_PRICE_COMPARE), OutputSheetCol(ws, COL_PRICE_GUIDANCE)
            Else
                ApplyPriceGuidanceColumnLayout ws
            End If
            refreshedCount = refreshedCount + 1
        End If
    Next ws

    LogCI "既存施工指示書等シートの参照単価再読込: 対象=" & refreshedCount

Cleanup:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
End Sub

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
    If Not IsConstructionVendorOutputSheet(ws) Then Exit Function
    If FindHeaderColumn(ws, "整理番号") = 0 Then Exit Function
    IsConstructionDocumentOutputSheet = (FindHeaderColumn(ws, "単価比較") > 0)
End Function

Private Sub RefreshConstructionReferencePricesOnSheet( _
    ByVal ws As Worksheet, _
    ByVal unitPriceSheetName As String, _
    ByVal changedPriceRows As Object, _
    ByVal lineSheetMap As Object)

    ' (溶接)シートは取込時に溶接単価シートからM列を確定するため、
    ' 工事用単価シートの変更による参照単価の再計算対象から除外する。
    If Right$(ws.Name, Len(CONSTRUCTION_SHEET_SUFFIX_WELDING)) = CONSTRUCTION_SHEET_SUFFIX_WELDING Then Exit Sub

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
                lineSheetMap, CommonNzText(ws.Cells(r, FindHeaderColumn(ws, "契約線区名")).value))

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

Private Function BuildConstructionLineSheetMap() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

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
    If Err.Number <> 0 Then
        LogCI "線区名マスタ読込エラー Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection

    LogCI "線区名→単価シート対応数=" & result.Count & " master=[" & masterPath & "]"
    Err.Clear
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
    If lineSheetMap Is Nothing Then GoTo DirectLookup

    Dim key As String
    key = "E|" & NormalizeLineLookupText(importedLineName, False)
    If Len(key) > 2 And lineSheetMap.Exists(key) Then
        ResolveUnitPriceSheetName = CStr(lineSheetMap(key))
        Exit Function
    End If

    key = "S|" & NormalizeLineLookupText(importedLineName, True)
    If Len(key) > 2 And lineSheetMap.Exists(key) Then
        ResolveUnitPriceSheetName = CStr(lineSheetMap(key))
        Exit Function
    End If

DirectLookup:
    ResolveUnitPriceSheetName = FindImportedUnitPriceSheetName(importedLineName)
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
        Dim masterFolder As String
        masterFolder = fso.BuildPath(CStr(documentRoot), MASTER_DATA_FOLDER & "\" & lineType)
        LogCI "工事件名別マスタ探索 folder=[" & masterFolder & "]"
        ResolveProjectLineMasterPath = FindProjectMasterFile(masterFolder, projectName)
        If ResolveProjectLineMasterPath <> "" Then
            LogCI "工事件名別マスタ解決 path=[" & ResolveProjectLineMasterPath & "]"
            Exit Function
        End If
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
        If mod_WeldingUnitPrice.IsWeldingUnitPriceSheet(priceSheet) Then
            recordKey = BuildWeldingLookupKey(priceSheet.Cells(r, COL_SEIRI).value)
        Else
            recordKey = NormalizeRecordKey(priceSheet.Cells(r, COL_SEIRI).value)
        End If
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
    If Not IsArray(dayNightPrices) Then Exit Function

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
        ws, rowIndex, unitPriceSheetName, OutputSheetCol(ws, COL_AUTO_PRICE), _
        OutputSheetCol(ws, COL_JR_PRICE), OutputSheetCol(ws, COL_PRICE_COMPARE), _
        OutputSheetCol(ws, COL_PRICE_GUIDANCE), includeGuidance
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

Private Function ResolveGuidanceDocumentNameFromOutputSheet(ByVal ws As Worksheet) As String
    Dim mgrColumn As Long
    mgrColumn = FindHeaderColumn(ws, "管理室")
    If mgrColumn > 0 Then
        If ws.Columns(mgrColumn).Hidden Then
            ResolveGuidanceDocumentNameFromOutputSheet = "施工通知書"
            Exit Function
        End If
    End If

    Dim normalizedName As String
    normalizedName = CommonRemoveAllSpaces(CommonNormalizeText(ws.Name))
    If InStr(1, normalizedName, "通知", vbTextCompare) > 0 Then
        ResolveGuidanceDocumentNameFromOutputSheet = "施工通知書"
    Else
        ResolveGuidanceDocumentNameFromOutputSheet = "施工指示書"
    End If
End Function

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

Private Sub SortWorksSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim colSide As Long
    Dim colWeld As Long
    Dim colLine As Long
    Dim colKind As Long
    Dim colDayNight As Long
    Dim colSeiri As Long
    colSide = OutputSheetCol(ws, COL_FLAG_SIDE)
    colWeld = OutputSheetCol(ws, COL_FLAG_WELD)
    colLine = OutputSheetCol(ws, COL_LINE)
    colKind = OutputSheetCol(ws, COL_KIND)
    colDayNight = OutputSheetCol(ws, COL_DAYNIGHT)
    colSeiri = OutputSheetSeiriColumn(ws)

    Dim r As Long, lineName As String, kindName As String
    For r = 2 To lastRow
        lineName = CommonRemoveAllSpaces(CommonNzText(ws.Cells(r, colLine).value))
        kindName = CommonRemoveAllSpaces(CommonNzText(ws.Cells(r, colKind).value))
        ws.Cells(r, colSide).value = IIf(InStr(1, lineName, SIDELINE_KEYWORD) > 0, 1, 0)
        ws.Cells(r, colWeld).value = IIf(InStr(1, kindName, WELDING_KEYWORD) > 0, 1, 0)
    Next r

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(2, colSide), ws.Cells(lastRow, colSide)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, colLine), ws.Cells(lastRow, colLine)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, colWeld), ws.Cells(lastRow, colWeld)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, colKind), ws.Cells(lastRow, colKind)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, colDayNight), ws.Cells(lastRow, colDayNight)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, colSeiri), ws.Cells(lastRow, colSeiri)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, colWeld))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ws.Range(ws.Cells(1, colSide), ws.Cells(lastRow, colWeld)).ClearContents
End Sub

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

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, outputLastColumn))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, jrPriceColumn), _
                 ws.Cells(lastRow, outAmountColumn)).NumberFormatLocal = "#,##0"
    End If

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

    ws.Range(ws.Cells(1, 1), ws.Cells(1, outputLastColumn)).EntireColumn.AutoFit

    If CommonNzText(ws.Cells(1, autoAmountColumn).value) <> "" Then
        ws.Columns(autoPriceColumn).ShrinkToFit = False
        ws.Columns(autoPriceColumn).ColumnWidth = ws.Columns(autoAmountColumn).ColumnWidth
        ws.Cells(1, autoPriceColumn).ShrinkToFit = True
    End If

    ws.Cells.RowHeight = 18

    Dim centerCols As Variant, cv As Variant
    centerCols = Array(dataKeyColumn, dayNightColumn, unitColumn, lineColumn, mgrColumn, kindColumn)
    For Each cv In centerCols
        ws.Range(ws.Cells(1, CLng(cv)), ws.Cells(lastRow, CLng(cv))).HorizontalAlignment = xlCenter
    Next cv
End Sub

Private Sub RedrawOutputSheetDataBorders(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim seiriColumn As Long
    seiriColumn = FindHeaderColumn(ws, "整理番号")
    If seiriColumn = 0 Then Exit Sub

    Dim kindColumn As Long
    kindColumn = FindHeaderColumn(ws, "工種分類")
    If kindColumn = 0 Then Exit Sub

    Dim comparisonColumn As Long
    comparisonColumn = FindHeaderColumn(ws, "単価比較")

    Dim autoPriceColumn As Long
    Dim autoAmountColumn As Long
    autoPriceColumn = 0
    autoAmountColumn = 0
    If comparisonColumn >= 3 Then
        autoPriceColumn = comparisonColumn - 2
        autoAmountColumn = comparisonColumn - 1
    End If

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, seiriColumn)
    If lastRow < 1 Then lastRow = 1

    If autoPriceColumn > kindColumn + 1 Then
        ws.Range(ws.Cells(1, kindColumn + 1), _
                 ws.Cells(lastRow, autoPriceColumn - 1)).Borders.LineStyle = xlNone
    End If

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, kindColumn)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With

    Dim subconFirstCol As Long
    subconFirstCol = OutputSheetSubconPriceFirstCol(ws)
    If kindColumn > subconFirstCol Then
        With ws.Range(ws.Cells(1, subconFirstCol), ws.Cells(lastRow, kindColumn - 1)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With
    End If

    If autoPriceColumn > 0 And autoAmountColumn > 0 Then
        Dim borderLastColumn As Long
        borderLastColumn = autoAmountColumn
        If comparisonColumn > 0 Then borderLastColumn = comparisonColumn

        With ws.Range(ws.Cells(1, autoPriceColumn), ws.Cells(lastRow, borderLastColumn)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With

        If comparisonColumn > 0 Then
            ws.Range(ws.Cells(1, comparisonColumn + 1), _
                     ws.Cells(lastRow, comparisonColumn + 1)).Borders.LineStyle = xlNone
        End If
    End If
End Sub

Private Function GetWeldingPriceRowMap(ByVal weldingSheetName As String, _
                                       ByVal vendorPriceCaches As Object) As Object
    If vendorPriceCaches Is Nothing Then Exit Function

    Dim cacheKey As String
    cacheKey = "WELDING_ROWS|" & weldingSheetName
    If vendorPriceCaches.Exists(cacheKey) Then
        Set GetWeldingPriceRowMap = vendorPriceCaches(cacheKey)
        Exit Function
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(weldingSheetName)
    On Error GoTo 0
    If Not priceSheet Is Nothing Then
        Dim lastRow As Long
        lastRow = priceSheet.Cells(priceSheet.rows.Count, WELDING_PRICE_SEIRI_COL).End(xlUp).Row

        Dim r As Long
        For r = WELDING_PRICE_DATA_START_ROW To lastRow
            Dim lookupKey As String
            lookupKey = BuildWeldingLookupKey(priceSheet.Cells(r, WELDING_PRICE_SEIRI_COL).value)
            If lookupKey <> "" And Not result.Exists(lookupKey) Then
                result.Add lookupKey, r
            End If
        Next r
    End If

    vendorPriceCaches.Add cacheKey, result
    Set GetWeldingPriceRowMap = result
End Function

Private Function FindWeldingRailVendorDayColumn(ByVal priceSheet As Worksheet, _
                                                  ByVal vendorName As String) As Long
    Dim vendorKey As String
    vendorKey = NormalizeVendorPriceName(vendorName)
    If vendorKey = "" Then Exit Function

    Dim lastColumn As Long
    lastColumn = priceSheet.Cells(WELDING_PRICE_VENDOR_NAME_ROW, priceSheet.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = WELDING_PRICE_FIRST_RAIL_DAY_COL To lastColumn Step 2
        Dim headerName As String
        headerName = NormalizeVendorPriceName(CommonNzText(priceSheet.Cells(WELDING_PRICE_VENDOR_NAME_ROW, c).value))
        If headerName <> "" Then
            If StrComp(headerName, vendorKey, vbTextCompare) = 0 Then
                FindWeldingRailVendorDayColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function LookupWeldingOutputVendorPrice( _
    ByVal weldingSheetName As String, _
    ByVal vendorPriceCaches As Object, _
    ByVal recordKey As String, _
    ByVal vendorName As String, _
    ByVal isWeldingVendorSlot As Boolean, _
    ByVal dayNightText As String) As Variant

    LookupWeldingOutputVendorPrice = Empty
    If weldingSheetName = "" Or recordKey = "" Then Exit Function

    Dim rowMap As Object
    Set rowMap = GetWeldingPriceRowMap(weldingSheetName, vendorPriceCaches)
    If rowMap Is Nothing Or Not rowMap.Exists(recordKey) Then Exit Function

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(weldingSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then Exit Function

    Dim priceRow As Long
    priceRow = CLng(rowMap(recordKey))

    Dim dayColumn As Long
    If isWeldingVendorSlot Then
        dayColumn = WELDING_PRICE_WELDING_DAY_COL
    Else
        dayColumn = FindWeldingRailVendorDayColumn(priceSheet, vendorName)
        If dayColumn = 0 Then Exit Function
    End If

    LookupWeldingOutputVendorPrice = SelectDayNightPrice(dayNightText, _
        Array(priceSheet.Cells(priceRow, dayColumn).Value2, priceSheet.Cells(priceRow, dayColumn + 1).Value2))
End Function

Private Sub ApplyOutputSheetHeaderAutoFilter( _
    ByVal ws As Worksheet, _
    Optional ByVal dataKeyColumn As Long = 0, _
    Optional ByVal kindColumn As Long = COL_KIND, _
    Optional ByVal autoAmountColumn As Long = COL_AUTO_AMOUNT, _
    Optional ByVal comparisonColumn As Long = COL_PRICE_COMPARE, _
    Optional ByVal guidanceColumn As Long = COL_PRICE_GUIDANCE)

    If ws Is Nothing Then Exit Sub

    If dataKeyColumn = 0 Then dataKeyColumn = OutputSheetSeiriColumn(ws)

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, dataKeyColumn)
    If lastRow < 1 Then lastRow = 1

    Dim lastCol As Long
    lastCol = GetOutputLastColumn(ws, kindColumn, autoAmountColumn, comparisonColumn, guidanceColumn)

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
End Sub

Private Sub ApplyOutputSheetHeaderFreezePanes(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim prevSheet As Worksheet
    On Error Resume Next
    Set prevSheet = ActiveSheet
    On Error GoTo 0

    ws.Activate
    With ws.Application.ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        ws.Range("A2").Select
        .FreezePanes = True
    End With

    On Error Resume Next
    If Not prevSheet Is Nothing Then prevSheet.Activate
    On Error GoTo 0
End Sub

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

Private Sub ApplyPurchaseNoticeColumnExclusions(ByVal ws As Worksheet)
    Dim excludedColumns As Range
    Set excludedColumns = Application.Union(ws.Columns(PURCHASE_NOTICE_MGR_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_PRICE_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_AMOUNT_COL))
    excludedColumns.EntireColumn.Hidden = True
End Sub

Private Function GetLastDataRow(ByVal ws As Worksheet, _
                                Optional ByVal dataKeyColumn As Long = 0) As Long
    If dataKeyColumn = 0 Then dataKeyColumn = OutputSheetSeiriColumn(ws)

    Dim scanStartRow As Long
    scanStartRow = 2
    If Not ws.UsedRange Is Nothing Then
        Dim usedLastRow As Long
        usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If usedLastRow > scanStartRow Then scanStartRow = usedLastRow
    End If

    Dim rowIndex As Long
    For rowIndex = scanStartRow To 2 Step -1
        If Trim$(CommonNzText(ws.Cells(rowIndex, dataKeyColumn).value)) <> "" Then
            GetLastDataRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    GetLastDataRow = 1
End Function

Public Function IsWeldingOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsWeldingOutputSheet = (FindHeaderColumn(ws, WELDING_VENDOR_HEADER) > 0)
End Function

Public Function OutputSheetCol(ByVal ws As Worksheet, ByVal baseCol As Long) As Long
    If ws Is Nothing Then
        OutputSheetCol = baseCol + WELDING_OUTPUT_COL_OFFSET
    ElseIf IsWeldingOutputSheet(ws) Then
        OutputSheetCol = baseCol + WELDING_OUTPUT_COL_OFFSET
    Else
        OutputSheetCol = baseCol
    End If
End Function

Public Function OutputSheetSeiriColumn(ByVal ws As Worksheet) As Long
    OutputSheetSeiriColumn = OutputSheetCol(ws, COL_SEIRI)
End Function

Public Function OutputSheetSubconPriceFirstCol(ByVal ws As Worksheet) As Long
    Dim jrAmountCol As Long
    jrAmountCol = FindHeaderColumn(ws, "JR金額")
    If jrAmountCol > 0 Then
        OutputSheetSubconPriceFirstCol = jrAmountCol + 1
        Exit Function
    End If

    If IsWeldingOutputSheet(ws) Then
        OutputSheetSubconPriceFirstCol = WELDING_SUBCON_PRICE_FIRST_COL
    Else
        OutputSheetSubconPriceFirstCol = SUBCON_PRICE_FIRST_COL
    End If
End Function

Public Function OutputSheetVendorColumns(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If IsWeldingOutputSheet(ws) Then
        result.Add WELD_COL_WELDING_VENDOR
        result.Add WELD_COL_TRACK_VENDOR
    Else
        result.Add COL_VENDOR
    End If
    Set OutputSheetVendorColumns = result
End Function

Private Function WeldingRowArrayIndex(ByVal baseCol As Long) As Long
    If baseCol = COL_VENDOR Then
        WeldingRowArrayIndex = WELD_COL_WELDING_VENDOR - 1
    Else
        WeldingRowArrayIndex = baseCol + WELDING_OUTPUT_COL_OFFSET - 1
    End If
End Function

Private Function IsConstructionVendorOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If FindHeaderColumn(ws, "施工業者") > 0 Then
        IsConstructionVendorOutputSheet = True
        Exit Function
    End If
    IsConstructionVendorOutputSheet = IsWeldingOutputSheet(ws)
End Function

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

Private Function BuildConstructionSheetName(ByVal sourceA3Text As String, _
                                            ByVal isWelding As Boolean) As String
    Dim suffix As String
    If isWelding Then
        suffix = CONSTRUCTION_SHEET_SUFFIX_WELDING
    Else
        suffix = CONSTRUCTION_SHEET_SUFFIX_WORKS
    End If

    Dim baseName As String
    baseName = SanitizeSheetName(sourceA3Text)
    If baseName = "" Then Exit Function

    Dim maxBaseLen As Long
    maxBaseLen = 31 - Len(suffix)
    If maxBaseLen < 1 Then maxBaseLen = 1
    If Len(baseName) > maxBaseLen Then baseName = Left$(baseName, maxBaseLen)

    BuildConstructionSheetName = baseName & suffix
End Function

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

Private Sub LogCI(ByVal msg As String)
    Debug.Print "[ConstructionImport] " & Format(Now, "hh:mm:ss") & "  " & msg
End Sub

Private Sub FillPurchaseUnitPrices(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, PURCHASE_NOTICE_SEIRI_COL).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

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

Private Function ResolveWeldingPriceSheetName() As String
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
    Dim recordset As Object
    actualSheetName = FindAdoWorksheetName(connection, PRICE_LINE_SHEET)

    Dim resultName As String
    If actualSheetName = "" Then
        LogCI "レール溶接単価: マスタに「" & PRICE_LINE_SHEET & "」シートがありません"
    Else
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
                    resultName = nameText & WELDING_PRICE_SHEET_SUFFIX
                    Exit Do
                End If
            End If
            recordset.MoveNext
        Loop
    End If

Cleanup:
    If Err.Number <> 0 Then
        LogCI "レール溶接単価マスタADO読込エラー Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    ResolveWeldingPriceSheetName = resultName
End Function

Private Function BuildWeldingLookupKey(ByVal seiriValue As Variant) As String
    Dim keyText As String
    keyText = CommonRemoveAllSpaces(CommonNzText(seiriValue))
    If keyText = "" Then Exit Function
    If IsNumeric(keyText) Then
        Dim seiriNumber As Long
        seiriNumber = CLng(CDbl(keyText))
        ' 5桁の整理番号は下4桁を溶接単価シートの整理番号として照合する。
        If seiriNumber >= 10000 And seiriNumber <= 99999 Then
            BuildWeldingLookupKey = Right$(Format$(seiriNumber, "00000"), 4)
        Else
            BuildWeldingLookupKey = CStr(seiriNumber)
        End If
    Else
        If Len(keyText) = 5 And IsNumeric(keyText) Then
            BuildWeldingLookupKey = Right$(keyText, 4)
        Else
            BuildWeldingLookupKey = keyText
        End If
    End If
End Function

Private Function SheetExistsByName(ByVal sheetName As String) As Boolean
    If sheetName = "" Then Exit Function
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.worksheets(sheetName)
    On Error GoTo 0
    SheetExistsByName = Not ws Is Nothing
End Function

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
            MASTER_DATA_FOLDER & "\" & UNIT_PRICE_LINE_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText candidates, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
            "線路出張所用_注文書_請求書アクセスサイト - ドキュメント\" & _
            MASTER_DATA_FOLDER & "\" & UNIT_PRICE_LINE_MASTER_FILE
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

Private Sub WriteReferenceValueToBasicInfo(ByVal refValue As Variant, ByVal sourceCell As String)
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
    LogCI BASIC_INFO_REF_VALUE_CELL & " に参照シート " & sourceCell & " を転記"
End Sub
