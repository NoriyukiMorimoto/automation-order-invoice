Option Explicit

Public SharedMasterData As Variant
Private mClearingImportedLineNames As Boolean
Private mImportingUnitPriceData As Boolean
' 直近の単価表取込1回分を識別するID。取込開始時に採番し、生成した各シートの
' XFD2セルへ書き込む。運用上「参照は常に最後に取り込んだ単価シートから」で
' あるため、以前の取込で作られた古いシートは重い再展開処理の対象から外す
' (mod_VendorMaster / mod_WeldingUnitPrice 側でこのバッチIDと突き合わせて判定)。
Private mCurrentImportBatchId As String

Private Type UnitPriceRequest
    Nendo As String
    BranchName As String
    OfficeName As String
    lineType As String
    projectName As String
    UnitPriceKind As String
End Type

Private Type UnitPriceMasterRow
    BranchName As String
    OfficeName As String
    BranchGroupName As String
    sectionName As String
    UnitPriceSectionName As String
End Type

Private Const MASTER_SHEET_NAME As String = "単価適用線区"
Private Const PROJECT_MASTER_SHEET_NAME As String = "単価適用工事件名マスタ"
Private Const ZAIRAISEN_DISTINCTION_SHEET_NAME As String = "在幹区分"
Private Const UNIT_PRICE_DATA_FOLDER As String = "単価データ"
Private Const UNIT_PRICE_MASTER_FILE As String = "出張所別_単価適用線区.xlsx"

Private Const ZAIRAISEN_NAME As String = "在来線"
Private Const SHINKANSEN_NAME As String = "新幹線"
Private Const INITIAL_PRICE_NAME As String = "年初単価"
Private Const DESIGN_CHANGE_PRICE_NAME As String = "設計変更単価"
Private Const NORMAL_PRICE_FOLDER As String = "通常単価"

Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_WORK_NAME_CELL As String = "C10"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C22"
Private Const BASIC_INFO_PRICE_KIND_FALLBACK_CELL As String = "B22"
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_CELL As String = "C24"
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_MONITOR_RANGE As String = "C24:C28"
Private Const BASIC_INFO_WELDING_FLAG_CELL As String = "C23"
' 施工会社ブロック最大10社(AH列まで)と衝突しないよう、AJ(業者一覧)の右隣から配置
Private Const PROJECT_NAME_LIST_COL As String = "AM"
Private Const LINE_TYPE_LIST_COL As String = "AN"
Private Const PRICE_KIND_LIST_COL As String = "AO"
Private Const LIST_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_LAST_ROW As Long = 1048576
Private Const IMPORTED_SHEET_MARKER_ADDRESS As String = "XFD1"
Private Const IMPORTED_SHEET_LEGACY_MARKER_ADDRESS As String = "ZZ1"
Private Const IMPORTED_SHEET_MARKER_VALUE As String = "UnitPriceImported"
Private Const IMPORTED_SHEET_BATCH_ADDRESS As String = "XFD2"
Private Const CURRENT_BATCH_DEFINED_NAME As String = "CurrentUnitPriceImportBatchId"

Private Const UNIT_PRICE_SHEET_TAB_R As Long = 165
Private Const UNIT_PRICE_SHEET_TAB_G As Long = 171
Private Const UNIT_PRICE_SHEET_TAB_B As Long = 229
Private Const UNIT_PRICE_FILE_KEYWORD As String = "軌道材料購入充当"
Private Const PURCHASE_SHEET_NAME_SUFFIX As String = "_購入充当単価"
Private Const PURCHASE_SHEET_TAB_R As Long = 153
Private Const PURCHASE_SHEET_TAB_G As Long = 255
Private Const PURCHASE_SHEET_TAB_B As Long = 204
Private Const WELDING_REQUIRED_VALUE As String = "溶接工事あり"
Private Const WELDING_FILE_KEYWORD As String = "⑧レール溶接工事"
Private Const WELDING_FILE_FALLBACK_KEYWORD As String = "レール溶接工事"
Private Const WELDING_SHEET_NAME_SUFFIX As String = "_レール溶接単価"
Private Const WELDING_SHEET_TAB_R As Long = 252
Private Const WELDING_SHEET_TAB_G As Long = 213
Private Const WELDING_SHEET_TAB_B As Long = 180
Private Const SOURCE_HEADER_ROW As Long = 1
Private Const UNIT_PRICE_BLANK_FILL_DATA_START_ROW As Long = 7
Private Const UNIT_PRICE_BLANK_FILL_LAST_ROW_COL As Long = 2
Private Const UNIT_PRICE_BLANK_FILL_COL_START As Long = 5
Private Const UNIT_PRICE_BLANK_FILL_COL_END As Long = 6
Private Const UNIT_PRICE_BLANK_FILL_COLOR_R As Long = 128
Private Const UNIT_PRICE_BLANK_FILL_COLOR_G As Long = 128
Private Const UNIT_PRICE_BLANK_FILL_COLOR_B As Long = 128
Private Const WELDING_BLANK_FILL_COL_I_START As Long = 9             ' I列以降の無値塗りつぶし開始列
Private Const WELDING_YEAR_HEADER_SKIP_ROWS As Long = 4              ' 年度ヘッダー行および直下のスキップ行数
Private Const IMPORTED_LINE_NAME_BASE_FONT_SIZE As Double = 16
Private Const IMPORTED_LINE_NAME_MIN_FONT_SIZE As Double = 5
Private Const IMPORTED_LINE_NAME_LINE_HEIGHT_RATIO As Double = 1.25
Private Const IMPORTED_LINE_NAME_VERTICAL_PADDING As Double = 4
Private Const IMPORTED_UNIT_PRICE_CENTER_COL As Long = 2
Private Const IMPORTED_UNIT_PRICE_MERGE_FIRST_ROW As Long = 1   ' A~F結合対象の開始行(1行目)
Private Const IMPORTED_UNIT_PRICE_MERGE_LAST_ROW As Long = 2    ' A~F結合対象の終了行(2行目)
Private Const IMPORTED_UNIT_PRICE_MERGE_FIRST_COL As Long = 1   ' A列
Private Const IMPORTED_UNIT_PRICE_MERGE_LAST_COL As Long = 6    ' F列
Private Const IMPORTED_UNIT_PRICE_COL_A_WIDTH As Double = 10#   ' A列のセル幅（工事単価・溶接単価）
Private Const PURCHASE_IMPORTED_UNIT_PRICE_COL_A_WIDTH As Double = 13#   ' A列のセル幅（購入充当単価）
Private Const IMPORTED_UNIT_PRICE_LEFT_ALIGN_ROW_FIRST As Long = 3  ' B3
Private Const IMPORTED_UNIT_PRICE_LEFT_ALIGN_ROW_LAST As Long = 4   ' B4

Private Const ZAIRAISEN_PROJECT_NAME_FIRST_END_ROW  As Long = 8
Private Const ZAIRAISEN_PROJECT_NAME_SKIP_ROW       As Long = 9
Private Const ZAIRAISEN_PROJECT_NAME_SECOND_END_ROW As Long = 11
Private Const SHINKANSEN_PROJECT_NAME_END_ROW       As Long = 6

Private Sub LogUP(ByVal msg As String)
    On Error Resume Next
    mod_DebugLog.Log "[UnitPrice] " & msg
    On Error GoTo 0
End Sub

Private Function LogUPB(ByVal msg As String, ByVal returnValue As Boolean) As Boolean
    LogUP msg
    LogUPB = returnValue
End Function

Public Function GetMasterFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim candidates As Collection
    Set candidates = New Collection

    If Len(ThisWorkbook.Path) > 0 Then
        candidates.Add fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                       MasterDataFolderText() & "\" & UNIT_PRICE_MASTER_FILE)
        candidates.Add fso.BuildPath(ThisWorkbook.Path, UNIT_PRICE_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(userProfilePath)) > 0 Then
        candidates.Add userProfilePath & "\" & CommonCompanyNameText() & "\" & _
                       OrderInvoiceDocumentFolderText() & "\" & MasterDataFolderText() & "\" & UNIT_PRICE_MASTER_FILE
    End If

    GetMasterFilePath = FirstExistingFilePath(candidates)
End Function

Public Sub AutoFillLineTypeFromWorkName(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim workName As String
    workName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_WORK_NAME_CELL).value))
    If workName = "" Then Exit Sub

    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath = "" Then Exit Sub

    Dim lineType As String
    lineType = LookupLineTypeByWorkName(masterFilePath, workName)
    If lineType = "" Then Exit Sub

    If StrComp(CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value)), lineType, vbTextCompare) = 0 Then Exit Sub

    wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value = lineType
End Sub

Public Sub AutoFillUnitPriceFieldsFromWorkName(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    AutoFillLineTypeFromWorkName wsInfo
    RefreshUnitPriceProjectNameValidation wsInfo, False
    AutoFillProjectNameFromWorkName wsInfo
End Sub

Public Sub AutoFillProjectNameFromWorkName(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim workName As String
    workName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_WORK_NAME_CELL).value))
    If workName = "" Then Exit Sub

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    If lineType = "" Then Exit Sub

    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath = "" Then Exit Sub

    Dim sheetName As String
    Dim dummy As Long
    ResolveUnitPriceProjectNameMasterConfig lineType, sheetName, dummy
    If sheetName = "" Then Exit Sub

    Dim projectName As String
    projectName = LookupProjectNameByWorkName(masterFilePath, sheetName, workName)
    If projectName = "" Then Exit Sub

    If StrComp(CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value)), projectName, vbTextCompare) = 0 Then Exit Sub

    wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value = projectName
End Sub

Private Function LookupProjectNameByWorkName(ByVal masterFilePath As String, _
                                             ByVal sheetName As String, _
                                             ByVal workName As String) As String
    Dim cn As Object
    Dim rs As Object

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(masterFilePath)
    If cn Is Nothing Then Exit Function

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(cn, sheetName)
    If actualSheetName = "" Then GoTo Cleanup

    Dim sql As String
    sql = "SELECT F1, F2 FROM " & _
          BuildAdoSheetTableName(actualSheetName) & _
          " WHERE F2 IS NOT NULL"

    Set rs = cn.Execute(sql)

    Dim normalizedWorkName As String
    normalizedWorkName = NormalizeMatchText(workName)

    Do Until rs.EOF
        Dim keyword As String
        keyword = NormalizeMatchText(CommonNzText(CommonGetAdoFieldValue(rs, 1)))
        If keyword <> "" Then
            If InStr(1, normalizedWorkName, keyword, vbTextCompare) > 0 Then
                Dim projectNameRaw As String
                projectNameRaw = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
                If projectNameRaw <> "" Then
                    LookupProjectNameByWorkName = CommonNormalizeText(projectNameRaw)
                    GoTo Cleanup
                End If
            End If
        End If
        rs.MoveNext
    Loop

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Resume Cleanup
End Function

Private Function LookupLineTypeByWorkName(ByVal masterFilePath As String, _
                                          ByVal workName As String) As String
    Dim cn As Object
    Dim rs As Object

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(masterFilePath)
    If cn Is Nothing Then Exit Function

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(cn, ZAIRAISEN_DISTINCTION_SHEET_NAME)
    If actualSheetName = "" Then GoTo Cleanup

    Dim sql As String
    sql = "SELECT F1, F2 FROM " & _
          BuildAdoSheetTableName(actualSheetName) & _
          " WHERE F1 IS NOT NULL"

    Set rs = cn.Execute(sql)

    Dim normalizedWorkName As String
    normalizedWorkName = NormalizeMatchText(workName)

    Do Until rs.EOF
        Dim keyword As String
        keyword = NormalizeMatchText(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If keyword <> "" Then
            If InStr(1, normalizedWorkName, keyword, vbTextCompare) > 0 Then
                Dim lineTypeRaw As String
                lineTypeRaw = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 1)))
                If lineTypeRaw <> "" Then
                    LookupLineTypeByWorkName = CommonNormalizeText(lineTypeRaw)
                    GoTo Cleanup
                End If
            End If
        End If
        rs.MoveNext
    Loop

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Resume Cleanup
End Function

Public Sub RefreshUnitPriceProjectNameValidation(Optional ByVal wsInfo As Worksheet, _
                                                  Optional ByVal keepProjectName As Boolean = True)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    WriteUnitPriceLineTypeValidation wsInfo
    WriteUnitPriceKindValidation wsInfo

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    Dim projectNames As Collection
    Set projectNames = LoadUnitPriceProjectNamesForBasicInfo(GetMasterFilePath(), lineType)
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

Public Sub ImportConstructionUnitPriceForBasicInfo()
    LogUP "ImportConstructionUnitPriceForBasicInfo 開始"
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        LogUP "基本情報シート取得失敗 -> 中断"
        MsgBox UiMsgBasicInfoSheetNotFoundText(), vbExclamation
        Exit Sub
    End If
    LogUP "基本情報シート取得成功 sheet=[" & wsInfo.Name & "]"

    ClearAndImportUnitPriceForBasicInfo wsInfo
End Sub

Public Sub ClearAndImportUnitPriceForBasicInfo(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    mImportingUnitPriceData = True
    On Error GoTo FinallyExit

    ' 既存の単価表/購入充当単価/レール溶接単価シートを残したまま新規取込みすると、
    ' シートが積み上がって書式(セルスタイル)数がExcelの上限に近づき、
    ' 新しいシートの.Copyが失敗する原因になる。運用上も参照は常に最後に
    ' 取り込んだ単価シートからのみ行うため、取込み前に必ずクリアする。
    LogUP "ClearAndImportUnitPriceForBasicInfo: 既存単価シートをクリア"
    SilentClearUnitPriceForBasicInfo ws

    ImportUnitPriceData ws
FinallyExit:
    mImportingUnitPriceData = False
End Sub

Public Function IsImportingUnitPriceData() As Boolean
    IsImportingUnitPriceData = mImportingUnitPriceData
End Function

Private Sub ImportUnitPriceData(ByVal wsInfo As Worksheet)
    LogUP "ImportUnitPriceData 開始 wb=[" & wsInfo.Parent.Name & "]"
    mCurrentImportBatchId = GenerateImportBatchId()
    LogUP "採番バッチID=[" & mCurrentImportBatchId & "]"

    Dim request As UnitPriceRequest
    If Not TryReadUnitPriceRequest(wsInfo, request) Then
        LogUP "TryReadUnitPriceRequest -> False 中断"
        Exit Sub
    End If
    LogUP "TryReadUnitPriceRequest -> True"

    Dim masterRow As UnitPriceMasterRow
    If Not TryLoadUnitPriceMasterRow(request, masterRow) Then
        LogUP "TryLoadUnitPriceMasterRow -> False 中断"
        Exit Sub
    End If
    LogUP "TryLoadUnitPriceMasterRow -> True BranchGroup=[" & masterRow.BranchGroupName & _
          "] UnitPriceSection=[" & masterRow.UnitPriceSectionName & "]"

    Dim priceFolderPath As String
    Dim sectionFolderPath As String
    Dim sourceFilePaths As Collection
    Set sourceFilePaths = ResolveUnitPriceSourceFilePaths(request, masterRow, priceFolderPath, sectionFolderPath)
    If sourceFilePaths Is Nothing Then
        LogUP "ResolveUnitPriceSourceFilePaths -> Nothing 中断"
        Exit Sub
    End If
    If sourceFilePaths.Count = 0 Then
        LogUP "ResolveUnitPriceSourceFilePaths -> 0件 中断"
        Exit Sub
    End If
    LogUP "ResolveUnitPriceSourceFilePaths: ブック数=" & CStr(sourceFilePaths.Count) & _
          " priceFolder=[" & priceFolderPath & "] sectionFolder=[" & sectionFolderPath & "]"

    Dim sheetSourceFileMap As Object
    Dim sheetSourceSheetMap As Object
    Dim sheetNames As Collection
    Set sheetNames = LoadWorksheetNameCandidatesFromWorkbooks(sourceFilePaths, sheetSourceFileMap, sheetSourceSheetMap)
    If sheetNames Is Nothing Then
        LogUP "LoadWorksheetNameCandidatesFromWorkbooks -> Nothing 中断"
        MsgBox UiMsgUnitPriceImportableSheetNotFoundText() & vbCrLf & JoinCollectionText(sourceFilePaths, vbCrLf), vbExclamation
        Exit Sub
    End If
    If sheetNames.Count = 0 Then
        LogUP "LoadWorksheetNameCandidatesFromWorkbooks -> 0件 中断"
        MsgBox UiMsgUnitPriceImportableSheetNotFoundText() & vbCrLf & JoinCollectionText(sourceFilePaths, vbCrLf), vbExclamation
        Exit Sub
    End If
    LogUP "LoadWorksheetNameCandidatesFromWorkbooks: 候補シート数=" & CStr(sheetNames.Count)

    ' C20=在来線の場合のみ、単価シート候補を「工事件名別マスタ D列(積算線区コード)の番号順」に
    ' 並べ替える。マスタD列に該当しないシートは取込対象から除外する。
    ' ここで並べ替えることで、選択ダイアログの表示順・取込順・C24の記載順がすべて揃う。
    If StrComp(request.lineType, ZAIRAISEN_NAME, vbTextCompare) = 0 Then
        Set sheetNames = OrderUnitPriceSheetNamesByProjectMasterFColumn(sheetNames, sheetSourceSheetMap)
        If sheetNames Is Nothing Or sheetNames.Count = 0 Then
            LogUP "D列コード順並べ替え後に対象シートが0件 -> 中断"
            MsgBox UiMsgUnitPriceImportableSheetNotFoundText() & vbCrLf & JoinCollectionText(sourceFilePaths, vbCrLf), vbExclamation
            Exit Sub
        End If
        LogUP "D列コード順並べ替え後の候補シート数=" & CStr(sheetNames.Count)
    End If

    Dim selectedSheetNames As Collection
    Set selectedSheetNames = PromptLineNameSelection(sheetNames)
    If selectedSheetNames Is Nothing Then
        LogUP "PromptLineNameSelection -> Nothing (キャンセル) 中断"
        Exit Sub
    End If
    If selectedSheetNames.Count = 0 Then
        LogUP "PromptLineNameSelection -> 0件 中断"
        Exit Sub
    End If
    LogUP "PromptLineNameSelection: 選択件数=" & CStr(selectedSheetNames.Count) & _
          " 選択=[" & JoinCollectionText(selectedSheetNames, "|") & "]"

    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation
    Dim previousEnableEvents As Boolean
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    previousEnableEvents = Application.EnableEvents
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ImportHeavyCleanup

    If Not ImportSelectedUnitPriceSheets(selectedSheetNames, wsInfo.Parent, sheetSourceFileMap, sheetSourceSheetMap) Then
        LogUP "ImportSelectedUnitPriceSheets -> False 中断"
        GoTo ImportHeavyCleanup
    End If
    LogUP "ImportSelectedUnitPriceSheets -> True"

    mod_VendorMaster.ApplyImportedUnitPriceJrHeadersForBasicInfo wsInfo

    WriteSelectedLineNames wsInfo, selectedSheetNames

    Dim purchaseSheetName As String
    LogUP "ImportPurchaseUnitPriceSheetsByReference 呼び出し"
    Call ImportPurchaseUnitPriceSheetsByReference(request, masterRow, priceFolderPath, sectionFolderPath, wsInfo.Parent, purchaseSheetName)
    LogUP "ImportPurchaseUnitPriceSheetsByReference 完了 createdSheet=[" & purchaseSheetName & "]"

    Dim weldingSheetName As String
    LogUP "ImportWeldingUnitPriceSheetsIfRequired 呼び出し"
    If Not ImportWeldingUnitPriceSheetsIfRequired(wsInfo, masterRow, priceFolderPath, sectionFolderPath, selectedSheetNames, wsInfo.Parent, weldingSheetName) Then
        LogUP "ImportWeldingUnitPriceSheetsIfRequired -> False 中断"
        GoTo ImportHeavyCleanup
    End If
    LogUP "ImportWeldingUnitPriceSheetsIfRequired -> True createdSheet=[" & weldingSheetName & "]"

    ' ここまでの単価表/購入充当単価/レール溶接単価シートがすべて作成できたので、
    ' このバッチを「現行」として確定する。以降の重い再展開処理はこのバッチの
    ' シートだけを対象にする。
    SetCurrentUnitPriceImportBatchId wsInfo.Parent, mCurrentImportBatchId
    LogUP "現行バッチID確定=[" & mCurrentImportBatchId & "]"

    LogUP "RefreshAllVendorUnitPricesForBasicInfo 開始"
    mod_VendorMaster.RefreshAllVendorUnitPricesForBasicInfo wsInfo, True
    LogUP "RefreshAllVendorUnitPricesForBasicInfo 完了"

    LogUP "ApplyWeldingVendorUnitPricesForBasicInfo 開始"
    mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, 0, True
    LogUP "ApplyWeldingVendorUnitPricesForBasicInfo 完了"

    mod_Construction_Order_Import.RefreshConstructionReferenceUnitPricesOnExistingSheets

    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0

    wsInfo.Activate
    LogUP "ImportUnitPriceData 完了"

    MsgBox BuildImportCompleteMessage(selectedSheetNames, JoinCollectionText(sourceFilePaths, vbCrLf), purchaseSheetName, weldingSheetName), vbInformation, UiMsgImportCompleteTitleText()

ImportHeavyCleanup:
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    Application.EnableEvents = previousEnableEvents
End Sub

Private Function TryReadUnitPriceRequest(ByVal wsInfo As Worksheet, ByRef request As UnitPriceRequest) As Boolean
    LogUP "TryReadUnitPriceRequest 開始"
    request.Nendo = CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).value))
    request.BranchName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    request.OfficeName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    request.lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
    request.projectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value))
    request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL).value))
    If request.UnitPriceKind = "" Then
        request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_FALLBACK_CELL).value))
    End If
    If InStr(1, NormalizeMatchText(request.UnitPriceKind), NormalizeMatchText("単価適用区分"), vbTextCompare) > 0 Then
        request.UnitPriceKind = ""
    End If

    If request.Nendo = "" Then
        LogUP "年度未取得 -> 中断"
        MsgBox UiMsgBasicInfoYearNotFoundB4Text(), vbExclamation
        Exit Function
    End If
    If request.BranchName = "" Or request.OfficeName = "" Then
        LogUP "支店/出張所未入力 Branch=[" & request.BranchName & "] Office=[" & request.OfficeName & "] -> 中断"
        MsgBox UiMsgBasicInfoBranchOfficeEmptyText(), vbExclamation
        Exit Function
    End If
    If request.lineType = "" Then
        LogUP "C20 線区区分未入力 -> 中断"
        MsgBox UiMsgBasicInfoLineTypeEmptyC20Text(), vbExclamation
        Exit Function
    End If
    If request.projectName = "" Then
        LogUP "C21 工事件名未入力 -> 中断"
        MsgBox UiMsgBasicInfoProjectNameEmptyC21Text(), vbExclamation
        Exit Function
    End If
    If IsPurchaseUnitPriceProjectName(request.projectName) Then
        LogUP "C21 が購入充当系工事件名 ProjectName=[" & request.projectName & "] -> 中断"
        MsgBox UiMsgBasicInfoProjectNamePurchaseExcludedC21Text() & vbCrLf & _
               UiMsgPurchaseUnitPriceAutoCreateAfterC24Text(), vbExclamation
        Exit Function
    End If
    If request.UnitPriceKind = "" Then
        LogUP "C22 単価区分未入力 -> 中断"
        MsgBox UiMsgBasicInfoPriceKindEmptyC22Text(), vbExclamation
        Exit Function
    End If

    LogUP "読み取り値 Nendo=[" & request.Nendo & "] Branch=[" & request.BranchName & _
          "] Office=[" & request.OfficeName & "] LineType=[" & request.lineType & _
          "] ProjectName=[" & request.projectName & "] PriceKind=[" & request.UnitPriceKind & "]"
    TryReadUnitPriceRequest = True
End Function

Private Function TryLoadUnitPriceMasterRow(ByRef request As UnitPriceRequest, _
                                           ByRef masterRow As UnitPriceMasterRow) As Boolean
    LogUP "TryLoadUnitPriceMasterRow 開始"
    Dim sourceFilePath As String
    sourceFilePath = GetMasterFilePath()
    LogUP "masterFilePath=[" & sourceFilePath & "]"
    If sourceFilePath = "" Then
        LogUP "masterFilePath 未取得 -> 中断"
        MsgBox UiMsgUnitPriceMasterFileNotFoundText(), vbExclamation
        Exit Function
    End If

    Dim cn As Object
    Dim rs As Object
    Dim adoErrDescription As String
    On Error GoTo ErrorHandler

    LogUP "ADO試行 path=[" & sourceFilePath & "]"
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then
        LogUP "ADO接続失敗 -> 中断"
        MsgBox UiMsgUnitPriceMasterFileUnreadableText() & vbCrLf & sourceFilePath, vbExclamation
        Exit Function
    End If
    LogUP "ADO接続成功"

    Dim masterSheetName As String
    masterSheetName = FindAdoWorksheetName(cn, MASTER_SHEET_NAME)
    LogUP "masterSheetName=[" & masterSheetName & "]"
    If masterSheetName = "" Then
        LogUP "単価適用線区シート未検出 -> 中断"
        MsgBox UiMsgUnitPriceMasterSheetNotFoundText() & vbCrLf & _
               UiMsgUnitPriceMasterSheetNameCheckText() & vbCrLf & sourceFilePath, vbExclamation
        GoTo Cleanup
    End If

    Dim sql As String
    sql = "SELECT F2, F3, F4, F5, F6 FROM " & _
          BuildAdoSheetTableName(masterSheetName) & _
          " WHERE F2 IS NOT NULL"
    Set rs = cn.Execute(sql)

    Dim matchedCount As Long
    Dim bestScore As Long
    Dim bestRow As UnitPriceMasterRow

    Do Until rs.EOF
        If MasterTextMatches(CommonNzText(CommonGetAdoFieldValue(rs, 0)), request.BranchName) And _
           MasterTextMatches(CommonNzText(CommonGetAdoFieldValue(rs, 1)), request.OfficeName) Then
            Dim candidateRow As UnitPriceMasterRow
            FillUnitPriceMasterRowFromRecordset rs, candidateRow

            matchedCount = matchedCount + 1

            Dim score As Long
            score = GetUnitPriceMasterRowScore(request, candidateRow)
            If score > bestScore Then
                bestScore = score
                bestRow = candidateRow
            End If
        End If
        rs.MoveNext
    Loop

    LogUP "マッチ結果 matchedCount=" & CStr(matchedCount) & " bestScore=" & CStr(bestScore)
    If matchedCount > 0 Then
        If bestScore > 1 Or matchedCount = 1 Then
            masterRow = bestRow
            LogUP "masterRow確定 BranchGroup=[" & masterRow.BranchGroupName & _
                  "] Section=[" & masterRow.sectionName & "] UnitPriceSection=[" & masterRow.UnitPriceSectionName & "]"
            TryLoadUnitPriceMasterRow = True
            GoTo Cleanup
        End If

        LogUP "線区区分に一致する保線区を特定できず -> 中断"
        MsgBox UiMsgUnitPriceMasterLineTypeAmbiguousText() & vbCrLf & _
               UiMsgBranchLabelText() & request.BranchName & vbCrLf & _
               UiMsgOfficeLabelText() & request.OfficeName & vbCrLf & _
               UiMsgLineTypeLabelText() & request.lineType, vbExclamation
        GoTo Cleanup
    End If

    LogUP "該当する支店・出張所なし -> 中断"
    MsgBox UiMsgUnitPriceMasterBranchOfficeNotFoundText() & vbCrLf & _
           UiMsgBranchLabelText() & request.BranchName & vbCrLf & _
           UiMsgOfficeLabelText() & request.OfficeName, vbExclamation

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    adoErrDescription = Err.Description
    LogUP "ADO例外発生 Err=[" & adoErrDescription & "] -> Workbookフォールバック試行"
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    If TryLoadUnitPriceMasterRowFromWorkbook(sourceFilePath, request, masterRow) Then
        LogUP "Workbookフォールバック成功"
        TryLoadUnitPriceMasterRow = True
        Exit Function
    End If
    LogUP "Workbookフォールバック失敗"
    MsgBox UiMsgUnitPriceMasterLoadFailedText() & vbCrLf & adoErrDescription, vbExclamation
    Resume Cleanup
End Function

Private Function TryLoadUnitPriceMasterRowFromWorkbook(ByVal sourceFilePath As String, _
                                                       ByRef request As UnitPriceRequest, _
                                                       ByRef masterRow As UnitPriceMasterRow) As Boolean
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set sourceSheet = FindWorksheetByName(sourceBook, MASTER_SHEET_NAME)
    If sourceSheet Is Nothing Then GoTo Cleanup

    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then GoTo Cleanup

    Dim matchedCount As Long
    Dim bestScore As Long
    Dim bestRow As UnitPriceMasterRow

    Dim rr As Long
    For rr = 2 To lastRow
        If MasterTextMatches(CommonNzText(sourceSheet.Cells(rr, 2).value), request.BranchName) And _
           MasterTextMatches(CommonNzText(sourceSheet.Cells(rr, 3).value), request.OfficeName) Then
            Dim candidateRow As UnitPriceMasterRow
            FillUnitPriceMasterRowFromWorksheet sourceSheet, rr, candidateRow

            matchedCount = matchedCount + 1

            Dim score As Long
            score = GetUnitPriceMasterRowScore(request, candidateRow)
            If score > bestScore Then
                bestScore = score
                bestRow = candidateRow
            End If
        End If
    Next rr

    If matchedCount > 0 Then
        If bestScore > 1 Or matchedCount = 1 Then
            masterRow = bestRow
            TryLoadUnitPriceMasterRowFromWorkbook = True
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    On Error GoTo 0
    Exit Function

ErrorHandler:
    TryLoadUnitPriceMasterRowFromWorkbook = False
    Resume Cleanup
End Function

Private Function FindWorksheetByName(ByVal sourceBook As Workbook, ByVal expectedSheetName As String) As Worksheet
    If sourceBook Is Nothing Then Exit Function

    Dim normalizedExpected As String
    normalizedExpected = NormalizeMatchText(expectedSheetName)

    Dim ws As Worksheet
    For Each ws In sourceBook.worksheets
        If StrComp(NormalizeMatchText(ws.Name), normalizedExpected, vbTextCompare) = 0 Then
            Set FindWorksheetByName = ws
            Exit Function
        End If
    Next ws

    For Each ws In sourceBook.worksheets
        If InStr(1, NormalizeMatchText(ws.Name), normalizedExpected, vbTextCompare) > 0 Then
            Set FindWorksheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindAdoWorksheetName(ByVal cn As Object, ByVal expectedSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)
    If sheetNames Is Nothing Then Exit Function

    Dim normalizedExpected As String
    normalizedExpected = NormalizeMatchText(expectedSheetName)

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(NormalizeMatchText(CStr(sheetName)), normalizedExpected, vbTextCompare) = 0 Then
            FindAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName

    For Each sheetName In sheetNames
        If InStr(1, NormalizeMatchText(CStr(sheetName)), normalizedExpected, vbTextCompare) > 0 Then
            FindAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

Private Function BuildAdoSheetTableName(ByVal sheetName As String) As String
    BuildAdoSheetTableName = "[" & Replace$(sheetName, "]", "]]") & "$]"
End Function

Private Sub FillUnitPriceMasterRowFromRecordset(ByVal rs As Object, _
                                                ByRef masterRow As UnitPriceMasterRow)
    masterRow.BranchName = CommonNzText(CommonGetAdoFieldValue(rs, 0))
    masterRow.OfficeName = CommonNzText(CommonGetAdoFieldValue(rs, 1))
    masterRow.BranchGroupName = CommonNzText(CommonGetAdoFieldValue(rs, 2))
    masterRow.sectionName = CommonNzText(CommonGetAdoFieldValue(rs, 3))
    masterRow.UnitPriceSectionName = CommonNzText(CommonGetAdoFieldValue(rs, 4))
End Sub

Private Sub FillUnitPriceMasterRowFromWorksheet(ByVal sourceSheet As Worksheet, _
                                                ByVal rowIndex As Long, _
                                                ByRef masterRow As UnitPriceMasterRow)
    masterRow.BranchName = CommonNzText(sourceSheet.Cells(rowIndex, 2).value)
    masterRow.OfficeName = CommonNzText(sourceSheet.Cells(rowIndex, 3).value)
    masterRow.BranchGroupName = CommonNzText(sourceSheet.Cells(rowIndex, 4).value)
    masterRow.sectionName = CommonNzText(sourceSheet.Cells(rowIndex, 5).value)
    masterRow.UnitPriceSectionName = CommonNzText(sourceSheet.Cells(rowIndex, 6).value)
End Sub

Private Function GetUnitPriceMasterRowScore(ByRef request As UnitPriceRequest, _
                                            ByRef candidateRow As UnitPriceMasterRow) As Long
    GetUnitPriceMasterRowScore = 1

    If UnitPriceMasterRowFolderExistsForLineType(request, candidateRow) Then
        GetUnitPriceMasterRowScore = 3
    ElseIf UnitPriceMasterRowTextMatchesLineType(request.lineType, candidateRow) Then
        GetUnitPriceMasterRowScore = 2
    End If
End Function

Private Function UnitPriceMasterRowFolderExistsForLineType(ByRef request As UnitPriceRequest, _
                                                           ByRef candidateRow As UnitPriceMasterRow) As Boolean
    Dim dataRoot As String
    dataRoot = GetUnitPriceDataRootPath()
    If dataRoot = "" Or Dir(dataRoot, vbDirectory) = "" Then Exit Function

    Dim lineFolder As String
    lineFolder = FindChildFolderByKey(dataRoot, request.lineType, True)
    If lineFolder = "" Then Exit Function

    Dim branchGroupFolder As String
    branchGroupFolder = FindChildFolderByKey(lineFolder, candidateRow.BranchGroupName, False)
    If branchGroupFolder = "" Then Exit Function

    UnitPriceMasterRowFolderExistsForLineType = _
        (FindChildFolderByKey(branchGroupFolder, candidateRow.UnitPriceSectionName, True) <> "")
End Function

Private Function UnitPriceMasterRowTextMatchesLineType(ByVal lineType As String, _
                                                       ByRef candidateRow As UnitPriceMasterRow) As Boolean
    Dim rowText As String
    rowText = NormalizeMatchText(candidateRow.sectionName & candidateRow.UnitPriceSectionName)

    Select Case CommonNormalizeText(lineType)
        Case SHINKANSEN_NAME
            UnitPriceMasterRowTextMatchesLineType = _
                (InStr(1, rowText, NormalizeMatchText(SHINKANSEN_NAME), vbTextCompare) > 0)
        Case ZAIRAISEN_NAME
            UnitPriceMasterRowTextMatchesLineType = _
                (InStr(1, rowText, NormalizeMatchText(SHINKANSEN_NAME), vbTextCompare) = 0)
        Case Else
            UnitPriceMasterRowTextMatchesLineType = True
    End Select
End Function

Private Function ResolveUnitPriceSourceFilePaths(ByRef request As UnitPriceRequest, _
                                                 ByRef masterRow As UnitPriceMasterRow, _
                                                 ByRef priceFolderPath As String, _
                                                 ByRef sectionFolderPath As String) As Collection
    priceFolderPath = ResolveUnitPricePriceFolderPath(request, masterRow, sectionFolderPath)
    If priceFolderPath = "" Then Exit Function

    Dim sourceFilePaths As Collection
    Set sourceFilePaths = FindUnitPriceWorkbooks(priceFolderPath, request.projectName)
    If sourceFilePaths Is Nothing Then
        LogUP "FindUnitPriceWorkbooks -> Nothing"
        Exit Function
    End If
    LogUP "FindUnitPriceWorkbooks: hits=" & CStr(sourceFilePaths.Count) & " projectName=[" & request.projectName & "]"
    If sourceFilePaths.Count = 0 Then
        LogUP "工事件名一致ブックなし -> 中断"
        MsgBox UiMsgUnitPriceBookByProjectNotFoundText() & vbCrLf & _
               UiMsgProjectNameLabelText() & request.projectName & vbCrLf & _
               priceFolderPath, vbExclamation
        Exit Function
    End If
    Dim sourceIndex As Long
    For sourceIndex = 1 To sourceFilePaths.Count
        LogUP "単価表ブック#" & CStr(sourceIndex) & "=[" & CStr(sourceFilePaths(sourceIndex)) & "]"
    Next sourceIndex
    Set ResolveUnitPriceSourceFilePaths = sourceFilePaths
End Function

Private Function ResolveUnitPricePriceFolderPath(ByRef request As UnitPriceRequest, _
                                                 ByRef masterRow As UnitPriceMasterRow, _
                                                 ByRef sectionFolderPath As String) As String
    LogUP "ResolveUnitPricePriceFolderPath 開始"
    Dim dataRoot As String
    dataRoot = GetUnitPriceDataRootPath()
    LogUP "dataRoot=[" & dataRoot & "]"
    If dataRoot = "" Or Dir(dataRoot, vbDirectory) = "" Then
        LogUP "単価データフォルダ未検出 -> 中断"
        MsgBox UiMsgUnitPriceDataFolderNotFoundText() & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If

    Dim lineFolder As String
    lineFolder = FindChildFolderByKey(dataRoot, request.lineType, True)
    LogUP "lineFolder=[" & lineFolder & "]"
    If lineFolder = "" Then
        LogUP "線区区分フォルダ未検出 key=[" & request.lineType & "] -> 中断"
        MsgBox UiMsgLineTypeFolderNotFoundText() & vbCrLf & request.lineType & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If

    Dim branchGroupFolder As String
    branchGroupFolder = FindChildFolderByKey(lineFolder, masterRow.BranchGroupName, False)
    LogUP "branchGroupFolder=[" & branchGroupFolder & "]"
    If branchGroupFolder = "" Then
        LogUP "支社フォルダ未検出 key=[" & masterRow.BranchGroupName & "] -> 中断"
        MsgBox UiMsgBranchGroupFolderNotFoundText() & vbCrLf & masterRow.BranchGroupName & vbCrLf & lineFolder, vbExclamation
        Exit Function
    End If

    Dim sectionFolder As String
    sectionFolder = FindChildFolderByKey(branchGroupFolder, masterRow.UnitPriceSectionName, True)
    LogUP "sectionFolder=[" & sectionFolder & "]"
    If sectionFolder = "" Then
        LogUP "保線区フォルダ未検出 key=[" & masterRow.UnitPriceSectionName & "] -> 中断"
        MsgBox UiMsgUnitPriceSectionFolderNotFoundText() & vbCrLf & masterRow.UnitPriceSectionName & vbCrLf & branchGroupFolder, vbExclamation
        Exit Function
    End If

    Dim yearFolder As String
    yearFolder = sectionFolder & "\" & request.Nendo
    LogUP "yearFolder=[" & yearFolder & "]"
    If Dir(yearFolder, vbDirectory) = "" Then
        LogUP "年度フォルダ未検出 -> 中断"
        MsgBox UiMsgYearFolderNotFoundText() & vbCrLf & yearFolder, vbExclamation
        Exit Function
    End If

    Dim priceFolder As String
    priceFolder = yearFolder & "\" & ResolveUnitPriceFolderName(request.UnitPriceKind)
    LogUP "priceFolder=[" & priceFolder & "]"
    If Dir(priceFolder, vbDirectory) = "" Then
        LogUP "単価区分フォルダ未検出 -> 中断"
        MsgBox UiMsgPriceKindFolderNotFoundText() & vbCrLf & priceFolder, vbExclamation
        Exit Function
    End If

    sectionFolderPath = sectionFolder
    ResolveUnitPricePriceFolderPath = priceFolder
End Function

Private Function ResolveUnitPriceFolderName(ByVal priceKind As String) As String
    Dim normalizedKind As String
    normalizedKind = NormalizeMatchText(priceKind)
    If InStr(1, normalizedKind, NormalizeMatchText(INITIAL_PRICE_NAME), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = NORMAL_PRICE_FOLDER
    ElseIf InStr(1, normalizedKind, NormalizeMatchText(DESIGN_CHANGE_PRICE_NAME), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = DESIGN_CHANGE_PRICE_NAME
    ElseIf InStr(1, normalizedKind, NormalizeMatchText(NORMAL_PRICE_FOLDER), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = NORMAL_PRICE_FOLDER
    Else
        ResolveUnitPriceFolderName = priceKind
    End If
End Function

Private Function NormalizeWorkbookPathForCompare(ByVal filePath As String) As String
    Dim trimmedPath As String
    trimmedPath = Trim$(filePath)
    If trimmedPath = "" Then Exit Function

    On Error GoTo Fallback
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(trimmedPath) Then
        NormalizeWorkbookPathForCompare = LCase$(fso.GetAbsolutePathName(trimmedPath))
        Exit Function
    End If

Fallback:
    NormalizeWorkbookPathForCompare = LCase$(Replace$(trimmedPath, "/", "\"))
End Function

Private Function FindOpenWorkbookByPath(ByVal sourceFilePath As String) As Workbook
    Dim normalizedPath As String
    normalizedPath = NormalizeWorkbookPathForCompare(sourceFilePath)
    If normalizedPath = "" Then Exit Function

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If NormalizeWorkbookPathForCompare(wb.FullName) = normalizedPath Then
            Set FindOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
End Function

Private Function LoadWorksheetNamesFromWorkbookByAdo(ByVal sourceFilePath As String) As Collection
    Dim cn As Object

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Set LoadWorksheetNamesFromWorkbookByAdo = CommonGetAdoWorksheetNames(cn)

Cleanup:
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Set LoadWorksheetNamesFromWorkbookByAdo = Nothing
    Resume Cleanup
End Function

Private Function LoadWorksheetNamesFromWorkbookByExcel(ByVal sourceFilePath As String) As Collection
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim result As Collection
    Dim openedByThisCall As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.screenUpdating
    Application.DisplayAlerts = False
    Application.screenUpdating = False

    Set sourceBook = FindOpenWorkbookByPath(sourceFilePath)
    If sourceBook Is Nothing Then
        Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
        openedByThisCall = True
    End If
    If sourceBook Is Nothing Then Err.Raise 91

    Set result = New Collection
    For Each sourceSheet In sourceBook.Worksheets
        result.Add sourceSheet.Name
    Next sourceSheet
    If result.Count = 0 Then Err.Raise 91

    Set LoadWorksheetNamesFromWorkbookByExcel = result

Cleanup:
    On Error Resume Next
    If openedByThisCall Then
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = previousDisplayAlerts
    Application.screenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Set LoadWorksheetNamesFromWorkbookByExcel = Nothing
    MsgBox UiMsgUnitPriceBookOpenFailedText() & vbCrLf & sourceFilePath & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function LoadWorksheetNamesFromWorkbook(ByVal sourceFilePath As String) As Collection
    Dim result As Collection

    If Len(Trim$(sourceFilePath)) = 0 Then Exit Function

    Set result = LoadWorksheetNamesFromWorkbookByAdo(sourceFilePath)
    If Not result Is Nothing Then
        If result.Count > 0 Then
            Set LoadWorksheetNamesFromWorkbook = result
            Exit Function
        End If
    End If

    Set LoadWorksheetNamesFromWorkbook = LoadWorksheetNamesFromWorkbookByExcel(sourceFilePath)
End Function

Private Function LoadWorksheetNameCandidatesFromWorkbooks(ByVal sourceFilePaths As Collection, _
                                                          ByRef sheetSourceFileMap As Object, _
                                                          ByRef sheetSourceSheetMap As Object) As Collection
    If sourceFilePaths Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection
    Set sheetSourceFileMap = CreateObject("Scripting.Dictionary")
    Set sheetSourceSheetMap = CreateObject("Scripting.Dictionary")

    Dim sourceFilePath As Variant
    For Each sourceFilePath In sourceFilePaths
        Dim workbookSheetNames As Collection
        Set workbookSheetNames = LoadWorksheetNamesFromWorkbook(CStr(sourceFilePath))
        If workbookSheetNames Is Nothing Then Exit Function

        Dim sheetName As Variant
        For Each sheetName In workbookSheetNames
            Dim displayName As String
            displayName = BuildUniqueUnitPriceSheetDisplayName(result, CStr(sheetName), CStr(sourceFilePath))
            result.Add displayName
            sheetSourceFileMap.Add displayName, CStr(sourceFilePath)
            sheetSourceSheetMap.Add displayName, CStr(sheetName)
        Next sheetName
    Next sourceFilePath

    Set LoadWorksheetNameCandidatesFromWorkbooks = result
End Function

Private Function BuildUniqueUnitPriceSheetDisplayName(ByVal currentNames As Collection, _
                                                      ByVal sheetName As String, _
                                                      ByVal sourceFilePath As String) As String
    Dim candidate As String
    candidate = sheetName
    If Not CollectionContainsText(currentNames, candidate) Then
        BuildUniqueUnitPriceSheetDisplayName = candidate
        Exit Function
    End If

    Dim fileBaseName As String
    fileBaseName = RemoveFileExtension(GetPathBaseName(sourceFilePath))
    candidate = sheetName & "（" & fileBaseName & "）"
    If Not CollectionContainsText(currentNames, candidate) Then
        BuildUniqueUnitPriceSheetDisplayName = candidate
        Exit Function
    End If

    Dim i As Long
    For i = 2 To 99
        candidate = sheetName & "（" & fileBaseName & " " & CStr(i) & "）"
        If Not CollectionContainsText(currentNames, candidate) Then
            BuildUniqueUnitPriceSheetDisplayName = candidate
            Exit Function
        End If
    Next i

    BuildUniqueUnitPriceSheetDisplayName = sheetName
End Function

Private Function PromptLineNameSelection(ByVal sheetNames As Collection) As Collection
    On Error GoTo ErrorHandler

    SelectLineName.InitLineNames sheetNames
    SelectLineName.Show vbModal

    If SelectLineName.SelectionConfirmed Then
        Set PromptLineNameSelection = SelectLineName.GetSelectedLineNames()
    End If

Cleanup:
    On Error Resume Next
    Unload SelectLineName
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox UiMsgLineNameFormShowFailedText() & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportSelectedUnitPriceSheets(ByVal selectedSheetNames As Collection, _
                                               ByVal targetBook As Workbook, _
                                               ByVal sheetSourceFileMap As Object, _
                                               ByVal sheetSourceSheetMap As Object) As Boolean
    Dim sourceBook As Workbook
    Dim stagedSheets As Collection
    Dim targetSheetNames As Collection
    Dim openedSourceBooks As Object
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    Set stagedSheets = New Collection
    Set targetSheetNames = New Collection
    Set openedSourceBooks = CreateObject("Scripting.Dictionary")
    openedSourceBooks.CompareMode = vbTextCompare

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim displayName As Variant
    For Each displayName In selectedSheetNames
        Dim sourceFilePath As String
        Dim sourceSheetName As String
        Dim targetSheetName As String
        Dim stagedSheet As Worksheet
        sourceFilePath = CStr(sheetSourceFileMap(CStr(displayName)))
        sourceSheetName = CStr(sheetSourceSheetMap(CStr(displayName)))
        targetSheetName = CStr(displayName)
        LogUP "シートコピー displayName=[" & targetSheetName & "] srcSheet=[" & sourceSheetName & _
              "] srcFile=[" & sourceFilePath & "]"

        If openedSourceBooks.Exists(sourceFilePath) Then
            Set sourceBook = openedSourceBooks(sourceFilePath)
        Else
            Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
            openedSourceBooks.Add sourceFilePath, sourceBook
        End If

        sourceBook.worksheets(sourceSheetName).Copy After:=targetBook.worksheets(targetBook.worksheets.Count)
        Set stagedSheet = targetBook.worksheets(targetBook.worksheets.Count)
        stagedSheet.Name = MakeUniqueWorksheetName(targetBook, "__UP_NEW_" & CStr(stagedSheets.Count + 1), stagedSheet.Name)
        With stagedSheet
            .Tab.Color = RGB(UNIT_PRICE_SHEET_TAB_R, UNIT_PRICE_SHEET_TAB_G, UNIT_PRICE_SHEET_TAB_B)
            MarkImportedUnitPriceSheet stagedSheet
            FillBlankUnitPriceEFCells stagedSheet
            ApplyImportedUnitPriceSheetFormat stagedSheet
            LogUP "一時シート作成完了 name=[" & .Name & "]"
        End With
        stagedSheets.Add stagedSheet
        targetSheetNames.Add targetSheetName
    Next displayName

    LogUP "ImportSelectedUnitPriceSheets: 新規シート作成成功、既存シートを入替"
    DeleteImportedUnitPriceSheetsExcept targetBook, stagedSheets

    Dim i As Long
    For i = 1 To stagedSheets.Count
        Set stagedSheet = stagedSheets(i)
        stagedSheet.Name = MakeUniqueWorksheetName(targetBook, CStr(targetSheetNames(i)), stagedSheet.Name)
        LogUP "シート確定 name=[" & stagedSheet.Name & "]"
    Next i

    ImportSelectedUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    Dim openedKey As Variant
    For Each openedKey In openedSourceBooks.Keys
        openedSourceBooks(openedKey).Close SaveChanges:=False
    Next openedKey
    Set sourceBook = Nothing
    If Not ImportSelectedUnitPriceSheets Then DeleteStagedWorksheets stagedSheets
    Application.DisplayAlerts = previousDisplayAlerts
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    CommonGetBasicInfoWorksheet(targetBook).Activate
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox UiMsgUnitPriceImportFailedText() & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportAndMergePurchaseUnitPriceSheets(ByVal sourceFilePath As String, _
                                                       ByVal selectedSheetNames As Collection, _
                                                       ByVal targetBook As Workbook, _
                                                       ByVal newSheetName As String) As Boolean
    Dim sourceBook As Workbook
    Dim newSheet As Worksheet
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)

    Dim isFirst As Boolean
    isFirst = True

    Dim sheetName As Variant
    For Each sheetName In selectedSheetNames
        Dim srcSheet As Worksheet
        Set srcSheet = sourceBook.worksheets(CStr(sheetName))

        If isFirst Then
            srcSheet.Copy After:=targetBook.worksheets(targetBook.worksheets.Count)
            Set newSheet = targetBook.worksheets(targetBook.worksheets.Count)
            newSheet.Name = MakeUniqueWorksheetName(targetBook, "__UP_PURCHASE_NEW", newSheet.Name)
            isFirst = False
        Else
            AppendSheetDataExcludingHeader srcSheet, newSheet
        End If
    Next sheetName

    If Not newSheet Is Nothing Then
        newSheet.Tab.Color = RGB(PURCHASE_SHEET_TAB_R, PURCHASE_SHEET_TAB_G, PURCHASE_SHEET_TAB_B)
        MarkImportedUnitPriceSheet newSheet
        ApplyImportedUnitPriceSheetFormat newSheet
        newSheet.Columns(IMPORTED_UNIT_PRICE_MERGE_FIRST_COL).ColumnWidth = PURCHASE_IMPORTED_UNIT_PRICE_COL_A_WIDTH
        DeleteWorksheetIfExists targetBook, newSheetName, newSheet
        newSheet.Name = MakeUniqueWorksheetName(targetBook, newSheetName, newSheet.Name)
    End If

    ImportAndMergePurchaseUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Not ImportAndMergePurchaseUnitPriceSheets Then
        If Not newSheet Is Nothing Then newSheet.Delete
    End If
    Application.DisplayAlerts = previousDisplayAlerts
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    Application.CutCopyMode = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox UiMsgPurchaseUnitPriceImportFailedText() & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportAndMergeWeldingUnitPriceSheets(ByVal sourceFilePath As String, _
                                                      ByVal selectedSheetNames As Collection, _
                                                      ByVal targetBook As Workbook, _
                                                      ByVal newSheetName As String) As Boolean
    Dim sourceBook As Workbook
    Dim newSheet As Worksheet
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)

    Dim isFirst As Boolean
    isFirst = True

    Dim sheetName As Variant
    For Each sheetName In selectedSheetNames
        Dim srcSheet As Worksheet
        Set srcSheet = sourceBook.worksheets(CStr(sheetName))

        If isFirst Then
            srcSheet.Copy After:=targetBook.worksheets(targetBook.worksheets.Count)
            Set newSheet = targetBook.worksheets(targetBook.worksheets.Count)
            newSheet.Name = MakeUniqueWorksheetName(targetBook, "__UP_WELDING_NEW", newSheet.Name)
            isFirst = False
        Else
            AppendSheetDataExcludingHeader srcSheet, newSheet
        End If
    Next sheetName

    If Not newSheet Is Nothing Then
        newSheet.Tab.Color = RGB(WELDING_SHEET_TAB_R, WELDING_SHEET_TAB_G, WELDING_SHEET_TAB_B)
        MarkImportedUnitPriceSheet newSheet
        FillBlankWeldingUnitPriceCells newSheet
        ApplyImportedUnitPriceSheetFormat newSheet
        ApplyWeldingUnitPriceSheetSectionFormat newSheet
        DeleteWorksheetIfExists targetBook, newSheetName, newSheet
        newSheet.Name = MakeUniqueWorksheetName(targetBook, newSheetName, newSheet.Name)
    End If

    ImportAndMergeWeldingUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If Not ImportAndMergeWeldingUnitPriceSheets Then
        If Not newSheet Is Nothing Then newSheet.Delete
    End If
    Application.DisplayAlerts = previousDisplayAlerts
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    Application.CutCopyMode = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox UiMsgWeldingUnitPriceImportFailedText() & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Sub AppendSheetDataExcludingHeader(ByVal srcSheet As Worksheet, ByVal destSheet As Worksheet)
    Dim srcUsed As Range
    Set srcUsed = srcSheet.UsedRange
    If srcUsed Is Nothing Then Exit Sub

    Dim srcLastRow As Long, srcLastCol As Long
    srcLastRow = srcUsed.Row + srcUsed.rows.Count - 1
    srcLastCol = srcUsed.Column + srcUsed.Columns.Count - 1
    If srcLastRow <= SOURCE_HEADER_ROW Then Exit Sub

    Dim destLastRow As Long
    destLastRow = FindLastUsedRow(destSheet)
    If destLastRow < 1 Then destLastRow = 1

    srcSheet.Range(srcSheet.Cells(SOURCE_HEADER_ROW + 1, 1), srcSheet.Cells(srcLastRow, srcLastCol)).Copy
    destSheet.Cells(destLastRow + 1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
End Sub

Private Function FindLastUsedRow(ByVal targetSheet As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = targetSheet.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If foundCell Is Nothing Then FindLastUsedRow = 0 Else FindLastUsedRow = foundCell.Row
End Function

Private Function ImportPurchaseUnitPriceSheetsByReference(ByRef request As UnitPriceRequest, _
                                                          ByRef masterRow As UnitPriceMasterRow, _
                                                          ByVal priceFolderPath As String, _
                                                          ByVal sectionFolderPath As String, _
                                                          ByVal targetBook As Workbook, _
                                                          ByRef createdSheetName As String) As Boolean
    createdSheetName = ""
    LogUP "ImportPurchaseUnitPriceSheetsByReference 開始"

    Dim purchaseFilePath As String
    purchaseFilePath = FindPurchaseUnitPriceWorkbook(priceFolderPath)
    LogUP "purchaseFilePath=[" & purchaseFilePath & "]"
    If purchaseFilePath = "" Then
        LogUP "購入充当ブック未検出 -> スキップ"
        Exit Function
    End If

    Dim referenceKey As String
    referenceKey = BuildPurchaseReferenceKey(sectionFolderPath)
    LogUP "referenceKey=[" & referenceKey & "]"
    If referenceKey = "" Then
        LogUP "referenceKey 取得失敗 -> 中断"
        MsgBox UiMsgPurchaseReferenceKeyNotFoundText() & vbCrLf & sectionFolderPath, vbExclamation
        Exit Function
    End If

    Dim purchaseSheetNames As Collection
    Set purchaseSheetNames = LoadPurchaseSheetNamesByReference(purchaseFilePath, referenceKey)
    If purchaseSheetNames Is Nothing Then
        LogUP "LoadPurchaseSheetNamesByReference -> Nothing 中断"
        Exit Function
    End If
    LogUP "購入充当シート候補数=" & CStr(purchaseSheetNames.Count)
    If purchaseSheetNames.Count = 0 Then
        LogUP "参照キー一致シートなし key=[" & referenceKey & "] -> 中断"
        MsgBox UiMsgPurchaseSheetByKeyNotFoundPrefixText() & referenceKey & UiMsgPurchaseSheetByKeyNotFoundSuffixText() & vbCrLf & purchaseFilePath, vbExclamation
        Exit Function
    End If

    Dim newSheetName As String
    newSheetName = BuildPurchaseSheetName(GetPathBaseName(sectionFolderPath))
    If newSheetName = "" Then newSheetName = BuildPurchaseSheetName(masterRow.UnitPriceSectionName)
    LogUP "購入充当 newSheetName=[" & newSheetName & "]"
    If newSheetName = "" Then
        LogUP "newSheetName 生成失敗 -> 中断"
        MsgBox UiMsgPurchaseImportSheetNameFailedText(), vbExclamation
        Exit Function
    End If

    If Not ImportAndMergePurchaseUnitPriceSheets(purchaseFilePath, purchaseSheetNames, targetBook, newSheetName) Then
        LogUP "ImportAndMergePurchaseUnitPriceSheets -> False"
        Exit Function
    End If

    createdSheetName = newSheetName
    LogUP "購入充当単価シート作成完了 createdSheet=[" & createdSheetName & "]"
    ImportPurchaseUnitPriceSheetsByReference = True
End Function

Private Function ImportWeldingUnitPriceSheetsIfRequired(ByVal wsInfo As Worksheet, _
                                                        ByRef masterRow As UnitPriceMasterRow, _
                                                        ByVal priceFolderPath As String, _
                                                        ByVal sectionFolderPath As String, _
                                                        ByVal selectedLineNames As Collection, _
                                                        ByVal targetBook As Workbook, _
                                                        ByRef createdSheetName As String) As Boolean
    createdSheetName = ""
    LogUP "ImportWeldingUnitPriceSheetsIfRequired 開始"
    If Not IsWeldingUnitPriceRequired(wsInfo) Then
        LogUP "C23 溶接不要 -> スキップ"
        ImportWeldingUnitPriceSheetsIfRequired = True
        Exit Function
    End If
    LogUP "C23 溶接必要"

    Dim weldingFilePath As String
    weldingFilePath = FindWeldingUnitPriceWorkbook(priceFolderPath)
    LogUP "weldingFilePath=[" & weldingFilePath & "]"
    If weldingFilePath = "" Then
        LogUP "溶接ブック未検出 -> 警告表示してスキップ"
        MsgBox BuildMissingWeldingUnitPriceMessage(sectionFolderPath, masterRow), vbExclamation
        ImportWeldingUnitPriceSheetsIfRequired = True
        Exit Function
    End If

    Dim weldingSheetNames As Collection
    Set weldingSheetNames = LoadWeldingSheetNames(weldingFilePath, selectedLineNames)
    If weldingSheetNames Is Nothing Then
        LogUP "LoadWeldingSheetNames -> Nothing 中断"
        Exit Function
    End If
    LogUP "溶接シート候補数=" & CStr(weldingSheetNames.Count)
    If weldingSheetNames.Count = 0 Then
        LogUP "溶接対象シートなし -> 中断"
        MsgBox UiMsgWeldingImportableSheetNotFoundText() & vbCrLf & weldingFilePath, vbExclamation
        Exit Function
    End If

    Dim newSheetName As String
    newSheetName = BuildWeldingSheetName(GetPathBaseName(sectionFolderPath))
    If newSheetName = "" Then newSheetName = BuildWeldingSheetName(masterRow.UnitPriceSectionName)
    LogUP "溶接 newSheetName=[" & newSheetName & "]"
    If newSheetName = "" Then
        LogUP "newSheetName 生成失敗 -> 中断"
        MsgBox UiMsgWeldingImportSheetNameFailedText(), vbExclamation
        Exit Function
    End If

    If Not ImportAndMergeWeldingUnitPriceSheets(weldingFilePath, weldingSheetNames, targetBook, newSheetName) Then
        LogUP "ImportAndMergeWeldingUnitPriceSheets -> False"
        Exit Function
    End If

    createdSheetName = newSheetName
    LogUP "溶接単価シート作成完了 createdSheet=[" & createdSheetName & "]"
    ImportWeldingUnitPriceSheetsIfRequired = True
End Function

Private Function BuildMissingWeldingUnitPriceMessage(ByVal sectionFolderPath As String, _
                                                     ByRef masterRow As UnitPriceMasterRow) As String
    Dim sectionName As String
    sectionName = TrimLeadingDigitsAndSeparators(GetPathBaseName(sectionFolderPath))
    If sectionName = "" Then sectionName = TrimLeadingDigitsAndSeparators(masterRow.UnitPriceSectionName)
    If sectionName = "" Then sectionName = masterRow.UnitPriceSectionName

    BuildMissingWeldingUnitPriceMessage = sectionName & UiMsgWeldingUnitPriceSettingMissingSuffixText()
End Function

Private Function IsWeldingUnitPriceRequired(ByVal wsInfo As Worksheet) As Boolean
    If wsInfo Is Nothing Then Exit Function

    Dim targetCell As Range
    Set targetCell = wsInfo.Range(BASIC_INFO_WELDING_FLAG_CELL).MergeArea.Cells(1, 1)

    Dim normalizedFlag As String
    normalizedFlag = NormalizeMatchText(CStr(targetCell.value))
    IsWeldingUnitPriceRequired = (StrComp(normalizedFlag, NormalizeMatchText(WELDING_REQUIRED_VALUE), vbTextCompare) = 0 Or _
                                  (InStr(1, normalizedFlag, NormalizeMatchText("溶接"), vbTextCompare) > 0 And _
                                   InStr(1, normalizedFlag, NormalizeMatchText("あり"), vbTextCompare) > 0))
End Function

Private Function LoadWeldingSheetNames(ByVal sourceFilePath As String, _
                                       ByVal selectedLineNames As Collection) As Collection
    Dim sourceSheetNames As Collection
    Set sourceSheetNames = LoadWorksheetNamesFromWorkbook(sourceFilePath)
    If sourceSheetNames Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection

    Dim sheetName As Variant
    If Not selectedLineNames Is Nothing Then
        For Each sheetName In sourceSheetNames
            If CollectionContainsText(selectedLineNames, CStr(sheetName)) Then result.Add CStr(sheetName)
        Next sheetName
    End If

    If result.Count = 0 Then
        For Each sheetName In sourceSheetNames
            result.Add CStr(sheetName)
        Next sheetName
    End If

    Set LoadWeldingSheetNames = result
End Function

Private Function LoadPurchaseSheetNamesByReference(ByVal sourceFilePath As String, ByVal referenceKey As String) As Collection
    Dim sheetNames As Collection
    Set sheetNames = LoadWorksheetNamesFromWorkbook(sourceFilePath)
    If sheetNames Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection

    AddPurchaseSheetNamesByReferenceKey sheetNames, referenceKey, result
    If result.Count = 0 Then
        AddPurchaseSheetNamesByReferenceKey sheetNames, BuildPurchaseFallbackReferenceKey(referenceKey), result
    End If

    If result.Count = 0 And sheetNames.Count = 1 Then
        result.Add CStr(sheetNames(1))
    End If

    Set LoadPurchaseSheetNamesByReference = result
End Function

Private Sub AddPurchaseSheetNamesByReferenceKey(ByVal sheetNames As Collection, _
                                                ByVal referenceKey As String, _
                                                ByVal result As Collection)
    If sheetNames Is Nothing Or result Is Nothing Then Exit Sub
    If referenceKey = "" Then Exit Sub

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If PurchaseSheetNameMatchesReferenceKey(CStr(sheetName), referenceKey) Then
            If Not CollectionContainsText(result, CStr(sheetName)) Then result.Add CStr(sheetName)
        End If
    Next sheetName
End Sub

Private Function BuildPurchaseFallbackReferenceKey(ByVal referenceKey As String) As String
    Dim result As String
    result = NormalizeMatchText(referenceKey)

    result = TrimPurchaseReferenceSuffix(result, SHINKANSEN_NAME & "保線技術センター")
    result = TrimPurchaseReferenceSuffix(result, SHINKANSEN_NAME & "保線区")
    result = TrimPurchaseReferenceSuffix(result, "保線技術センター")
    result = TrimPurchaseReferenceSuffix(result, "保線区")
    result = TrimPurchaseReferenceSuffix(result, "地域鉄道部")
    result = TrimPurchaseReferenceSuffix(result, "鉄道部")

    If result <> NormalizeMatchText(referenceKey) Then BuildPurchaseFallbackReferenceKey = result
End Function

Private Function TrimPurchaseReferenceSuffix(ByVal sourceText As String, ByVal suffixText As String) As String
    Dim suffix As String
    suffix = NormalizeMatchText(suffixText)
    If sourceText = "" Or suffix = "" Then
        TrimPurchaseReferenceSuffix = sourceText
        Exit Function
    End If

    If Len(sourceText) > Len(suffix) Then
        If Right$(sourceText, Len(suffix)) = suffix Then
            TrimPurchaseReferenceSuffix = Left$(sourceText, Len(sourceText) - Len(suffix))
            Exit Function
        End If
    End If

    TrimPurchaseReferenceSuffix = sourceText
End Function

Private Function PurchaseSheetNameMatchesReferenceKey(ByVal sheetName As String, ByVal referenceKey As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(sheetName)
    b = NormalizeMatchText(referenceKey)
    If a = "" Or b = "" Then Exit Function
    PurchaseSheetNameMatchesReferenceKey = (a = b Or _
                                            Left$(a, Len(b) + 1) = b & "-" Or _
                                            Left$(a, Len(b) + 1) = b & "_" Or _
                                            InStr(1, a, b, vbTextCompare) > 0)
End Function

Private Function BuildPurchaseReferenceKey(ByVal sectionFolderPath As String) As String
    Dim baseName As String
    baseName = GetPathBaseName(sectionFolderPath)

    Dim digits As String
    digits = ExtractLeadingDigits(baseName)
    If digits <> "" Then
        BuildPurchaseReferenceKey = digits
        Exit Function
    End If

    BuildPurchaseReferenceKey = NormalizeMatchText(baseName)
End Function

Private Function ExtractLeadingDigits(ByVal value As String) As String
    Dim result As String, i As Long, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[0-9]" Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            Exit For
        End If
    Next i
    ExtractLeadingDigits = result
End Function

Private Function GetPathBaseName(ByVal sourcePath As String) As String
    Dim pos As Long
    pos = InStrRev(sourcePath, "\")
    If pos > 0 Then GetPathBaseName = Mid$(sourcePath, pos + 1) Else GetPathBaseName = sourcePath
End Function

Private Function BuildPurchaseSheetName(ByVal rawSectionName As String) As String
    Dim trimmed As String
    trimmed = TrimLeadingDigitsAndSeparators(rawSectionName)
    If trimmed = "" Then Exit Function
    BuildPurchaseSheetName = Left$(MakeSafeWorksheetName(trimmed & PURCHASE_SHEET_NAME_SUFFIX), 31)
End Function

Private Function BuildWeldingSheetName(ByVal rawSectionName As String) As String
    Dim trimmed As String
    trimmed = TrimLeadingDigitsAndSeparators(rawSectionName)
    If trimmed = "" Then Exit Function
    BuildWeldingSheetName = Left$(MakeSafeWorksheetName(trimmed & WELDING_SHEET_NAME_SUFFIX), 31)
End Function

Private Function TrimLeadingDigitsAndSeparators(ByVal rawValue As String) As String
    Dim result As String
    result = Trim$(CommonNormalizeText(rawValue))
    Do While Len(result) > 0
        If Mid$(result, 1, 1) Like "[0-9]" Then result = Mid$(result, 2) Else Exit Do
    Loop
    Do While Len(result) > 0 And (Left$(result, 1) = "_" Or Left$(result, 1) = "-" Or Left$(result, 1) = " ")
        result = Mid$(result, 2)
    Loop
    TrimLeadingDigitsAndSeparators = Trim$(result)
End Function

' ============================================================
' 単価シート名候補(displayName)を、工事件名別マスタ D列(積算線区コード)の
' 番号順に並べ替える。照合は実シート名(sheetSourceSheetMap)で行う。
' マスタD列に該当しない(rank<0)候補は除外する。
' 同順位内は元の出現順を保持する(安定ソート)。
Private Function OrderUnitPriceSheetNamesByProjectMasterFColumn( _
        ByVal sheetNames As Collection, _
        ByVal sheetSourceSheetMap As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    If sheetNames Is Nothing Then
        Set OrderUnitPriceSheetNamesByProjectMasterFColumn = result
        Exit Function
    End If

    ' (rank, 元順, displayName) を集め、rank 昇順→元順昇順で並べ替える。
    Dim ranks() As Long
    Dim origins() As Long
    Dim names() As String
    Dim n As Long
    n = 0
    ReDim ranks(1 To sheetNames.Count)
    ReDim origins(1 To sheetNames.Count)
    ReDim names(1 To sheetNames.Count)

    Dim originIndex As Long
    originIndex = 0

    Dim displayName As Variant
    For Each displayName In sheetNames
        originIndex = originIndex + 1

        Dim lookupName As String
        lookupName = CStr(displayName)
        If Not sheetSourceSheetMap Is Nothing Then
            If sheetSourceSheetMap.Exists(CStr(displayName)) Then
                lookupName = CStr(sheetSourceSheetMap(CStr(displayName)))
            End If
        End If

        Dim rank As Long
        rank = mod_Construction_Order_Import.GetProjectMasterLineOrderRank(lookupName)
        If rank >= 0 Then
            n = n + 1
            ranks(n) = rank
            origins(n) = originIndex
            names(n) = CStr(displayName)
        Else
            LogUP "D列コード順除外(マスタD列に未登録) sheet=[" & CStr(displayName) & "] lookup=[" & lookupName & "]"
        End If
    Next displayName

    If n = 0 Then
        Set OrderUnitPriceSheetNamesByProjectMasterFColumn = result
        Exit Function
    End If

    ' 挿入ソート(安定): キー=(rank, 元順)
    Dim i As Long, j As Long
    Dim kr As Long, ko As Long
    Dim kn As String
    For i = 2 To n
        kr = ranks(i)
        ko = origins(i)
        kn = names(i)
        j = i - 1
        Do While j >= 1
            If (ranks(j) > kr) Or (ranks(j) = kr And origins(j) > ko) Then
                ranks(j + 1) = ranks(j)
                origins(j + 1) = origins(j)
                names(j + 1) = names(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        ranks(j + 1) = kr
        origins(j + 1) = ko
        names(j + 1) = kn
    Next i

    For i = 1 To n
        result.Add names(i)
    Next i

    Set OrderUnitPriceSheetNamesByProjectMasterFColumn = result
End Function

Private Sub WriteSelectedLineNames(ByVal wsInfo As Worksheet, ByVal selectedSheetNames As Collection)
    If wsInfo Is Nothing Then Exit Sub
    Dim joinedText As String
    joinedText = JoinCollectionText(selectedSheetNames, ChrW$(&H3001))
    LogUP "WriteSelectedLineNames C24=[" & joinedText & "]"
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).value = joinedText
    FormatImportedLineNamesCell wsInfo
End Sub

Public Sub FormatImportedLineNamesCell(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetCell As Range
    Set targetCell = wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.Cells(1, 1)

    Dim formattedText As String
    formattedText = NormalizeImportedLineNameText(CStr(targetCell.value))
    If CStr(targetCell.value) <> formattedText Then targetCell.value = formattedText

    targetCell.MergeArea.WrapText = True
    targetCell.MergeArea.VerticalAlignment = xlVAlignCenter
    targetCell.MergeArea.Font.Size = CalculateImportedLineNameFontSize(targetCell.MergeArea, formattedText)

    ' 線区名(C24:C28)はEnableEvents=Falseの取込経路でも書込まれるため、
    ' ガイド(塗色・コメント)を明示的に更新して黄色塗りを解除/復帰する。
    On Error Resume Next
    mod_BasicInfoGuide.OnCellChanged wsInfo, wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL)
    On Error GoTo 0
End Sub

Private Function CalculateImportedLineNameFontSize(ByVal targetArea As Range, ByVal LineNameText As String) As Double
    If targetArea Is Nothing Then
        CalculateImportedLineNameFontSize = IMPORTED_LINE_NAME_BASE_FONT_SIZE
        Exit Function
    End If

    Dim lineCount As Long
    lineCount = CountImportedLineNameLines(LineNameText)
    If lineCount < 1 Then lineCount = 1

    Dim availableHeight As Double
    availableHeight = targetArea.Height - IMPORTED_LINE_NAME_VERTICAL_PADDING
    If availableHeight <= 0 Then availableHeight = targetArea.Height

    Dim fontSize As Double
    fontSize = Int(availableHeight / (lineCount * IMPORTED_LINE_NAME_LINE_HEIGHT_RATIO))

    If fontSize > IMPORTED_LINE_NAME_BASE_FONT_SIZE Then fontSize = IMPORTED_LINE_NAME_BASE_FONT_SIZE
    If fontSize < IMPORTED_LINE_NAME_MIN_FONT_SIZE Then fontSize = IMPORTED_LINE_NAME_MIN_FONT_SIZE

    CalculateImportedLineNameFontSize = fontSize
End Function

Private Function NormalizeImportedLineNameText(ByVal sourceText As String) As String
    Dim normalized As String
    normalized = Replace$(Replace$(sourceText, vbCrLf, vbLf), vbCr, vbLf)
    normalized = Replace$(normalized, ChrW$(&H3001), vbLf)

    Dim parts As Variant
    parts = Split(normalized, vbLf)

    Dim result As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim lineText As String
        lineText = Trim$(CStr(parts(i)))
        If lineText <> "" Then
            If result <> "" Then result = result & vbLf
            result = result & lineText
        End If
    Next i
    NormalizeImportedLineNameText = result
End Function

Private Function CountImportedLineNameLines(ByVal LineNameText As String) As Long
    If LineNameText = "" Then
        CountImportedLineNameLines = 1
        Exit Function
    End If

    Dim i As Long
    CountImportedLineNameLines = 1
    For i = 1 To Len(LineNameText)
        If Mid$(LineNameText, i, 1) = vbLf Then CountImportedLineNameLines = CountImportedLineNameLines + 1
    Next i
End Function

Private Function JoinCollectionText(ByVal values As Collection, ByVal delimiter As String) As String
    If values Is Nothing Then Exit Function
    Dim result As String, item As Variant
    For Each item In values
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next item
    JoinCollectionText = result
End Function

Public Sub ConfirmAndClearUnitPriceForBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If MsgBox(UiMsgUnitPriceClearConfirmPromptText() & vbCrLf & _
              UiMsgUnitPriceClearConfirmYesLineText() & vbCrLf & _
              UiMsgUnitPriceClearConfirmNoLineText(), vbQuestion + vbYesNo, UiMsgUnitPriceClearConfirmTitleText()) <> vbYes Then Exit Sub
    ClearUnitPriceSheets wsInfo.Parent
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.ClearContents
    FormatImportedLineNamesCell wsInfo
End Sub

Public Sub SilentClearUnitPriceForBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    ClearUnitPriceSheets wsInfo.Parent
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.ClearContents
    FormatImportedLineNamesCell wsInfo
End Sub

'  HandleImportedLineNamesCellChange
Public Function IsClearingImportedLineNames() As Boolean
    IsClearingImportedLineNames = mClearingImportedLineNames
End Function

Public Sub HandleImportedLineNamesCellChange(ByVal wsInfo As Worksheet, ByVal changedRange As Range)
    If wsInfo Is Nothing Or changedRange Is Nothing Then Exit Sub
    If mClearingImportedLineNames Then Exit Sub

    If Not IsImportedLineNamesMonitorRangeChanged(wsInfo, changedRange) Then Exit Sub
    If Not IsImportedLineNamesCellEmpty(wsInfo) Then Exit Sub

    LogUP "C24 cleared -> delete imported unit price sheets changed=[" & changedRange.Address(False, False) & "]"
    mClearingImportedLineNames = True
    On Error GoTo FinallyExit
    ClearUnitPriceSheets wsInfo.Parent
    FormatImportedLineNamesCell wsInfo
FinallyExit:
    mClearingImportedLineNames = False
End Sub

Public Function IsConstructionUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function
    If IsProtectedSystemWorksheet(targetSheet) Then Exit Function
    If IsSpecialImportedUnitPriceSheet(targetSheet) Then Exit Function

    If IsImportedUnitPriceSheetByMarker(targetSheet) Then
        IsConstructionUnitPriceSheet = True
        Exit Function
    End If

    On Error Resume Next
    If targetSheet.Tab.ColorIndex <> xlColorIndexNone Then
        IsConstructionUnitPriceSheet = _
            (targetSheet.Tab.Color = RGB(UNIT_PRICE_SHEET_TAB_R, UNIT_PRICE_SHEET_TAB_G, UNIT_PRICE_SHEET_TAB_B))
    End If
    On Error GoTo 0
End Function

Private Function IsSpecialImportedUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = CommonNormalizeText(CStr(targetSheet.Name))
    IsSpecialImportedUnitPriceSheet = _
        (InStr(1, sheetName, NormalizeMatchText(PURCHASE_SHEET_NAME_SUFFIX), vbTextCompare) > 0) Or _
        (InStr(1, sheetName, NormalizeMatchText(WELDING_SHEET_NAME_SUFFIX), vbTextCompare) > 0)
End Function

Public Sub ClearUnitPriceSheets(Optional ByVal targetBook As Workbook)
    Dim savedErrNumber As Long, savedErrSource As String, savedErrDescription As String
    If targetBook Is Nothing Then Set targetBook = ThisWorkbook
    Dim previousDisplayAlerts As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    previousDisplayAlerts = Application.DisplayAlerts
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.screenUpdating
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.screenUpdating = False
    DeleteImportedUnitPriceSheets targetBook
    ClearCurrentUnitPriceImportBatchId targetBook
Cleanup:
    Application.screenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts
    If savedErrNumber <> 0 Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Exit Sub
ErrorHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.source
    savedErrDescription = Err.Description
    Resume Cleanup
End Sub

Private Sub DeleteImportedUnitPriceSheets(ByVal targetBook As Workbook)
    DeleteImportedUnitPriceSheetsExcept targetBook, Nothing
End Sub

Private Sub DeleteImportedUnitPriceSheetsExcept(ByVal targetBook As Workbook, _
                                                ByVal keepSheets As Collection)
    Dim i As Long
    Dim deletedCount As Long
    For i = targetBook.worksheets.Count To 1 Step -1
        If IsImportedUnitPriceSheet(targetBook.worksheets(i)) Then
            If Not WorksheetIsInCollection(targetBook.worksheets(i), keepSheets) And _
               targetBook.worksheets.Count > 1 Then
                LogUP "DeleteImportedUnitPriceSheets: delete [" & targetBook.worksheets(i).Name & "]"
                targetBook.worksheets(i).Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next i
    LogUP "DeleteImportedUnitPriceSheets: deletedCount=" & CStr(deletedCount)
End Sub

Private Sub DeleteStagedWorksheets(ByVal stagedSheets As Collection)
    If stagedSheets Is Nothing Then Exit Sub

    Dim i As Long
    On Error Resume Next
    For i = stagedSheets.Count To 1 Step -1
        stagedSheets(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function WorksheetIsInCollection(ByVal targetSheet As Worksheet, _
                                         ByVal worksheets As Collection) As Boolean
    If targetSheet Is Nothing Or worksheets Is Nothing Then Exit Function

    Dim item As Worksheet
    For Each item In worksheets
        If item Is targetSheet Then
            WorksheetIsInCollection = True
            Exit Function
        End If
    Next item
End Function

Private Sub DeleteWorksheetIfExists(ByVal targetBook As Workbook, _
                                    ByVal sheetName As String, _
                                    Optional ByVal keepSheet As Worksheet)
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetBook.worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If Not keepSheet Is Nothing Then
        If targetSheet Is keepSheet Then Exit Sub
    End If
    If targetBook.worksheets.Count <= 1 Then Exit Sub
    If Not IsImportedUnitPriceSheet(targetSheet) Then Exit Sub
    targetSheet.Delete
End Sub

Private Function IsImportedUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function
    If IsProtectedSystemWorksheet(targetSheet) Then Exit Function
    If IsImportedUnitPriceSheetByMarker(targetSheet) Then
        IsImportedUnitPriceSheet = True
        Exit Function
    End If
    IsImportedUnitPriceSheet = IsImportedUnitPriceSheetByNameSuffix(targetSheet)
End Function

Private Function IsProtectedSystemWorksheet(ByVal targetSheet As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = CommonNormalizeText(CStr(targetSheet.Name))
    IsProtectedSystemWorksheet = _
        (StrComp(sheetName, CommonNormalizeText(CommonBasicInfoSheetNameText()), vbTextCompare) = 0) Or _
        (StrComp(sheetName, CommonNormalizeText(CommonCoverSheetNameText()), vbTextCompare) = 0) Or _
        (StrComp(sheetName, "Cover", vbTextCompare) = 0) Or _
        (StrComp(sheetName, "DebugLog", vbTextCompare) = 0)
End Function

Private Function IsImportedUnitPriceSheetByNameSuffix(ByVal targetSheet As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = CommonNormalizeText(CStr(targetSheet.Name))
    IsImportedUnitPriceSheetByNameSuffix = _
        TextEndsWith(sheetName, NormalizeMatchText(PURCHASE_SHEET_NAME_SUFFIX)) Or _
        TextEndsWith(sheetName, NormalizeMatchText(WELDING_SHEET_NAME_SUFFIX))
End Function

Private Function TextEndsWith(ByVal sourceText As String, ByVal suffixText As String) As Boolean
    If suffixText = "" Or Len(sourceText) < Len(suffixText) Then Exit Function
    TextEndsWith = (StrComp(Right$(sourceText, Len(suffixText)), suffixText, vbTextCompare) = 0)
End Function

Private Function IsImportedUnitPriceSheetByMarker(ByVal targetSheet As Worksheet) As Boolean
    On Error Resume Next
    IsImportedUnitPriceSheetByMarker = _
        (CStr(targetSheet.Range(IMPORTED_SHEET_MARKER_ADDRESS).value) = IMPORTED_SHEET_MARKER_VALUE) Or _
        (CStr(targetSheet.Range(IMPORTED_SHEET_LEGACY_MARKER_ADDRESS).value) = IMPORTED_SHEET_MARKER_VALUE)
    On Error GoTo 0
End Function

Private Function IsImportedUnitPriceSheetByTabColor(ByVal targetSheet As Worksheet) As Boolean
    On Error Resume Next
    If targetSheet.Tab.ColorIndex = xlColorIndexNone Then Exit Function
    On Error GoTo 0

    Dim tabColor As Long
    tabColor = targetSheet.Tab.Color

    IsImportedUnitPriceSheetByTabColor = _
        (tabColor = RGB(UNIT_PRICE_SHEET_TAB_R, UNIT_PRICE_SHEET_TAB_G, UNIT_PRICE_SHEET_TAB_B)) Or _
        (tabColor = RGB(PURCHASE_SHEET_TAB_R, PURCHASE_SHEET_TAB_G, PURCHASE_SHEET_TAB_B)) Or _
        (tabColor = RGB(WELDING_SHEET_TAB_R, WELDING_SHEET_TAB_G, WELDING_SHEET_TAB_B))
End Function

Private Sub MarkImportedUnitPriceSheet(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    targetSheet.Range(IMPORTED_SHEET_MARKER_ADDRESS).value = IMPORTED_SHEET_MARKER_VALUE
    targetSheet.Range(IMPORTED_SHEET_BATCH_ADDRESS).value = mCurrentImportBatchId
    On Error GoTo 0
End Sub

' 取込1回分を識別するバッチIDを採番する(一意性より「前回と区別できること」を重視)。
Private Function GenerateImportBatchId() As String
    GenerateImportBatchId = Format(Now, "yyyymmddhhnnss") & "_" & _
                            Format(CLng(Timer * 1000) Mod 100000, "00000")
End Function

' 現行バッチIDをブックに永続化する(セルを使わずDefined Nameに保持し、
' UsedRange/書式の肥大化に影響しないようにする)。
Public Sub SetCurrentUnitPriceImportBatchId(ByVal wb As Workbook, ByVal batchId As String)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Names(CURRENT_BATCH_DEFINED_NAME).Delete
    On Error GoTo 0
    On Error Resume Next
    wb.Names.Add Name:=CURRENT_BATCH_DEFINED_NAME, RefersTo:="=""" & batchId & """"
    On Error GoTo 0
End Sub

' 現行バッチIDを取得する。未設定(旧仕様のブック・一度も取込していない等)の場合は空文字を返す。
Public Function GetCurrentUnitPriceImportBatchId(Optional ByVal wb As Workbook) As String
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim raw As String
    On Error Resume Next
    raw = wb.Names(CURRENT_BATCH_DEFINED_NAME).RefersTo
    On Error GoTo 0
    If Len(raw) = 0 Then Exit Function

    ' raw は ="20260701153045_00123" の形式。前後の ="  " を除去する。
    If Left$(raw, 2) = "=""" Then raw = Mid$(raw, 3)
    If Len(raw) > 0 And Right$(raw, 1) = """" Then raw = Left$(raw, Len(raw) - 1)
    GetCurrentUnitPriceImportBatchId = raw
End Function

' 単価表クリア時は現行バッチも失効させる(参照先が無くなるため)。
Public Sub ClearCurrentUnitPriceImportBatchId(Optional ByVal wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    wb.Names(CURRENT_BATCH_DEFINED_NAME).Delete
    On Error GoTo 0
End Sub

' 指定シートが「現行バッチ」の単価表/購入充当単価/レール溶接単価シートかどうかを判定する。
' 現行バッチIDが未設定の場合(旧仕様のブック等)は安全側に倒して True を返す
' (=以前と同じ「全シート対象」の挙動にフォールバックする)。
Public Function IsCurrentImportBatchUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    If targetSheet Is Nothing Then Exit Function

    Dim currentBatchId As String
    currentBatchId = GetCurrentUnitPriceImportBatchId(targetSheet.Parent)
    If currentBatchId = "" Then
        IsCurrentImportBatchUnitPriceSheet = True
        Exit Function
    End If

    On Error Resume Next
    IsCurrentImportBatchUnitPriceSheet = _
        (CStr(targetSheet.Range(IMPORTED_SHEET_BATCH_ADDRESS).value) = currentBatchId)
    On Error GoTo 0
End Function

Private Sub ApplyImportedUnitPriceSheetFormat(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then Exit Sub

    ' 書式設定は装飾処理。失敗しても単価表の取り込み(データ)自体は止めないよう、
    ' 各手順を個別にエラートラップし、エラーは呼び出し元へ伝播させない。
    Dim wasProtected As Boolean
    On Error Resume Next
    wasProtected = targetSheet.ProtectContents
    If wasProtected Then targetSheet.Unprotect
    On Error GoTo 0

    Dim fmtErr As Long
    fmtErr = 0

    Dim fontRange As Range
    Set fontRange = targetSheet.UsedRange
    If fontRange Is Nothing Then Set fontRange = targetSheet.Range("A1")

    On Error Resume Next
    fontRange.Font.Name = ImportedUnitPriceSheetFontNameText()
    If Err.Number <> 0 Then
        fmtErr = Err.Number
        Err.Clear
    End If
    On Error GoTo 0

    On Error Resume Next
    targetSheet.Columns(IMPORTED_UNIT_PRICE_CENTER_COL).HorizontalAlignment = xlCenter
    If Err.Number <> 0 Then
        If fmtErr = 0 Then fmtErr = Err.Number
        Err.Clear
    End If

    targetSheet.Range( _
        targetSheet.Cells(IMPORTED_UNIT_PRICE_MERGE_FIRST_ROW, IMPORTED_UNIT_PRICE_MERGE_FIRST_COL), _
        targetSheet.Cells(IMPORTED_UNIT_PRICE_MERGE_LAST_ROW, IMPORTED_UNIT_PRICE_MERGE_LAST_COL)).UnMerge

    Dim mergeRowIndex As Long
    For mergeRowIndex = IMPORTED_UNIT_PRICE_MERGE_FIRST_ROW To IMPORTED_UNIT_PRICE_MERGE_LAST_ROW
        With targetSheet.Range( _
                targetSheet.Cells(mergeRowIndex, IMPORTED_UNIT_PRICE_MERGE_FIRST_COL), _
                targetSheet.Cells(mergeRowIndex, IMPORTED_UNIT_PRICE_MERGE_LAST_COL))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        If Err.Number <> 0 Then
            If fmtErr = 0 Then fmtErr = Err.Number
            Err.Clear
        End If
    Next mergeRowIndex

    targetSheet.Columns(IMPORTED_UNIT_PRICE_MERGE_FIRST_COL).ColumnWidth = IMPORTED_UNIT_PRICE_COL_A_WIDTH
    If Err.Number <> 0 Then
        If fmtErr = 0 Then fmtErr = Err.Number
        Err.Clear
    End If

    targetSheet.Range( _
        targetSheet.Cells(IMPORTED_UNIT_PRICE_LEFT_ALIGN_ROW_FIRST, IMPORTED_UNIT_PRICE_CENTER_COL), _
        targetSheet.Cells(IMPORTED_UNIT_PRICE_LEFT_ALIGN_ROW_LAST, IMPORTED_UNIT_PRICE_CENTER_COL)).HorizontalAlignment = xlLeft
    If Err.Number <> 0 Then
        If fmtErr = 0 Then fmtErr = Err.Number
        Err.Clear
    End If
    On Error GoTo 0

    If fmtErr <> 0 Then
        LogUP "ApplyImportedUnitPriceSheetFormat: 一部書式の適用に失敗 err=" & CStr(fmtErr) & _
              " sheet=[" & targetSheet.Name & "]"
    End If

    If wasProtected Then
        On Error Resume Next
        targetSheet.Protect
        On Error GoTo 0
    End If
End Sub

' レール溶接単価シート: A列の「工事件名：」「積算線区：」行の右隣(B列)を左詰めにする
Private Sub ApplyWeldingUnitPriceSheetSectionFormat(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = FindLastUsedRow(targetSheet)
    If lastRow < 1 Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        Dim labelText As String
        labelText = CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(targetSheet.Cells(rowIndex, 1).value)))
        If Len(labelText) = 0 Then GoTo NextRow

        If InStr(1, labelText, WeldingProjectNameLabelKeywordText(), vbTextCompare) > 0 Or _
           InStr(1, labelText, WeldingLineSectionLabelKeywordText(), vbTextCompare) > 0 Then
            targetSheet.Cells(rowIndex, IMPORTED_UNIT_PRICE_CENTER_COL).HorizontalAlignment = xlLeft
        End If
NextRow:
    Next rowIndex
End Sub

Private Function WeldingProjectNameLabelKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H4EF6) & ChrW$(&H540D)
    WeldingProjectNameLabelKeywordText = cached
End Function

Private Function WeldingLineSectionLabelKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H7A4D) & ChrW$(&H7B97) & ChrW$(&H7DDA) & ChrW$(&H533A)
    WeldingLineSectionLabelKeywordText = cached
End Function

Private Function IsUnitPriceSheetDataSeiriRow(ByVal cellValue As Variant) As Boolean
    Dim textValue As String
    textValue = Trim$(StrConv(CommonNzText(cellValue), vbNarrow))
    If Len(textValue) = 0 Then Exit Function
    IsUnitPriceSheetDataSeiriRow = IsNumeric(textValue)
End Function

Private Function ImportedUnitPriceSheetFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    ImportedUnitPriceSheetFontNameText = cached
End Function

Private Function IsImportedLineNamesMonitorRangeChanged(ByVal wsInfo As Worksheet, ByVal changedRange As Range) As Boolean
    Dim monitorRange As Range
    Dim importedLineRange As Range

    Set monitorRange = wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_MONITOR_RANGE)
    Set importedLineRange = wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea

    IsImportedLineNamesMonitorRangeChanged = _
        (Not Intersect(changedRange, monitorRange) Is Nothing) Or _
        (Not Intersect(changedRange, importedLineRange) Is Nothing)
End Function

Private Function IsImportedLineNamesCellEmpty(ByVal wsInfo As Worksheet) As Boolean
    IsImportedLineNamesCellEmpty = _
        (Len(Trim$(CStr(wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.Cells(1, 1).value))) = 0)
End Function

Private Sub FillBlankUnitPriceEFCells(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then Exit Sub
    If IsPurchaseUnitPriceProjectName(CStr(targetSheet.Name)) Then Exit Sub

    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.rows.Count, UNIT_PRICE_BLANK_FILL_LAST_ROW_COL).End(xlUp).Row
    If lastRow < UNIT_PRICE_BLANK_FILL_DATA_START_ROW Then Exit Sub

    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = UNIT_PRICE_BLANK_FILL_DATA_START_ROW To lastRow
        If Not IsUnitPriceSheetDataSeiriRow(targetSheet.Cells(rowIndex, UNIT_PRICE_BLANK_FILL_LAST_ROW_COL).value) Then GoTo NextFillRow
        For colIndex = UNIT_PRICE_BLANK_FILL_COL_START To UNIT_PRICE_BLANK_FILL_COL_END
            With targetSheet.Cells(rowIndex, colIndex)
                If Len(Trim$(CStr(.value))) = 0 Then
                    .Interior.Color = RGB(UNIT_PRICE_BLANK_FILL_COLOR_R, _
                                          UNIT_PRICE_BLANK_FILL_COLOR_G, _
                                          UNIT_PRICE_BLANK_FILL_COLOR_B)
                End If
            End With
        Next colIndex
NextFillRow:
    Next rowIndex
End Sub

Private Sub FillBlankWeldingUnitPriceCells(ByVal targetSheet As Worksheet)
    ' レール溶接単価シート専用の塗りつぶし処理。
    ' A列に「XXXX年度」形式のヘッダー行がある場合、その行と直下WELDING_YEAR_HEADER_SKIP_ROWS行は塗りつぶし対象外。
    ' 対象列: E/F列(準日/夜単価) および I列以降(軌道会社単価)。
    ' 塗りつぶし条件: 年度ヘッダー行以外のデータ行(B列数値)で、対象列の値が空のセル。
    If targetSheet Is Nothing Then Exit Sub
    If IsPurchaseUnitPriceProjectName(CStr(targetSheet.Name)) Then Exit Sub

    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, UNIT_PRICE_BLANK_FILL_LAST_ROW_COL).End(xlUp).Row
    If lastRow < UNIT_PRICE_BLANK_FILL_DATA_START_ROW Then Exit Sub

    Dim lastCol As Long
    lastCol = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column

    Dim skipUntilRow As Long
    skipUntilRow = 0

    Dim rowIndex As Long
    For rowIndex = UNIT_PRICE_BLANK_FILL_DATA_START_ROW To lastRow
        ' A列に年度ヘッダー（XXXX年度）があればその行と直下 WELDING_YEAR_HEADER_SKIP_ROWS 行をスキップ
        Dim aVal As String
        aVal = CommonRemoveAllSpaces(CommonNzText(targetSheet.Cells(rowIndex, 1).Value))
        If IsWeldingYearHeaderRow(aVal) Then
            skipUntilRow = rowIndex + WELDING_YEAR_HEADER_SKIP_ROWS
        End If
        If rowIndex <= skipUntilRow Then GoTo NextWeldFillRow

        ' B列が整数（データ行）でなければスキップ
        If Not IsUnitPriceSheetDataSeiriRow(targetSheet.Cells(rowIndex, UNIT_PRICE_BLANK_FILL_LAST_ROW_COL).Value) Then GoTo NextWeldFillRow

        ' E/F列：値が空の場合に塗りつぶし
        Dim colIndex As Long
        For colIndex = UNIT_PRICE_BLANK_FILL_COL_START To UNIT_PRICE_BLANK_FILL_COL_END
            With targetSheet.Cells(rowIndex, colIndex)
                If Len(Trim$(CStr(.Value))) = 0 Then
                    .Interior.Color = RGB(UNIT_PRICE_BLANK_FILL_COLOR_R, _
                                          UNIT_PRICE_BLANK_FILL_COLOR_G, _
                                          UNIT_PRICE_BLANK_FILL_COLOR_B)
                End If
            End With
        Next colIndex

        ' I列以降：計算式が入っているが値として表示されない（空文字列）セルのみ塗りつぶす
        If lastCol >= WELDING_BLANK_FILL_COL_I_START Then
            For colIndex = WELDING_BLANK_FILL_COL_I_START To lastCol
                With targetSheet.Cells(rowIndex, colIndex)
                    If .HasFormula Then
                        If Len(Trim$(CStr(.Value))) = 0 Then
                            .Interior.Color = RGB(UNIT_PRICE_BLANK_FILL_COLOR_R, _
                                                  UNIT_PRICE_BLANK_FILL_COLOR_G, _
                                                  UNIT_PRICE_BLANK_FILL_COLOR_B)
                        End If
                    End If
                End With
            Next colIndex
        End If

NextWeldFillRow:
    Next rowIndex
End Sub

Private Function IsWeldingYearHeaderRow(ByVal normalizedAColText As String) As Boolean
    ' A列の内容が「XXXX年度」パターンかどうかを判定する。
    ' 例: 2026年度、 2027年度 など。
    If Len(normalizedAColText) = 0 Then Exit Function
    Dim nenDoText As String
    nenDoText = ChrW$(&H5E74) & ChrW$(&H5EA6)   ' 年度
    If InStr(1, normalizedAColText, nenDoText, vbTextCompare) = 0 Then Exit Function
    ' 数字が4桁以上含まれるか確認（XXXX年度の形式）
    Dim i As Long
    Dim digitCount As Long
    For i = 1 To Len(normalizedAColText)
        If Mid$(normalizedAColText, i, 1) >= "0" And Mid$(normalizedAColText, i, 1) <= "9" Then
            digitCount = digitCount + 1
        End If
    Next i
    IsWeldingYearHeaderRow = (digitCount >= 4)
End Function

Private Function MakeUniqueWorksheetName(ByVal targetBook As Workbook, _
                                         ByVal requestedName As String, _
                                         ByVal currentName As String) As String
    Dim baseName As String
    baseName = Left$(MakeSafeWorksheetName(requestedName), 31)
    If baseName = "" Then baseName = "単価表"
    If StrComp(baseName, currentName, vbTextCompare) = 0 Then
        MakeUniqueWorksheetName = currentName
        Exit Function
    End If
    If Not WorksheetExists(targetBook, baseName) Then
        MakeUniqueWorksheetName = baseName
        Exit Function
    End If
    Dim i As Long, candidate As String
    For i = 2 To 99
        candidate = Left$(baseName, 28) & "(" & CStr(i) & ")"
        If Not WorksheetExists(targetBook, candidate) Then
            MakeUniqueWorksheetName = candidate
            Exit Function
        End If
    Next i
    MakeUniqueWorksheetName = currentName
End Function

Private Function MakeSafeWorksheetName(ByVal sourceName As String) As String
    Dim result As String
    result = Trim$(sourceName)
    Dim i As Long
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, CStr(invalidChars(i)), "_")
    Next i
    MakeSafeWorksheetName = result
End Function

Private Function WorksheetExists(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetBook.worksheets(sheetName)
    WorksheetExists = Not targetSheet Is Nothing
    On Error GoTo 0
End Function

Private Function LoadUnitPriceProjectNamesForBasicInfo(ByVal sourceFilePath As String, _
                                                       ByVal lineType As String) As Collection
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_NAME
            Set LoadUnitPriceProjectNamesForBasicInfo = LoadUnitPriceProjectNames(sourceFilePath, ZAIRAISEN_NAME)
        Case SHINKANSEN_NAME
            Set LoadUnitPriceProjectNamesForBasicInfo = LoadUnitPriceProjectNames(sourceFilePath, SHINKANSEN_NAME)
        Case Else
            Dim result As Collection
            Set result = New Collection
            AddUnitPriceProjectNames result, LoadUnitPriceProjectNames(sourceFilePath, ZAIRAISEN_NAME)
            AddUnitPriceProjectNames result, LoadUnitPriceProjectNames(sourceFilePath, SHINKANSEN_NAME)
            Set LoadUnitPriceProjectNamesForBasicInfo = result
    End Select
End Function

Private Sub AddUnitPriceProjectNames(ByVal target As Collection, ByVal source As Collection)
    If target Is Nothing Or source Is Nothing Then Exit Sub
    Dim item As Variant
    For Each item In source
        If Not CollectionContainsText(target, CStr(item)) Then target.Add CStr(item)
    Next item
End Sub

Private Function LoadUnitPriceProjectNames(ByVal sourceFilePath As String, _
                                           ByVal lineType As String) As Collection
    Dim sheetName As String, endRow As Long
    On Error GoTo ErrorHandler
    ResolveUnitPriceProjectNameMasterConfig lineType, sheetName, endRow
    If sheetName = "" Then Exit Function

    If CommonNormalizeText(lineType) = ZAIRAISEN_NAME Then
        Dim result As Collection
        Set result = New Collection
        AddUnitPriceProjectNames result, LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName, ZAIRAISEN_PROJECT_NAME_FIRST_END_ROW)
        AddUnitPriceProjectNames result, LoadUnitPriceProjectNamesByAdo2(sourceFilePath, sheetName, ZAIRAISEN_PROJECT_NAME_SKIP_ROW + 1, endRow)
        Set LoadUnitPriceProjectNames = result
    Else
        Set LoadUnitPriceProjectNames = LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName, endRow)
    End If
    Exit Function
ErrorHandler:
    Set LoadUnitPriceProjectNames = Nothing
End Function

Private Sub ResolveUnitPriceProjectNameMasterConfig(ByVal lineType As String, _
                                                    ByRef sheetName As String, _
                                                    ByRef endRow As Long)
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_NAME
            sheetName = PROJECT_MASTER_SHEET_NAME
            endRow = ZAIRAISEN_PROJECT_NAME_SECOND_END_ROW
        Case SHINKANSEN_NAME
            sheetName = SHINKANSEN_NAME & "_" & PROJECT_MASTER_SHEET_NAME
            endRow = SHINKANSEN_PROJECT_NAME_END_ROW
        Case Else
            sheetName = ""
            endRow = 0
    End Select
End Sub

Private Function LoadUnitPriceProjectNamesByAdo(ByVal sourceFilePath As String, _
                                                ByVal sheetName As String, _
                                                ByVal endRow As Long) As Collection
    Dim cn As Object, rs As Object, result As Collection
    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Set rs = cn.Execute("SELECT F1 FROM " & _
                        BuildAdoSheetRangeName(sheetName, "A", PROJECT_NAME_MASTER_START_ROW, "A", endRow) & _
                        " WHERE F1 IS NOT NULL")
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

Private Function LoadUnitPriceProjectNamesByAdo2(ByVal sourceFilePath As String, _
                                                 ByVal sheetName As String, _
                                                 ByVal startRow As Long, _
                                                 ByVal endRow As Long) As Collection
    Dim cn As Object, rs As Object, result As Collection
    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Set rs = cn.Execute("SELECT F1 FROM " & _
                        BuildAdoSheetRangeName(sheetName, "A", startRow, "A", endRow) & _
                        " WHERE F1 IS NOT NULL")
    Set result = New Collection
    Do Until rs.EOF
        Dim itemText As String
        itemText = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If itemText <> "" Then result.Add itemText
        rs.MoveNext
    Loop
    Set LoadUnitPriceProjectNamesByAdo2 = result
Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function
ErrorHandler:
    Set LoadUnitPriceProjectNamesByAdo2 = Nothing
    Resume Cleanup
End Function

Private Function BuildAdoSheetRangeName(ByVal sheetName As String, _
                                        ByVal startCol As String, ByVal startRow As Long, _
                                        ByVal endCol As String, ByVal endRow As Long) As String
    BuildAdoSheetRangeName = "[" & Replace$(sheetName, "]", "]]") & "$" & _
                             startCol & CStr(startRow) & ":" & endCol & CStr(endRow) & "]"
End Function

Private Sub WriteUnitPriceLineTypeValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(LINE_TYPE_LIST_COL & LIST_START_ROW).Resize(2, 1)
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).value = ZAIRAISEN_NAME
    listRange.Cells(2, 1).value = SHINKANSEN_NAME
    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL), listRange
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceKindValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(PRICE_KIND_LIST_COL & LIST_START_ROW).Resize(2, 1)
    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).value = INITIAL_PRICE_NAME
    listRange.Cells(2, 1).value = DESIGN_CHANGE_PRICE_NAME
    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL), listRange
    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                ByVal projectNames As Collection, _
                                                ByVal keepProjectName As Boolean)
    Dim maxRows As Long
    maxRows = Application.Max(1, projectNames.Count)
    Dim listRange As Range
    Set listRange = wsInfo.Range(PROJECT_NAME_LIST_COL & LIST_START_ROW).Resize(maxRows, 1)

    Dim currentProjectName As String
    currentProjectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).value))

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents

    Dim i As Long
    For i = 1 To projectNames.Count
        listRange.Cells(i, 1).value = projectNames(i)
    Next i

    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL), _
                             listRange.Resize(Application.Max(1, projectNames.Count))

    If Not keepProjectName Or currentProjectName = "" Or Not CollectionContainsText(projectNames, currentProjectName) Then
        wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    End If

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
End Sub

Private Sub ResetUnitPriceValidation(ByVal targetCell As Range, ByVal listRange As Range)
    Dim topLeft As Range
    Set topLeft = targetCell.MergeArea.Cells(1, 1)
    With topLeft.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
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

Private Function FindUnitPriceWorkbooks(ByVal folderPath As String, ByVal projectName As String) As Collection
    Dim normalizedProjectName As String
    normalizedProjectName = NormalizeMatchText(projectName)
    If normalizedProjectName = "" Then Exit Function

    Dim result As Collection
    Set result = New Collection

    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, fileName, UNIT_PRICE_FILE_KEYWORD, vbTextCompare) = 0 Then
                    If InStr(1, NormalizeMatchText(RemoveFileExtension(fileName)), normalizedProjectName, vbTextCompare) > 0 Then
                        Dim workbookPath As String
                        workbookPath = folderPath & "\" & fileName
                        If Not CollectionContainsText(result, workbookPath) Then result.Add workbookPath
                    End If
                End If
            End If
            fileName = Dir()
        Loop
    Next ext

    Set FindUnitPriceWorkbooks = result
End Function

Private Function FindPurchaseUnitPriceWorkbook(ByVal folderPath As String) As String
    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, fileName, UNIT_PRICE_FILE_KEYWORD, vbTextCompare) > 0 Then
                    FindPurchaseUnitPriceWorkbook = folderPath & "\" & fileName
                    Exit Function
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
End Function

Private Function FindWeldingUnitPriceWorkbook(ByVal folderPath As String) As String
    FindWeldingUnitPriceWorkbook = FindWorkbookByKeyword(folderPath, WELDING_FILE_KEYWORD)
    If FindWeldingUnitPriceWorkbook = "" Then
        FindWeldingUnitPriceWorkbook = FindWorkbookByKeyword(folderPath, WELDING_FILE_FALLBACK_KEYWORD)
    End If
End Function

Private Function FindWorkbookByKeyword(ByVal folderPath As String, ByVal keyword As String) As String
    Dim normalizedKeyword As String
    normalizedKeyword = NormalizeMatchText(keyword)
    If normalizedKeyword = "" Then Exit Function

    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, NormalizeMatchText(fileName), normalizedKeyword, vbTextCompare) > 0 Then
                    FindWorkbookByKeyword = folderPath & "\" & fileName
                    Exit Function
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
End Function

Private Function RemoveFileExtension(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then RemoveFileExtension = Left$(fileName, pos - 1) Else RemoveFileExtension = fileName
End Function

Private Function IsPurchaseUnitPriceProjectName(ByVal projectName As String) As Boolean
    IsPurchaseUnitPriceProjectName = (InStr(1, NormalizeMatchText(projectName), NormalizeMatchText(UNIT_PRICE_FILE_KEYWORD), vbTextCompare) > 0 Or _
                                      InStr(1, NormalizeMatchText(projectName), NormalizeMatchText("購入充当"), vbTextCompare) > 0)
End Function

Private Function FindChildFolderByKey(ByVal parentFolder As String, _
                                      ByVal keyText As String, _
                                      ByVal ignoreLeadingNumbers As Boolean) As String
    Dim folderName As String
    folderName = Dir(parentFolder & "\*", vbDirectory)
    Do While folderName <> ""
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(parentFolder & "\" & folderName) And vbDirectory) = vbDirectory Then
                If FolderTextMatches(keyText, NormalizeFolderName(folderName, ignoreLeadingNumbers)) Then
                    FindChildFolderByKey = parentFolder & "\" & folderName
                    Exit Function
                End If
            End If
        End If
        folderName = Dir()
    Loop
End Function

Private Function FolderTextMatches(ByVal keyText As String, ByVal folderText As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(keyText)
    b = NormalizeMatchText(folderText)
    If a = "" Or b = "" Then Exit Function
    FolderTextMatches = (StrComp(a, b, vbTextCompare) = 0 Or _
                         Left$(a, Len(b)) = b Or Left$(b, Len(a)) = a Or _
                         InStr(1, a, b, vbTextCompare) > 0 Or InStr(1, b, a, vbTextCompare) > 0)
End Function

Private Function MasterTextMatches(ByVal masterText As String, ByVal searchText As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(masterText)
    b = NormalizeMatchText(searchText)
    If a = "" Or b = "" Then Exit Function
    MasterTextMatches = (StrComp(a, b, vbTextCompare) = 0 Or Left$(a, Len(b)) = b)
End Function

Private Function NormalizeMatchText(ByVal value As String) As String
    NormalizeMatchText = CommonRemoveAllSpaces(CommonNormalizeText(value))
End Function

Private Function NormalizeFolderName(ByVal folderName As String, ByVal ignoreLeadingNumbers As Boolean) As String
    Dim result As String
    result = folderName
    If ignoreLeadingNumbers Then
        Do While Len(result) > 0 And Mid$(result, 1, 1) Like "[0-9]"
            result = Mid$(result, 2)
        Loop
        Do While Len(result) > 0 And (Left$(result, 1) = "_" Or Left$(result, 1) = "-" Or Left$(result, 1) = " ")
            result = Mid$(result, 2)
        Loop
    End If
    NormalizeFolderName = result
End Function

Private Function GetUnitPriceDataRootPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath <> "" Then
        GetUnitPriceDataRootPath = fso.BuildPath(fso.GetParentFolderName(masterFilePath), _
                                   UnitPriceMasterFolderText() & "\" & UNIT_PRICE_DATA_FOLDER)
        Exit Function
    End If
    If Len(ThisWorkbook.Path) > 0 Then
        GetUnitPriceDataRootPath = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                                   MasterDataFolderText() & "\" & UnitPriceMasterFolderText() & "\" & UNIT_PRICE_DATA_FOLDER)
    End If
End Function

Private Function FirstExistingFilePath(ByVal candidates As Collection) As String
    Dim candidate As Variant, found As Boolean
    For Each candidate In candidates
        found = False
        On Error Resume Next
        found = (Len(Dir(CStr(candidate), vbNormal)) > 0)
        On Error GoTo 0
        If found Then
            FirstExistingFilePath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Function BuildImportCompleteMessage(ByVal selectedSheetNames As Collection, _
                                            ByVal sourceFilePath As String, _
                                            ByVal purchaseSheetName As String, _
                                            ByVal weldingSheetName As String) As String
    BuildImportCompleteMessage = CStr(selectedSheetNames.Count) & UiMsgImportCompleteCountSuffixText() & vbCrLf & sourceFilePath
    If Len(purchaseSheetName) > 0 Then
        BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & UiMsgPurchaseSheetCreatedPrefixText() & purchaseSheetName & UiMsgSheetCreatedSuffixText()
    End If
    If Len(weldingSheetName) > 0 Then
        BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & UiMsgWeldingSheetCreatedPrefixText() & weldingSheetName & UiMsgSheetCreatedSuffixText()
    End If
End Function

Private Function OrderInvoiceDocumentFolderText() As String
    OrderInvoiceDocumentFolderText = mod_common.CommonOrderInvoiceDocumentFolderText()
End Function

Private Function MasterDataFolderText() As String
    MasterDataFolderText = mod_common.CommonMasterDataFolderText()
End Function

Private Function UnitPriceMasterFolderText() As String
    Static cached As String
    If cached = "" Then cached = "単価マスタ"
    UnitPriceMasterFolderText = cached
End Function
