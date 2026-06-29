Option Explicit

' ????: CHANGELOG.md ??
' mod_Construction_LineMapping (split from mod_Construction_Order_Import)

Public Function BuildConstructionLineSheetMap() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim targetSheet As Worksheet
    For Each targetSheet In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(targetSheet) Then
            AddLineSheetAliases result, targetSheet.Name, targetSheet.Name
        End If
    Next targetSheet

    Dim lineNamePairs As Collection
    Set lineNamePairs = LoadProjectMasterLineNamePairs()
    If lineNamePairs Is Nothing Then
        LogCI "工事件名別マスタ未検出 -> 単価シート名の直接照合のみ"
        Set BuildConstructionLineSheetMap = result
        Exit Function
    End If

    EnsureProjectLineNameAliasMapsLoaded

    Dim pairItem As Variant
    For Each pairItem In lineNamePairs
        Dim unitPriceLineName As String
        Dim sourceLineName As String
        Dim actualSheetName As String

        unitPriceLineName = CStr(pairItem(0))
        sourceLineName = CStr(pairItem(1))

        actualSheetName = ResolveUnitPriceSheetName(result, unitPriceLineName, False)
        If actualSheetName = "" Then
            actualSheetName = ResolveUnitPriceSheetName(result, sourceLineName, False)
        End If
        If actualSheetName <> "" Then
            AddLineSheetAliases result, sourceLineName, actualSheetName
            AddLineSheetAliases result, unitPriceLineName, actualSheetName
        End If
    Next pairItem

    LogCI "線区名→単価シート対応数=" & result.Count
    Set BuildConstructionLineSheetMap = result
End Function

Public Sub AddLineSheetAliases(ByVal lineSheetMap As Object, _
                                ByVal sourceLineName As String, _
                                ByVal unitPriceSheetName As String)
    AddLineSheetAlias lineSheetMap, "W|" & NormalizeLineLookupText(sourceLineName, True), unitPriceSheetName
    AddLineSheetAlias lineSheetMap, "C|" & NormalizeLineLookupText(sourceLineName, False), unitPriceSheetName
End Sub

Public Sub AddLineSheetAlias(ByVal lineSheetMap As Object, _
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

Public Sub ClearProjectLineNameAliasCache()
    Set mProjectLineNameAliasMapWelding = Nothing
    Set mProjectLineNameAliasMapConstruction = Nothing
    Set mProjectLineNameReverseAliasWelding = Nothing
    Set mProjectMasterLineOrderRankMap = Nothing
End Sub

' ============================================================
' 在来線フォルダ各ファイルの F列(積算線区) 出現順を順位として返す。
' mod_MaterialPriceImport が単価シートの表示・取込・C24順を F列順に揃えるために使用する。
' 照合は線区名対応(エイリアス)と同じ正規化 NormalizeLineLookupText(_, False) を用いる。
' 該当無しは -1 を返す。
' ============================================================
Public Function GetProjectMasterLineOrderRank(ByVal lineName As String) As Long
    GetProjectMasterLineOrderRank = -1
    EnsureProjectMasterLineOrderMapLoaded
    If mProjectMasterLineOrderRankMap Is Nothing Then Exit Function

    Dim key As String
    key = NormalizeLineLookupText(lineName, False)
    If Len(key) = 0 Then Exit Function

    If mProjectMasterLineOrderRankMap.Exists(key) Then
        GetProjectMasterLineOrderRank = CLng(mProjectMasterLineOrderRankMap(key))
    End If
End Function

' 順位辞書のキャッシュを明示的に破棄する(取込前に呼んで最新化したい場合)。
Public Sub ClearProjectMasterLineOrderCache()
    Set mProjectMasterLineOrderRankMap = Nothing
End Sub

Public Sub EnsureProjectMasterLineOrderMapLoaded()
    If Not mProjectMasterLineOrderRankMap Is Nothing Then Exit Sub

    Set mProjectMasterLineOrderRankMap = CreateObject("Scripting.Dictionary")
    mProjectMasterLineOrderRankMap.CompareMode = vbTextCompare

    Dim lineNamePairs As Collection
    Set lineNamePairs = LoadProjectMasterLineNamePairs()
    If lineNamePairs Is Nothing Then Exit Sub

    Dim nextRank As Long
    nextRank = 0

    Dim pairItem As Variant
    For Each pairItem In lineNamePairs
        ' 同一マスタ行の F列(積算線区)と G列(施工指示書記載線区名) は
        ' 同一順位を共有させ、マスタの行順をそのまま順位とする。
        ' F列を主キー、G列を補助キー(単価シート同がG列と一致する場合のフォールバック)とする。
        Dim registeredF As Boolean
        registeredF = RegisterProjectMasterLineOrderKeyAtRank(CStr(pairItem(0)), nextRank)
        Dim registeredG As Boolean
        registeredG = RegisterProjectMasterLineOrderKeyAtRank(CStr(pairItem(1)), nextRank)
        If registeredF Or registeredG Then nextRank = nextRank + 1
    Next pairItem

    LogCI "単価シート並順キー数(F列順)=" & mProjectMasterLineOrderRankMap.Count
End Sub

' 指定 rank でキーを登録する。新規登録したら True、既出キーなら False を返す。
Public Function RegisterProjectMasterLineOrderKeyAtRank(ByVal lineName As String, _
                                                         ByVal rankValue As Long) As Boolean
    Dim key As String
    key = NormalizeLineLookupText(lineName, False)
    If Len(key) = 0 Then Exit Function
    If mProjectMasterLineOrderRankMap.Exists(key) Then Exit Function

    mProjectMasterLineOrderRankMap.Add key, rankValue
    RegisterProjectMasterLineOrderKeyAtRank = True
End Function

' 工事件名別マスタ F列(積算線区)・G列(施工指示書記載線区名) のペアを返す。
' C21 で指定された1ファイルだけでなく、C20 フォルダ内の全 .xlsx(①軌道整備他 等)を走査する。
Public Function LoadProjectMasterLineNamePairs() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim masterFiles As Collection
    Set masterFiles = CollectProjectLineMasterWorkbookPaths()
    If masterFiles Is Nothing Or masterFiles.Count = 0 Then Exit Function

    Dim masterPath As Variant
    Dim loadedFileCount As Long
    For Each masterPath In masterFiles
        If AppendProjectMasterLineNamePairsFromWorkbook(CStr(masterPath), result) Then
            loadedFileCount = loadedFileCount + 1
        End If
    Next masterPath

    LogCI "工事件名別マスタ 線区名ペア数=" & result.Count & _
          " 読込ファイル数=" & loadedFileCount & "/" & masterFiles.Count
    Set LoadProjectMasterLineNamePairs = result
End Function

Public Function CollectProjectLineMasterWorkbookPaths() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim masterFolder As String
    masterFolder = ResolveProjectLineMasterFolderPath()
    If masterFolder <> "" Then
        AppendProjectMasterExcelFilesInFolder masterFolder, result
        LogCI "工事件名別マスタ探索 folder=[" & masterFolder & "] files=" & result.Count
    End If

    If result.Count = 0 Then
        Dim singlePath As String
        singlePath = ResolveProjectLineMasterPath()
        If singlePath <> "" Then result.Add singlePath
    End If

    If result.Count > 0 Then Set CollectProjectLineMasterWorkbookPaths = result
End Function

Public Function ResolveProjectLineMasterFolderPath() As String
    Dim lineType As String
    lineType = GetBasicInfoLineTypeText()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim documentRoots As Collection
    Set documentRoots = New Collection
    CollectDocumentMasterDataRoots documentRoots

    Dim documentRoot As Variant
    Dim masterFolder As String

    If lineType <> "" Then
        For Each documentRoot In documentRoots
            masterFolder = fso.BuildPath(CStr(documentRoot), MASTER_DATA_FOLDER & "\" & lineType)
            If fso.FolderExists(masterFolder) Then
                ResolveProjectLineMasterFolderPath = masterFolder
                Exit Function
            End If
        Next documentRoot
    End If

    For Each documentRoot In documentRoots
        masterFolder = fso.BuildPath(CStr(documentRoot), _
            MASTER_DATA_FOLDER & "\" & mod_Construction_OutputLayout.CombinedLineTypeMasterFolderText())
        If fso.FolderExists(masterFolder) Then
            LogCI "工事件名別マスタフォルダ(既定) path=[" & masterFolder & "]"
            ResolveProjectLineMasterFolderPath = masterFolder
            Exit Function
        End If
    Next documentRoot
End Function

Public Function GetBasicInfoLineTypeText() As String
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function
    GetBasicInfoLineTypeText = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).value))
End Function

Public Sub CollectDocumentMasterDataRoots(ByVal documentRoots As Collection)
    If documentRoots Is Nothing Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(ThisWorkbook.Path) > 0 Then
        AddUniqueText documentRoots, ThisWorkbook.Path
        AddUniqueText documentRoots, fso.GetParentFolderName(ThisWorkbook.Path)
    End If

    Dim managerMasterPath As String
    managerMasterPath = mod_Construction_Import_Load.ResolveMasterFilePath()
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
                     mod_Construction_OutputLayout.OrderInvoiceDocumentFolderTextCI()
    End If
End Sub

Public Sub AppendProjectMasterExcelFilesInFolder(ByVal folderPath As String, _
                                                  ByVal filePaths As Collection)
    If filePaths Is Nothing Or Len(folderPath) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub

    Dim sourceFile As Object
    For Each sourceFile In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(sourceFile.Name)) = "xlsx" Then
            If Left$(sourceFile.Name, 2) <> "~$" Then
                AddUniqueText filePaths, sourceFile.Path
            End If
        End If
    Next sourceFile
End Sub

Public Function AppendProjectMasterLineNamePairsFromWorkbook(ByVal masterPath As String, _
                                                              ByVal lineNamePairs As Collection) As Boolean
    If lineNamePairs Is Nothing Or Len(masterPath) = 0 Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then
        LogCI "工事件名別マスタへADO接続できない path=[" & masterPath & "]"
        Exit Function
    End If

    Dim pairCountBefore As Long
    pairCountBefore = lineNamePairs.Count

    On Error GoTo Cleanup

    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim sheetName As Variant
    Dim recordset As Object
    For Each sheetName In sheetNames
        Set recordset = CreateObject("ADODB.Recordset")
        recordset.Open "SELECT [F6], [F7] FROM " & _
                       mod_Construction_OutputLayout.BuildAdoSheetTableName(CStr(sheetName)), connection, 0, 1, 1

        Dim rowNumber As Long
        rowNumber = 1
        Do Until recordset.EOF
            Dim unitPriceLineName As String
            Dim sourceLineName As String

            unitPriceLineName = CommonNzText(recordset.Fields(0).value)
            If rowNumber >= PROJECT_MASTER_START_ROW And Trim$(unitPriceLineName) <> "" Then
                sourceLineName = CommonNzText(recordset.Fields(1).value)
                If Trim$(sourceLineName) = "" Then sourceLineName = unitPriceLineName
                lineNamePairs.Add Array(unitPriceLineName, sourceLineName)
            End If
            recordset.MoveNext
            rowNumber = rowNumber + 1
        Loop

        CommonCloseAdoRecordset recordset
        Set recordset = Nothing
    Next sheetName

    AppendProjectMasterLineNamePairsFromWorkbook = True
    LogCI "工事件名別マスタ読込 path=[" & masterPath & "] 追加ペア=" & _
          CStr(lineNamePairs.Count - pairCountBefore)

Cleanup:
    If Err.Number <> 0 Then
        LogCI "工事件名別マスタ線区名読込エラー path=[" & masterPath & "] Err " & _
              Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    Err.Clear
End Function

Public Sub EnsureProjectLineNameAliasMapsLoaded()
    If Not mProjectLineNameAliasMapWelding Is Nothing Then Exit Sub

    Set mProjectLineNameAliasMapWelding = CreateObject("Scripting.Dictionary")
    mProjectLineNameAliasMapWelding.CompareMode = vbTextCompare
    Set mProjectLineNameAliasMapConstruction = CreateObject("Scripting.Dictionary")
    mProjectLineNameAliasMapConstruction.CompareMode = vbTextCompare
    Set mProjectLineNameReverseAliasWelding = CreateObject("Scripting.Dictionary")
    mProjectLineNameReverseAliasWelding.CompareMode = vbTextCompare

    Dim lineNamePairs As Collection
    Set lineNamePairs = LoadProjectMasterLineNamePairs()
    If lineNamePairs Is Nothing Then Exit Sub

    Dim pairItem As Variant
    For Each pairItem In lineNamePairs
        RegisterProjectLineNameAliasPair CStr(pairItem(0)), CStr(pairItem(1))
    Next pairItem

    LogCI "線区名エイリアス(溶接)=" & mProjectLineNameAliasMapWelding.Count & _
          " (工事)=" & mProjectLineNameAliasMapConstruction.Count
End Sub

Public Sub RegisterProjectLineNameAliasPair(ByVal unitPriceLineName As String, _
                                             ByVal sourceLineName As String)
    Dim normalizedUnitW As String
    Dim normalizedSourceW As String
    Dim normalizedUnitC As String
    Dim normalizedSourceC As String

    normalizedUnitW = NormalizeLineLookupText(unitPriceLineName, True)
    normalizedSourceW = NormalizeLineLookupText(sourceLineName, True)
    normalizedUnitC = NormalizeLineLookupText(unitPriceLineName, False)
    normalizedSourceC = NormalizeLineLookupText(sourceLineName, False)

    If normalizedSourceW <> "" And normalizedUnitW <> "" Then
        If StrComp(normalizedSourceW, normalizedUnitW, vbTextCompare) <> 0 Then
            AddProjectLineNameAlias mProjectLineNameAliasMapWelding, normalizedSourceW, normalizedUnitW
            AddProjectLineNameReverseAlias mProjectLineNameReverseAliasWelding, normalizedUnitW, normalizedSourceW
        End If
    End If

    If normalizedSourceC <> "" And normalizedUnitC <> "" Then
        If StrComp(normalizedSourceC, normalizedUnitC, vbTextCompare) <> 0 Then
            AddProjectLineNameAlias mProjectLineNameAliasMapConstruction, normalizedSourceC, normalizedUnitC
        End If
    End If

    RegisterSynthesizedConstructionLineAliases unitPriceLineName

    ' 施工指示書側の正規化差(例: (溶接指示書用)付き)でも溶接照合できるよう、工事正規化キーを溶接マップへ登録
    If normalizedSourceC <> "" And normalizedUnitW <> "" And _
       StrComp(normalizedSourceC, normalizedUnitW, vbTextCompare) <> 0 Then
        If StrComp(normalizedSourceC, normalizedSourceW, vbTextCompare) <> 0 Then
            AddProjectLineNameAlias mProjectLineNameAliasMapWelding, normalizedSourceC, normalizedUnitW
            AddProjectLineNameReverseAlias mProjectLineNameReverseAliasWelding, normalizedUnitW, normalizedSourceC
        End If
    End If
End Sub

Public Sub AddProjectLineNameAlias(ByVal aliasMap As Object, _
                                    ByVal sourceLineName As String, _
                                    ByVal unitPriceLineName As String)
    If aliasMap Is Nothing Then Exit Sub
    If Len(sourceLineName) = 0 Or Len(unitPriceLineName) = 0 Then Exit Sub

    If Not aliasMap.Exists(sourceLineName) Then
        aliasMap.Add sourceLineName, unitPriceLineName
    ElseIf StrComp(CStr(aliasMap(sourceLineName)), unitPriceLineName, vbTextCompare) <> 0 Then
        LogCI "線区名エイリアス重複 source=[" & sourceLineName & "] first=[" & _
              CStr(aliasMap(sourceLineName)) & "] ignored=[" & unitPriceLineName & "]"
    End If
End Sub

Public Sub AddProjectLineNameReverseAlias(ByVal reverseMap As Object, _
                                           ByVal unitLineName As String, _
                                           ByVal sourceLineName As String)
    If reverseMap Is Nothing Then Exit Sub
    If Len(unitLineName) = 0 Or Len(sourceLineName) = 0 Then Exit Sub

    Dim sources As Collection
    If reverseMap.Exists(unitLineName) Then
        Set sources = reverseMap(unitLineName)
    Else
        Set sources = New Collection
        reverseMap.Add unitLineName, sources
    End If

    Dim item As Variant
    For Each item In sources
        If StrComp(CStr(item), sourceLineName, vbTextCompare) = 0 Then Exit Sub
    Next item
    sources.Add sourceLineName
End Sub

Public Function ResolveProjectLineNameAlias(ByVal normalizedLineName As String, _
                                             ByVal isWeldingSheet As Boolean) As String
    If Len(normalizedLineName) = 0 Then Exit Function

    EnsureProjectLineNameAliasMapsLoaded

    Dim aliasMap As Object
    If isWeldingSheet Then
        Set aliasMap = mProjectLineNameAliasMapWelding
    Else
        Set aliasMap = mProjectLineNameAliasMapConstruction
    End If

    If Not aliasMap Is Nothing And aliasMap.Exists(normalizedLineName) Then
        ResolveProjectLineNameAlias = CStr(aliasMap(normalizedLineName))
    Else
        ResolveProjectLineNameAlias = normalizedLineName
    End If
End Function

Public Function ResolveWeldingLineSectionAlias(ByVal normalizedLineName As String) As String
    ResolveWeldingLineSectionAlias = ResolveProjectLineNameAlias(normalizedLineName, True)
End Function

Public Function FindWeldingPriceRecordKey(ByVal priceRows As Object, _
                                           ByVal lineText As String, _
                                           ByVal seiriValue As Variant) As String
    If priceRows Is Nothing Then Exit Function

    Dim primaryKey As String
    primaryKey = mod_Construction_OutputLayout.BuildWeldingLineSeiriLookupKey(lineText, seiriValue, True)
    If primaryKey <> "" And priceRows.Exists(primaryKey) Then
        FindWeldingPriceRecordKey = primaryKey
        Exit Function
    End If

    Dim fallbackKey As String
    fallbackKey = mod_Construction_OutputLayout.BuildWeldingLineSeiriLookupKey(lineText, seiriValue, False)
    If fallbackKey <> "" And StrComp(fallbackKey, primaryKey, vbTextCompare) <> 0 Then
        If priceRows.Exists(fallbackKey) Then FindWeldingPriceRecordKey = fallbackKey
    End If
End Function

Public Function ResolveUnitPriceSheetName(ByVal lineSheetMap As Object, _
                                           ByVal importedLineName As String, _
                                           Optional ByVal isWeldingSheet As Boolean = False) As String
    Dim candidates As Collection
    Set candidates = CollectLineLookupCandidates(importedLineName, isWeldingSheet)
    If candidates Is Nothing Or candidates.Count = 0 Then Exit Function

    If Not lineSheetMap Is Nothing Then
        Dim keyPrefix As String
        If isWeldingSheet Then
            keyPrefix = "W|"
        Else
            keyPrefix = "C|"
        End If

        Dim candidateItem As Variant
        For Each candidateItem In candidates
            Dim mapKey As String
            mapKey = keyPrefix & CStr(candidateItem)
            If Len(mapKey) > 2 And lineSheetMap.Exists(mapKey) Then
                ResolveUnitPriceSheetName = CStr(lineSheetMap(mapKey))
                Exit Function
            End If
        Next candidateItem
    End If

    ResolveUnitPriceSheetName = FindImportedUnitPriceSheetName(importedLineName, isWeldingSheet)
End Function

Public Function CollectLineLookupCandidates(ByVal lineName As String, _
                                             ByVal isWeldingSheet As Boolean) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim normalized As String
    normalized = NormalizeLineLookupText(lineName, isWeldingSheet)
    AddUniqueLineLookupCandidate result, normalized

    EnsureProjectLineNameAliasMapsLoaded

    Dim aliased As String
    aliased = ResolveProjectLineNameAlias(normalized, isWeldingSheet)
    AddUniqueLineLookupCandidate result, aliased

    If Not isWeldingSheet Then
        Dim synthesized As Collection
        Dim synthesizedItem As Variant
        Set synthesized = SynthesizeConstructionLineNameVariants(lineName)
        For Each synthesizedItem In synthesized
            AddUniqueLineLookupCandidate result, NormalizeLineLookupText(CStr(synthesizedItem), False)
            AddUniqueLineLookupCandidate result, _
                ResolveProjectLineNameAlias(NormalizeLineLookupText(CStr(synthesizedItem), False), False)
        Next synthesizedItem
    End If

    Set CollectLineLookupCandidates = result
End Function

Public Sub AddUniqueLineLookupCandidate(ByVal candidates As Collection, ByVal candidateText As String)
    If candidates Is Nothing Then Exit Sub
    If Len(candidateText) = 0 Then Exit Sub

    Dim item As Variant
    For Each item In candidates
        If StrComp(CStr(item), candidateText, vbTextCompare) = 0 Then Exit Sub
    Next item
    candidates.Add candidateText
End Sub

' 積算線区(F列)から、G列未記載でも施工指示書に現れやすい契約線区名の略称を合成する
Public Sub RegisterSynthesizedConstructionLineAliases(ByVal unitPriceLineName As String)
    If mProjectLineNameAliasMapConstruction Is Nothing Then Exit Sub
    If Len(Trim$(unitPriceLineName)) = 0 Then Exit Sub

    Dim normalizedUnit As String
    normalizedUnit = NormalizeLineLookupText(unitPriceLineName, False)
    If Len(normalizedUnit) = 0 Then Exit Sub

    Dim variants As Collection
    Dim variantItem As Variant
    Set variants = SynthesizeConstructionLineNameVariants(unitPriceLineName)
    For Each variantItem In variants
        Dim normalizedVariant As String
        normalizedVariant = NormalizeLineLookupText(CStr(variantItem), False)
        If Len(normalizedVariant) > 0 Then
            If StrComp(normalizedVariant, normalizedUnit, vbTextCompare) <> 0 Then
                AddProjectLineNameAlias mProjectLineNameAliasMapConstruction, normalizedVariant, normalizedUnit
            End If
        End If
    Next variantItem
End Sub

Public Function SynthesizeConstructionLineNameVariants(ByVal unitPriceLineName As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim sourceText As String
    sourceText = CommonRemoveAllSpaces(CommonNormalizeText(unitPriceLineName))
    If Len(sourceText) = 0 Then
        Set SynthesizeConstructionLineNameVariants = result
        Exit Function
    End If

    Dim honSenText As String
    honSenText = ChrW$(&H672C) & ChrW$(&H7DDA)
    Dim tokaiHonSenText As String
    tokaiHonSenText = ChrW$(&H6771) & ChrW$(&H6D77) & ChrW$(&H9053) & honSenText
    Dim tokaiText As String
    tokaiText = ChrW$(&H6771) & ChrW$(&H6D77) & ChrW$(&H9053)

    Dim honSenVariant As String
    honSenVariant = Replace$(sourceText, honSenText, ChrW$(&H672C), , , vbTextCompare)
    If StrComp(honSenVariant, sourceText, vbTextCompare) <> 0 Then result.Add honSenVariant

    Dim tokaiVariant As String
    tokaiVariant = Replace$(sourceText, tokaiHonSenText, tokaiText, , , vbTextCompare)
    If StrComp(tokaiVariant, sourceText, vbTextCompare) <> 0 Then result.Add tokaiVariant

    Dim segmentSeparator As String
    segmentSeparator = ChrW$(&H30FB)
    Dim segments() As String
    segments = Split(sourceText, segmentSeparator)
    If UBound(segments) >= LBound(segments) Then
        Dim segmentIndex As Long
        Dim changed As Boolean
        changed = False
        For segmentIndex = LBound(segments) To UBound(segments)
            Dim segmentText As String
            segmentText = segments(segmentIndex)
            If segmentIndex < UBound(segments) Then
                If Len(segmentText) > 1 Then
                    If Right$(segmentText, 1) = ChrW$(&H7DDA) Then
                        If InStr(1, segmentText, ChrW$(&H9023) & ChrW$(&H7D61) & ChrW$(&H7DDA), vbTextCompare) = 0 And _
                           InStr(1, segmentText, ChrW$(&H8CA) & ChrW$(&H7269) & ChrW$(&H7DDA), vbTextCompare) = 0 Then
                            segments(segmentIndex) = Left$(segmentText, Len(segmentText) - 1)
                            changed = True
                        End If
                    End If
                End If
            ElseIf segmentIndex = UBound(segments) And UBound(segments) = LBound(segments) Then
                If Right$(segmentText, Len(honSenText)) = honSenText Then
                    segments(segmentIndex) = Left$(segmentText, Len(segmentText) - Len(ChrW$(&H7DDA)))
                    changed = True
                End If
            End If
        Next segmentIndex
        If changed Then result.Add Join$(segments, segmentSeparator)
    End If

    Set SynthesizeConstructionLineNameVariants = result
End Function

Public Function NormalizeLineLookupText(ByVal sourceText As String, _
                                         Optional ByVal isWeldingSheet As Boolean = False) As String
    Dim result As String
    result = CommonNormalizeText(sourceText)
    result = NormalizeLineLookupBracketChars(result)
    result = NormalizeLineLookupDigitChars(result)
    result = Replace$(result, ChrW$(&HFF65), ChrW$(&H30FB))
    If isWeldingSheet Then
        result = RemoveWeldingInstructionMarker(result)
    Else
        result = RemoveTrackDesignationMarker(result)
    End If
    NormalizeLineLookupText = CommonRemoveAllSpaces(result)
End Function

Public Function NormalizeLineLookupBracketChars(ByVal sourceText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long

    result = ""
    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        codePoint = AscW(ch)
        Select Case codePoint
            Case &HFF08: ch = ChrW$(&H28)
            Case &HFF09: ch = ChrW$(&H29)
            Case &HFF3B: ch = ChrW$(&H5B)
            Case &HFF3D: ch = ChrW$(&H5D)
        End Select
        result = result & ch
    Next i
    NormalizeLineLookupBracketChars = result
End Function

Public Function NormalizeLineLookupDigitChars(ByVal sourceText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long

    result = ""
    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        codePoint = AscW(ch)
        If codePoint >= &HFF10 And codePoint <= &HFF19 Then
            ch = ChrW$(codePoint - &HFF10 + AscW("0"))
        End If
        result = result & ch
    Next i
    NormalizeLineLookupDigitChars = result
End Function

Public Function RemoveWeldingInstructionMarker(ByVal sourceText As String) As String
    Static halfWidthMarker As String
    Static fullWidthMarker As String

    If Len(halfWidthMarker) = 0 Then
        halfWidthMarker = ChrW$(&H28) & ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H6307) & _
                          ChrW$(&H793A) & ChrW$(&H66F8) & ChrW$(&H7528) & ChrW$(&H29)
        fullWidthMarker = ChrW$(&HFF08) & ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H6307) & _
                          ChrW$(&H793A) & ChrW$(&H66F8) & ChrW$(&H7528) & ChrW$(&HFF09)
    End If

    Dim result As String
    result = sourceText
    result = Replace$(result, halfWidthMarker, "", , , vbTextCompare)
    result = Replace$(result, fullWidthMarker, "", , , vbTextCompare)
    RemoveWeldingInstructionMarker = result
End Function

Public Function RemoveTrackDesignationMarker(ByVal sourceText As String) As String
    Static halfWidthMarker As String
    Static fullWidthMarker As String

    If Len(halfWidthMarker) = 0 Then
        halfWidthMarker = ChrW$(&H28) & ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H29)
        fullWidthMarker = ChrW$(&HFF08) & ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&HFF09)
    End If

    Dim result As String
    result = sourceText
    result = Replace$(result, halfWidthMarker, "", , , vbTextCompare)
    result = Replace$(result, fullWidthMarker, "", , , vbTextCompare)
    RemoveTrackDesignationMarker = result
End Function

' 線区テキストに(軌道)マーカーが含まれるか
Public Function LineTextHasTrackDesignation(ByVal lineText As String) As Boolean
    Dim normalized As String
    normalized = CommonNormalizeText(lineText)
    LineTextHasTrackDesignation = (RemoveTrackDesignationMarker(normalized) <> normalized)
End Function

' 整理番号が単価マスタに未登録の行を、線区シート別に収集する(同一整理番号は1件のみ)
Public Sub CollectMissingSeiriForLineSheet(ByVal pendingByLineSheet As Object, _
                                            ByVal lineSheetName As String, _
                                            ByVal recordKey As String, _
                                            ByVal seiriValue As Variant, _
                                            ByVal typeValue As Variant, _
                                            ByVal unitValue As Variant)
    If pendingByLineSheet Is Nothing Then Exit Sub
    If lineSheetName = "" Or recordKey = "" Then Exit Sub

    Dim sheetPending As Object
    If pendingByLineSheet.Exists(lineSheetName) Then
        Set sheetPending = pendingByLineSheet(lineSheetName)
    Else
        Set sheetPending = CreateObject("Scripting.Dictionary")
        sheetPending.CompareMode = vbTextCompare
        pendingByLineSheet.Add lineSheetName, sheetPending
    End If

    If Not sheetPending.Exists(recordKey) Then
        sheetPending.Add recordKey, Array(seiriValue, typeValue, unitValue)
    End If
End Sub

' 収集した未登録整理番号を各線区シートへ反映する
Public Sub RegisterMissingSeiriToLineSheets(ByVal pendingByLineSheet As Object)
    If pendingByLineSheet Is Nothing Then Exit Sub
    If pendingByLineSheet.Count = 0 Then Exit Sub

    ' 線区シートのB列書き込みが Worksheet_Change を誘発しないよう局所的に抑制する
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreEvents

    Dim sheetName As Variant
    For Each sheetName In pendingByLineSheet.Keys
        AppendMissingSeiriToLineSheet CStr(sheetName), pendingByLineSheet(sheetName)
    Next sheetName

RestoreEvents:
    If Err.Number <> 0 Then
        LogCI "整理番号追記エラー Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    Application.EnableEvents = prevEvents
End Sub

' 1つの線区シートに対し、既存に無い整理番号のみ最下部へ追記し、単価欄(E/F)を黄色化する
Public Sub AppendMissingSeiriToLineSheet(ByVal lineSheetName As String, _
                                          ByVal sheetPending As Object)
    If sheetPending Is Nothing Then Exit Sub
    If sheetPending.Count = 0 Then Exit Sub

    Dim lineWs As Worksheet
    On Error Resume Next
    Set lineWs = ThisWorkbook.worksheets(lineSheetName)
    On Error GoTo 0
    If lineWs Is Nothing Then
        LogCI "線区シート[" & lineSheetName & "] が見つからないため追記をスキップ"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = lineWs.Cells(lineWs.rows.Count, COL_SEIRI).End(xlUp).Row
    If lastRow < UNIT_PRICE_DATA_START_ROW - 1 Then lastRow = UNIT_PRICE_DATA_START_ROW - 1

    ' 既存整理番号の集合(重複登録防止)
    Dim existingKeys As Object
    Set existingKeys = CreateObject("Scripting.Dictionary")
    existingKeys.CompareMode = vbTextCompare

    Dim rr As Long
    Dim existingKey As String
    For rr = UNIT_PRICE_DATA_START_ROW To lastRow
        existingKey = NormalizeRecordKey(lineWs.Cells(rr, COL_SEIRI).value)
        If existingKey <> "" Then
            If Not existingKeys.Exists(existingKey) Then existingKeys.Add existingKey, True
        End If
    Next rr

    ' 既存に無い分のみ追記対象とする
    Dim rowsToAdd As Collection
    Set rowsToAdd = New Collection
    Dim k As Variant
    For Each k In sheetPending.Keys
        If Not existingKeys.Exists(CStr(k)) Then rowsToAdd.Add sheetPending(k)
    Next k
    If rowsToAdd.Count = 0 Then
        LogCI "線区シート[" & lineSheetName & "] 追記候補=" & sheetPending.Count & _
              "件は既存整理番号と重複のためスキップ"
        Exit Sub
    End If

    Dim firstRow As Long
    firstRow = lastRow + 1
    If firstRow < UNIT_PRICE_DATA_START_ROW Then firstRow = UNIT_PRICE_DATA_START_ROW

    Dim writeArr() As Variant
    ReDim writeArr(1 To rowsToAdd.Count, 1 To 4)   ' A(工種), B(整理番号), C(種別), D(単位)

    Dim workTypeLabel As String
    workTypeLabel = MissingSeiriWorkTypeLabelText()

    Dim idx As Long
    Dim rowData As Variant
    For idx = 1 To rowsToAdd.Count
        rowData = rowsToAdd(idx)
        writeArr(idx, 1) = workTypeLabel          ' A 工種
        writeArr(idx, 2) = rowData(0)             ' B 整理番号
        writeArr(idx, 3) = rowData(1)             ' C 種別
        writeArr(idx, 4) = rowData(2)             ' D 単位
    Next idx

    Dim lastAppendRow As Long
    lastAppendRow = firstRow + rowsToAdd.Count - 1

    lineWs.Range(lineWs.Cells(firstRow, UNIT_PRICE_WORK_KIND_COL), _
                 lineWs.Cells(lastAppendRow, UNIT_PRICE_UNIT_COL)).value = writeArr

    HighlightLineSheetPriceCellsUntilFilled lineWs, firstRow, lastAppendRow
    mod_VendorMaster.ApplyConstructionUnitPriceImportedRowDecorations _
        lineWs, firstRow, lastAppendRow

    LogCI "線区シート[" & lineSheetName & "] 整理番号未登録分を追記: " & _
          rowsToAdd.Count & "件 rows " & firstRow & "-" & lastAppendRow
End Sub

' 追記行の単価欄(E/F)を黄色で塗る。条件付き書式により値が入力されると自動解除される
Public Sub HighlightLineSheetPriceCellsUntilFilled(ByVal lineWs As Worksheet, _
                                                    ByVal firstRow As Long, _
                                                    ByVal lastRow As Long)
    If lineWs Is Nothing Then Exit Sub
    If lastRow < firstRow Then Exit Sub

    AddEmptyCellHighlightCondition lineWs, firstRow, lastRow, UNIT_PRICE_DAY_PRICE_COL
    AddEmptyCellHighlightCondition lineWs, firstRow, lastRow, UNIT_PRICE_NIGHT_PRICE_COL
End Sub

Public Sub AddEmptyCellHighlightCondition(ByVal lineWs As Worksheet, _
                                           ByVal firstRow As Long, _
                                           ByVal lastRow As Long, _
                                           ByVal columnIndex As Long)
    Dim target As Range
    Set target = lineWs.Range(lineWs.Cells(firstRow, columnIndex), _
                              lineWs.Cells(lastRow, columnIndex))

    Dim anchorAddress As String
    anchorAddress = lineWs.Cells(firstRow, columnIndex).Address(False, False)

    On Error Resume Next
    Dim fc As FormatCondition
    Set fc = target.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & anchorAddress & "=""""")
    If Not fc Is Nothing Then
        fc.Interior.Color = MISSING_SEIRI_FILL_COLOR
        fc.StopIfTrue = False
    End If
    On Error GoTo 0
End Sub

Public Function MissingSeiriWorkTypeLabelText() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H72EC) & ChrW$(&H81EA) & ChrW$(&H5DE5) & ChrW$(&H7A2E)
    End If
    MissingSeiriWorkTypeLabelText = cached
End Function

Public Function ResolveProjectLineMasterPath() As String
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
    CollectDocumentMasterDataRoots documentRoots

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

Public Function FindProjectMasterFile(ByVal masterFolder As String, _
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

Public Function NormalizeProjectMasterName(ByVal sourceText As String) As String
    Dim result As String
    result = CommonRemoveAllSpaces(CommonNormalizeText(sourceText))
    If LCase$(Right$(result, 5)) = ".xlsx" Then result = Left$(result, Len(result) - 5)
    result = Replace$(result, ChrW$(&H30FB), "")
    result = Replace$(result, ChrW$(&HFF65), "")
    NormalizeProjectMasterName = result
End Function

Public Sub AddUniqueText(ByVal values As Collection, ByVal newValue As String)
    If values Is Nothing Or Len(Trim$(newValue)) = 0 Then Exit Sub

    Dim item As Variant
    For Each item In values
        If StrComp(CStr(item), newValue, vbTextCompare) = 0 Then Exit Sub
    Next item

    values.Add newValue
End Sub

Public Function FindImportedUnitPriceSheetName(ByVal expectedSheetName As String, _
                                                Optional ByVal isWeldingSheet As Boolean = False) As String
    Dim candidates As Collection
    Set candidates = CollectLineLookupCandidates(expectedSheetName, isWeldingSheet)
    If candidates Is Nothing Or candidates.Count = 0 Then Exit Function

    Dim ws As Worksheet
    Dim candidateItem As Variant
    For Each ws In ThisWorkbook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(ws) Then
            Dim normalizedSheetName As String
            normalizedSheetName = NormalizeLineLookupText(ws.Name, isWeldingSheet)
            For Each candidateItem In candidates
                If StrComp(normalizedSheetName, CStr(candidateItem), vbTextCompare) = 0 Then
                    FindImportedUnitPriceSheetName = ws.Name
                    Exit Function
                End If
            Next candidateItem
        End If
    Next ws
End Function

Public Function GetUnitPriceRows(ByVal unitPriceSheetName As String, _
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

    If mod_WeldingUnitPrice.IsWeldingUnitPriceSheet(priceSheet) Then
        BuildWeldingUnitPriceRowCache priceSheet, result
    Else
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
    End If

    sheetPriceCaches.Add unitPriceSheetName, result
    Set GetUnitPriceRows = result
End Function

Public Function NormalizeRecordKey(ByVal value As Variant) As String
    NormalizeRecordKey = CommonRemoveAllSpaces(CommonNzText(value))
End Function

Public Function SelectDayNightPrice(ByVal dayNightText As String, _
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

Public Sub WritePriceComparison(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                                 ByVal unitPriceSheetName As String, _
                                 Optional ByVal includeGuidance As Boolean = True, _
                                 Optional ByVal guidanceDocumentName As String = "")
    WritePriceComparisonAtColumns _
        ws, rowIndex, unitPriceSheetName, mod_Construction_OutputLayout.OutputSheetCol(ws, COL_AUTO_PRICE), _
        mod_Construction_OutputLayout.OutputSheetCol(ws, COL_JR_PRICE), mod_Construction_OutputLayout.OutputSheetCol(ws, COL_PRICE_COMPARE), _
        mod_Construction_OutputLayout.OutputSheetCol(ws, COL_PRICE_GUIDANCE), includeGuidance
End Sub

Public Sub WritePriceComparisonAtColumns( _
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
                            guidanceSheetName = CommonNzText(ws.Cells(rowIndex, mod_Construction_OutputLayout.OutputSheetCol(ws, COL_LINE)).value)
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

Public Function ResolveGuidanceDocumentNameFromOutputSheet(ByVal ws As Worksheet) As String
    Dim mgrColumn As Long
    mgrColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "管理室")
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

Public Function ResolveGuidanceDocumentName(ByVal sourceA3Text As String, _
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

Public Function UnitPriceValuesMatch(ByVal leftValue As Variant, _
                                      ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then Exit Function
    If Len(Trim$(CommonNzText(leftValue))) = 0 Or Len(Trim$(CommonNzText(rightValue))) = 0 Then Exit Function
    If Not IsNumeric(leftValue) Or Not IsNumeric(rightValue) Then Exit Function

    UnitPriceValuesMatch = (Abs(CDbl(leftValue) - CDbl(rightValue)) < 0.0000001)
End Function

Public Sub SortWorksSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim colSide As Long
    Dim colWeld As Long
    Dim colLine As Long
    Dim colKind As Long
    Dim colDayNight As Long
    Dim colSeiri As Long
    colSide = mod_Construction_OutputLayout.OutputSheetCol(ws, COL_FLAG_SIDE)
    colWeld = mod_Construction_OutputLayout.OutputSheetCol(ws, COL_FLAG_WELD)
    colLine = mod_Construction_OutputLayout.OutputSheetCol(ws, COL_LINE)
    colKind = mod_Construction_OutputLayout.OutputSheetCol(ws, COL_KIND)
    colDayNight = mod_Construction_OutputLayout.OutputSheetCol(ws, COL_DAYNIGHT)
    colSeiri = mod_Construction_OutputLayout.OutputSheetSeiriColumn(ws)

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

Public Sub SortPurchaseSheet(ByVal ws As Worksheet, _
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

Public Sub FormatSheet(ByVal ws As Worksheet, _
                        Optional ByVal dataKeyColumn As Long = COL_SEIRI)
    FormatSheetAtColumns _
        ws, dataKeyColumn, COL_DAYNIGHT, COL_UNIT, COL_LINE, COL_MGR, _
        COL_JR_PRICE, COL_OUT_AMOUNT, COL_KIND, COL_AUTO_PRICE, _
        COL_AUTO_AMOUNT, COL_PRICE_COMPARE, COL_PRICE_GUIDANCE
End Sub

Public Sub FormatPurchaseNoticeSheet(ByVal ws As Worksheet)
    FormatSheetAtColumns _
        ws, PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_DAYNIGHT_COL, _
        PURCHASE_NOTICE_UNIT_COL, PURCHASE_NOTICE_LINE_COL, PURCHASE_NOTICE_MGR_COL, _
        PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_OUT_AMOUNT_COL, _
        PURCHASE_NOTICE_KIND_COL, PURCHASE_NOTICE_AUTO_PRICE_COL, _
        PURCHASE_NOTICE_AUTO_AMOUNT_COL, PURCHASE_NOTICE_PRICE_COMPARE_COL, _
        PURCHASE_NOTICE_PRICE_GUIDANCE_COL
End Sub

Public Sub FormatSheetAtColumns( _
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
    outputLastColumn = mod_Construction_Import_Load.GetOutputLastColumn( _
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

Public Sub RedrawOutputSheetDataBorders(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim seiriColumn As Long
    seiriColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "整理番号")
    If seiriColumn = 0 Then Exit Sub

    Dim kindColumn As Long
    kindColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "工種分類")
    If kindColumn = 0 Then Exit Sub

    Dim comparisonColumn As Long
    comparisonColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "単価比較")

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
    subconFirstCol = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstCol(ws)
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

Public Function GetWeldingPriceRowMap(ByVal weldingSheetName As String, _
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
        BuildWeldingUnitPriceRowCache priceSheet, Nothing, True, result
    End If

    vendorPriceCaches.Add cacheKey, result
    Set GetWeldingPriceRowMap = result
End Function

Public Function FindWeldingRailVendorDayColumn(ByVal priceSheet As Worksheet, _
                                                  ByVal vendorName As String) As Long
    Dim vendorKey As String
    vendorKey = mod_Construction_BasicTotals.NormalizeVendorPriceName(vendorName)
    If vendorKey = "" Then Exit Function

    Dim lastColumn As Long
    lastColumn = priceSheet.Cells(WELDING_PRICE_VENDOR_NAME_ROW, priceSheet.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = WELDING_PRICE_FIRST_RAIL_DAY_COL To lastColumn Step 2
        Dim headerName As String
        headerName = mod_Construction_BasicTotals.NormalizeVendorPriceName(CommonNzText(priceSheet.Cells(WELDING_PRICE_VENDOR_NAME_ROW, c).value))
        If headerName <> "" Then
            If StrComp(headerName, vendorKey, vbTextCompare) = 0 Then
                FindWeldingRailVendorDayColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Public Function LookupWeldingOutputVendorPrice( _
    ByVal weldingSheetName As String, _
    ByVal vendorPriceCaches As Object, _
    ByVal lineText As String, _
    ByVal seiriValue As Variant, _
    ByVal vendorName As String, _
    ByVal isWeldingVendorSlot As Boolean, _
    ByVal dayNightText As String) As Variant

    LookupWeldingOutputVendorPrice = Empty
    If weldingSheetName = "" Then Exit Function

    Dim rowMap As Object
    Set rowMap = GetWeldingPriceRowMap(weldingSheetName, vendorPriceCaches)
    If rowMap Is Nothing Then Exit Function

    Dim recordKey As String
    recordKey = FindWeldingPriceRecordKey(rowMap, lineText, seiriValue)
    If recordKey = "" Then Exit Function

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

Public Sub ApplyOutputSheetHeaderAutoFilter( _
    ByVal ws As Worksheet, _
    Optional ByVal dataKeyColumn As Long = 0, _
    Optional ByVal kindColumn As Long = COL_KIND, _
    Optional ByVal autoAmountColumn As Long = COL_AUTO_AMOUNT, _
    Optional ByVal comparisonColumn As Long = COL_PRICE_COMPARE, _
    Optional ByVal guidanceColumn As Long = COL_PRICE_GUIDANCE)

    If ws Is Nothing Then Exit Sub

    If dataKeyColumn = 0 Then dataKeyColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumn(ws)

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, dataKeyColumn)
    If lastRow < 1 Then lastRow = 1

    Dim lastCol As Long
    lastCol = mod_Construction_Import_Load.GetOutputLastColumn(ws, kindColumn, autoAmountColumn, comparisonColumn, guidanceColumn)

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
End Sub

Public Sub ApplyOutputSheetHeaderFreezePanes(ByVal ws As Worksheet)
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

Public Sub ApplyPriceGuidanceColumnLayout(ByVal ws As Worksheet)
    ApplyPriceGuidanceColumnLayoutAtColumns _
        ws, COL_PRICE_COMPARE, COL_PRICE_GUIDANCE
End Sub

Public Sub ApplyPriceGuidanceColumnLayoutAtColumns( _
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

Public Sub ApplyPurchaseNoticeColumnExclusions(ByVal ws As Worksheet)
    Dim excludedColumns As Range
    Set excludedColumns = Application.Union(ws.Columns(PURCHASE_NOTICE_MGR_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_PRICE_COL), _
                                            ws.Columns(PURCHASE_NOTICE_OUT_AMOUNT_COL))
    excludedColumns.EntireColumn.Hidden = True
End Sub

Public Function GetLastDataRow(ByVal ws As Worksheet, _
                                Optional ByVal dataKeyColumn As Long = 0) As Long
    If dataKeyColumn = 0 Then dataKeyColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumn(ws)

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

