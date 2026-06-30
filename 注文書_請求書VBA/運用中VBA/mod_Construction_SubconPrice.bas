Option Explicit

' ????: CHANGELOG.md ??
' mod_Construction_SubconPrice (split from mod_Construction_Order_Import)

Public Sub RefreshSubcontractorPriceColumnsCore(ByVal ws As Worksheet, _
                                            Optional ByVal changedRows As Collection = Nothing)
    If ws Is Nothing Then Exit Sub
    If Not mod_Construction_OutputLayout.IsConstructionVendorOutputSheet(ws) Then Exit Sub
    If mod_Construction_BasicTotals.FindHeaderColumn(ws, "整理番号") = 0 Then Exit Sub

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
    Dim savedAutoFilter As Object

    Set vendorNames = New Collection

    Set savedAutoFilter = mod_Construction_LineMapping.CaptureWorksheetAutoFilter(ws)
    mod_Construction_LineMapping.SuspendWorksheetAutoFilterView ws

    refreshStep = "CollectVendors"
    On Error Resume Next
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Set vendorNames = mod_Construction_BasicTotals.CollectSelectedSubcontractors(ws, lastRow)
    If Err.Number <> 0 Then GoTo RefreshSetupError
    Err.Clear

    refreshStep = "FindKindColumn"
    kindColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "工種分類")
    If Err.Number <> 0 Then GoTo RefreshSetupError
    If kindColumn = 0 Then GoTo RefreshExit
    Err.Clear

    refreshStep = "ResolveSubconFirstCol"
    subconFirstCol = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstColCore(ws)
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
            kindColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "工種分類")
            If Err.Number <> 0 Then GoTo RefreshSetupError
            If kindColumn = 0 Then GoTo RefreshExit
        End If
        Err.Clear
    End If
    On Error GoTo 0

    If vendorNames.Count = 0 Or lastRow < 2 Then
        lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
        If lastRow >= 2 Then
            Dim emptyVendors As New Collection
            WriteOutputTotalRows ws, emptyVendors, 0, 0
            RefreshOutputSheetVendorColumnColors ws, lastRow
        End If
        mod_Construction_BasicTotals.RefreshBasicInfoConstructionTotalsCore
        GoTo RefreshExit
    End If

    insertedColumnCount = vendorNames.Count * 2

    If Not layoutMatches Then
        refreshStep = "LayoutInsert"
        On Error Resume Next
        ws.Range(ws.Columns(subconFirstCol), _
                 ws.Columns(subconFirstCol + insertedColumnCount - 1)).Insert Shift:=xlToRight
        If Err.Number <> 0 Then GoTo RefreshSetupError

        lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
        If Err.Number <> 0 Then GoTo RefreshSetupError
        If lastRow < 2 Then
            On Error GoTo 0
            mod_Construction_BasicTotals.RefreshBasicInfoConstructionTotalsCore
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

        kindColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "工種分類")
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
    Set lineSheetMap = mod_Construction_LineMapping.BuildConstructionLineSheetMap()
    Err.Clear
    If lineSheetMap Is Nothing Then
        Set lineSheetMap = CreateObject("Scripting.Dictionary")
        lineSheetMap.CompareMode = vbTextCompare
    End If

    Set vendorPriceCaches = CreateObject("Scripting.Dictionary")
    vendorPriceCaches.CompareMode = vbTextCompare

    refreshStep = "PrepareApply"
    seiriColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumnCore(ws)
    dayNightColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_DAYNIGHT)
    lineColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_LINE)
    qtyColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_QTY)
    isWeldingSheet = mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws)
    Set vendorColumns = mod_Construction_OutputLayout.OutputSheetVendorColumnsCore(ws)
    partialUpdate = False
    If Not changedRows Is Nothing Then partialUpdate = (changedRows.Count > 0)
    If ws.AutoFilterMode Then partialUpdate = False
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
        LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
        matchedCount = 0
        lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
        ApplySubcontractorPricesBatch ws, lastRow, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
        If Err.Number <> 0 Then
            If coreErrNo = 0 Then
                coreErrNo = Err.Number
                coreErrDesc = Err.Description
            End If
            LogCI "RefreshSubcontractorPriceColumnsCore step=ApplyPricesRetry Err " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
    ElseIf partialUpdate And layoutMatches And matchedCount = 0 Then
        lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
        ApplySubcontractorPricesBatch ws, lastRow, vendorColumnMap, vendorColumns, _
            lineSheetMap, vendorPriceCaches, seiriColumn, dayNightColumn, lineColumn, _
            qtyColumn, isWeldingSheet, matchedCount
    End If
    On Error GoTo 0

    '  (2) 金額数式
    refreshStep = "AmountFormulas"
    On Error Resume Next
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    ApplySubcontractorAmountFormulas ws, lastRow, vendorColumnMap, qtyColumn
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (3) 列内塗りつぶし(単価有無の背景色)
    refreshStep = "ColumnInteriors"
    On Error Resume Next
    RefreshSubcontractorColumnInteriors ws, lastRow, subconFirstCol, insertedColumnCount
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (3b) 施工会社列の業者情報色
    refreshStep = "VendorColumnColors"
    On Error Resume Next
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    RefreshOutputSheetVendorColumnColors ws, lastRow
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
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
            LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    '  (5) 合計行
    refreshStep = "WriteTotals"
    On Error Resume Next
    WriteOutputTotalRows ws, vendorNames, subconFirstCol, insertedColumnCount
    If Err.Number <> 0 Then
        LogCI "RefreshSubcontractorPriceColumnsCore step=" & refreshStep & " Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    '  (6) 基本情報合計(内部に独自ハンドラあり)以降は再度 RefreshError で保護
    On Error GoTo RefreshError
    refreshStep = "BasicInfoTotals"
    mod_Construction_BasicTotals.RefreshBasicInfoConstructionTotalsCore

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
    LogCI "RefreshSubcontractorPriceColumnsCore Setup Err " & refreshErrNo & ": " & _
          refreshErrDesc & " (step=" & refreshStep & ")"
    GoTo RefreshExit

RefreshError:
    refreshErrNo = Err.Number
    refreshErrDesc = Err.Description

RefreshExit:
    mod_Construction_LineMapping.RestoreWorksheetAutoFilter ws, savedAutoFilter
    Set savedAutoFilter = Nothing
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
        LogCI "RefreshSubcontractorPriceColumnsCore Err " & refreshErrNo & ": " & refreshErrDesc & " (step=" & refreshStep & ")"
    ElseIf coreErrNo <> 0 And matchedCount = 0 Then
        '  単価適用が一件も成功しなかった場合のみ通知
        MsgBox "施工会社別の単価・金額列を更新できませんでした。" & vbCrLf & _
               coreErrDesc, vbExclamation
        LogCI "RefreshSubcontractorPriceColumnsCore core Err " & coreErrNo & ": " & coreErrDesc & " (step=ApplyPrices)"
    ElseIf matchedCount = 0 Then
        If Not vendorNames Is Nothing Then
            LogCI "施工会社別単価: 一致0件 (会社数=" & vendorNames.Count & ")"
        Else
            LogCI "施工会社別単価: 一致0件"
        End If
    End If
End Sub

Public Function BuildVendorPriceColumnMap(ByVal vendorNames As Collection, _
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
        vendorKey = mod_Construction_BasicTotals.NormalizeVendorPriceName(CStr(vendorNames(vendorIndex)))
        If vendorKey <> "" Then
            If Not result.Exists(vendorKey) Then
                result.Add vendorKey, subconFirstCol + ((vendorIndex - 1) * 2)
            End If
        End If
    Next vendorIndex

    Set BuildVendorPriceColumnMap = result
End Function

Public Function CollectExistingSubconColumnVendors(ByVal ws As Worksheet, _
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

Public Function SubconColumnLayoutMatches(ByVal ws As Worksheet, _
                                           ByVal vendorNames As Collection, _
                                           ByVal subconFirstCol As Long, _
                                           ByVal kindColumn As Long) As Boolean
    Dim existingNames As Collection
    Set existingNames = CollectExistingSubconColumnVendors(ws, subconFirstCol, kindColumn)
    If existingNames.Count <> vendorNames.Count Then Exit Function

    Dim i As Long
    For i = 1 To vendorNames.Count
        If mod_Construction_BasicTotals.NormalizeVendorPriceName(CStr(existingNames(i))) <> _
           mod_Construction_BasicTotals.NormalizeVendorPriceName(CStr(vendorNames(i))) Then Exit Function
    Next i

    SubconColumnLayoutMatches = True
End Function

Public Function LookupSubcontractorVendorPrice( _
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
    If mod_Construction_BasicTotals.IsSanpaiRow(ws, rowIndex) Then Exit Function

    Dim dayNightText As String
    dayNightText = CommonNzText(ws.Cells(rowIndex, dayNightColumn).value)

    If isWeldingSheet Then
        Dim weldingPriceSheetName As String
        weldingPriceSheetName = mod_Construction_OutputLayout.ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then Exit Function
        LookupSubcontractorVendorPrice = mod_Construction_LineMapping.LookupWeldingOutputVendorPrice( _
            weldingPriceSheetName, vendorPriceCaches, _
            CommonNzText(ws.Cells(rowIndex, lineColumn).value), ws.Cells(rowIndex, seiriColumn).value, _
            CommonNzText(ws.Cells(rowIndex, vendorColumn).value), _
            (vendorColumn = WELD_COL_WELDING_VENDOR), dayNightText)
    Else
        Dim recordKey As String
        recordKey = mod_Construction_LineMapping.NormalizeRecordKey(ws.Cells(rowIndex, seiriColumn).value)
        If recordKey = "" Then Exit Function

        Dim unitPriceSheetName As String
        unitPriceSheetName = mod_Construction_LineMapping.ResolveUnitPriceSheetName( _
            lineSheetMap, CommonNzText(ws.Cells(rowIndex, lineColumn).value))
        If unitPriceSheetName = "" Then
            LogCI "施工会社単価: 線区未解決 line=[" & CommonNzText(ws.Cells(rowIndex, lineColumn).value) & "]"
            Exit Function
        End If

        Dim vendorPriceRows As Object
        Set vendorPriceRows = mod_Construction_BasicTotals.GetVendorUnitPriceRows( _
            unitPriceSheetName, CommonNzText(ws.Cells(rowIndex, vendorColumn).value), vendorPriceCaches)
        If vendorPriceRows Is Nothing Then Exit Function
        If Not vendorPriceRows.Exists(recordKey) Then
            LogCI "施工会社単価: 整理番号未一致 sheet=[" & unitPriceSheetName & "] key=[" & recordKey & "]"
            Exit Function
        End If

        Dim dayNightPrices As Variant
        dayNightPrices = vendorPriceRows(recordKey)
        LookupSubcontractorVendorPrice = mod_Construction_BasicTotals.SelectUsableDayNightPrice(dayNightText, dayNightPrices)
    End If
End Function

Public Sub ApplySubcontractorPricesPartial( _
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
                rowVendorKey = mod_Construction_BasicTotals.NormalizeVendorPriceName( _
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

Public Sub ApplySubcontractorPricesBatch( _
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
    typeColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_TYPE)
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
            If mod_Construction_BasicTotals.IsSanpaiTypeText(CommonNzText(GetArrayCellValue(typeData, r, 1))) Then GoTo NextBatchRow

            vendorColIndex = 0
            For Each vendorCol In vendorColumns
                vendorColIndex = vendorColIndex + 1
                rowVendorKey = mod_Construction_BasicTotals.NormalizeVendorPriceName( _
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

Public Sub ApplySubcontractorAmountFormulas(ByVal ws As Worksheet, _
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

Public Sub RefreshSubcontractorColumnInteriors(ByVal ws As Worksheet, _
                                                  ByVal lastRow As Long, _
                                                  ByVal firstColumn As Long, _
                                                  ByVal columnCount As Long)
    If lastRow < 2 Or columnCount <= 0 Then Exit Sub

    Dim lastColumn As Long
    Dim sanpaiFillColor As Long
    Dim r As Long
    Dim priceCol As Long
    lastColumn = firstColumn + columnCount - 1
    sanpaiFillColor = mod_Construction_BasicTotals.GetSanpaiFillColor()

    For r = 2 To lastRow
        If mod_Construction_BasicTotals.IsSanpaiRow(ws, r) Then
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

Public Sub RefreshOutputSheetVendorColumnColors(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws Is Nothing Then Exit Sub
    If lastRow < 2 Then Exit Sub
    If Not mod_Construction_OutputLayout.IsConstructionVendorOutputSheet(ws) Then Exit Sub

    Dim vendorColumns As Collection
    Set vendorColumns = mod_Construction_OutputLayout.OutputSheetVendorColumnsCore(ws)

    Dim r As Long
    Dim vendorCol As Variant
    For r = 2 To lastRow
        If mod_Construction_BasicTotals.IsSanpaiRow(ws, r) Then GoTo NextColorRow

        For Each vendorCol In vendorColumns
            mod_VendorInfoColors.ApplyOutputSheetVendorCellColor _
                ws, r, CLng(vendorCol), ResolveVendorColumnWorkTypeKeyword(ws, CLng(vendorCol))
        Next vendorCol
NextColorRow:
    Next r
End Sub

Public Function ResolveVendorColumnWorkTypeKeyword(ByVal ws As Worksheet, _
                                                    ByVal vendorColumn As Long) As String
    If Not mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then Exit Function

    If vendorColumn = WELD_COL_WELDING_VENDOR Then
        ResolveVendorColumnWorkTypeKeyword = WELDING_WORK_TYPE_KEYWORD
    ElseIf vendorColumn = WELD_COL_TRACK_VENDOR Then
        ResolveVendorColumnWorkTypeKeyword = TRACK_WORK_TYPE_KEYWORD
    End If
End Function

Public Function GetArrayCellValue(ByVal arr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If IsArray(arr) Then
        GetArrayCellValue = arr(rowIndex, colIndex)
    Else
        If rowIndex = 1 And colIndex = 1 Then GetArrayCellValue = arr
    End If
End Function

Public Function LookupSubcontractorVendorPriceFromArrays( _
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
    If mod_Construction_BasicTotals.IsSanpaiTypeText(CommonNzText(GetArrayCellValue(typeData, rowIndex, 1))) Then Exit Function

    Dim dayNightText As String
    dayNightText = CommonNzText(GetArrayCellValue(dayNightData, rowIndex, 1))

    Dim vendorName As String
    vendorName = CommonNzText(GetArrayCellValue(vendorData, rowIndex, 1))

    If isWeldingSheet Then
        Dim weldingPriceSheetName As String
        weldingPriceSheetName = mod_Construction_OutputLayout.ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then Exit Function
        LookupSubcontractorVendorPriceFromArrays = mod_Construction_LineMapping.LookupWeldingOutputVendorPrice( _
            weldingPriceSheetName, vendorPriceCaches, _
            CommonNzText(GetArrayCellValue(lineData, rowIndex, 1)), _
            GetArrayCellValue(seiriData, rowIndex, 1), vendorName, _
            (vendorColumn = WELD_COL_WELDING_VENDOR), dayNightText)
    Else
        Dim recordKey As String
        recordKey = mod_Construction_LineMapping.NormalizeRecordKey(GetArrayCellValue(seiriData, rowIndex, 1))
        If recordKey = "" Then Exit Function

        Dim unitPriceSheetName As String
        unitPriceSheetName = mod_Construction_LineMapping.ResolveUnitPriceSheetName( _
            lineSheetMap, CommonNzText(GetArrayCellValue(lineData, rowIndex, 1)))
        If unitPriceSheetName = "" Then
            LogCI "施工会社単価: 線区未解決 line=[" & CommonNzText(GetArrayCellValue(lineData, rowIndex, 1)) & "]"
            Exit Function
        End If

        Dim vendorPriceRows As Object
        Set vendorPriceRows = mod_Construction_BasicTotals.GetVendorUnitPriceRows( _
            unitPriceSheetName, vendorName, vendorPriceCaches)
        If vendorPriceRows Is Nothing Then Exit Function
        If Not vendorPriceRows.Exists(recordKey) Then Exit Function

        Dim dayNightPrices As Variant
        dayNightPrices = vendorPriceRows(recordKey)
        LookupSubcontractorVendorPriceFromArrays = mod_Construction_BasicTotals.SelectUsableDayNightPrice(dayNightText, dayNightPrices)
    End If
End Function

