Option Explicit

Private Const COLOR_INPUT_REQUIRED As Long = &HFFFF&         ' #FFFF00 (BGR)
Private Const COLOR_FILLED         As Long = &H1D1106&        ' #06111D (BGR)
Private Const COLOR_PENDING_WORK_TYPE As Long = 8628721    ' #F1A983 RGB(241,169,131)
Private Const COLOR_INPUT_REQUIRED_FONT As Long = &H0&       ' 黒
Private Const COLOR_FILLED_FONT         As Long = &HFFFFFF&  ' 白
Private Const COMMENT_FONT_NAME As String = "BIZ UDGothic"
Private Const COMMENT_FONT_SIZE As Long = 12
Private Const VENDOR_COL_START  As Long = 6
Private Const VENDOR_COL_STEP   As Long = 3
Private Const VENDOR_MAX_COUNT  As Long = 10
Private Const ROW_CONSTRUCTION_TYPE As Long = 10   ' F10/I10/L10...
Private Const ROW_RAIL_PATTERN As Long = 30        ' 軌道工事: 溶接手元単価パターン
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_CELL As String = "C24"

Private Function GetImportedLineNamesMergeArea(ByVal ws As Worksheet) As Range
    Set GetImportedLineNamesMergeArea = ws.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea
End Function

Private Function GetVendorCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    ' IsNumeric()だとF9が"4社"のような接尾辞付き文字列の場合にFalse判定となり
    ' 常に0を返してしまう(mod_VendorMaster.GetVendorBlockCountと不整合になる)ため、
    ' 同じくVal()による先頭数値抽出に統一する。
    Dim v As Variant
    v = ws.Range("F9").value
    GetVendorCount = CLng(Val(StrConv(CStr(v), vbNarrow)))
    If GetVendorCount < 0 Then GetVendorCount = 0
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
' ベンダー行の監視レンジ（全10社分、変更検知用）
' F10行も含めて工事種別変更を検知する
' ----------------------------------------------------------------

Private Function GetVendorGuideMonitorRange(ByVal ws As Worksheet) As Range
    Dim colNames As Variant
    colNames = Array("F", "I", "L", "O", "R", "U", "X", "AA", "AD", "AG")
    Dim monitorRows As Variant
    monitorRows = Array(10, 11, 27, 29, 30, 31)   ' 10行目=工事種別も監視

    Dim result As Range
    Dim cellAddr As Range
    Dim c As Long
    Dim r As Long

    On Error Resume Next
    For c = 0 To UBound(colNames)
        For r = 0 To UBound(monitorRows)
            Set cellAddr = ws.Range(CStr(colNames(c)) & CStr(monitorRows(r)))
            If Not cellAddr Is Nothing Then
                If result Is Nothing Then
                    Set result = cellAddr
                Else
                    Set result = Union(result, cellAddr)
                End If
            End If
        Next r
    Next c
    On Error GoTo 0

    Set GetVendorGuideMonitorRange = result
End Function

Private Function GuideWritableCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As Range
    Dim targetCell As Range
    Set targetCell = ws.Cells(rowNum, colNum)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
    Set GuideWritableCell = targetCell
End Function

Private Function IsConstructionTypeUndetermined(ByVal constructionType As String) As Boolean
    IsConstructionTypeUndetermined = (Len(NormalizeConstructionTypeText(constructionType)) = 0)
End Function

Private Function IsEmpty_Cell(ByVal cell As Range) As Boolean
    On Error Resume Next
    IsEmpty_Cell = (Len(Trim$(CStr(cell.value))) = 0)
    On Error GoTo 0
End Function

Private Function IsImportedLineNamesEmpty(ByVal ws As Worksheet) As Boolean
    IsImportedLineNamesEmpty = IsEmpty_Cell(ws.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.Cells(1, 1))
End Function

Private Function IsOutsourceRatioPattern(ByVal value As String) As Boolean
    IsOutsourceRatioPattern = (StrComp(NormalizeConstructionTypeText(value), mod_BasicInfoGuideTexts.GetPatternOutsourceRatioText(), vbTextCompare) = 0)
End Function

Private Function IsRailConstructionType(ByVal constructionType As String) As Boolean
    IsRailConstructionType = (StrComp(NormalizeConstructionTypeText(constructionType), mod_BasicInfoGuideTexts.GetKidoKojiText(), vbTextCompare) = 0)
End Function

Private Function IsWeldingConstructionType(ByVal constructionType As String) As Boolean
    IsWeldingConstructionType = (StrComp(NormalizeConstructionTypeText(constructionType), mod_BasicInfoGuideTexts.GetYosetsuKojiText(), vbTextCompare) = 0)
End Function

Private Function NormalizeConstructionTypeText(ByVal value As Variant) As String
    NormalizeConstructionTypeText = Trim$(CStr(value))
End Function

Private Sub ApplyDisabledCell(ByVal cell As Range)
    On Error Resume Next
    cell.Comment.Delete
    cell.Interior.Color = COLOR_FILLED
    cell.Font.Color = COLOR_FILLED_FONT

    Dim edgeId As Variant
    For Each edgeId In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With cell.Borders(edgeId)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next edgeId

    With cell.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(255, 255, 255)
    End With
    With cell.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(255, 255, 255)
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyGuideCell(ByVal ws As Worksheet, ByVal addr As String, ByVal commentText As String, ByVal isEmpty As Boolean)
    ApplyGuideCellToRange ws.Range(addr), commentText, isEmpty
End Sub

Private Sub ApplyGuideCellByRowCol(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal commentText As String, ByVal isEmpty As Boolean)
    ' 斜線が残っていれば先に解除してからガイドを適用
    RemoveDiagonalBorders ws.Cells(rowNum, colNum)
    ApplyGuideCellToRange ws.Cells(rowNum, colNum), commentText, isEmpty
End Sub

Private Sub ApplyGuideCellToRange(ByVal cell As Range, ByVal commentText As String, ByVal isEmpty As Boolean)
    On Error Resume Next
    If isEmpty Then
        cell.Interior.Color = COLOR_INPUT_REQUIRED
        cell.Font.Color = COLOR_INPUT_REQUIRED_FONT
        cell.Comment.Delete
        Dim cmt As Comment
        Set cmt = cell.AddComment(commentText)
        If Not cmt Is Nothing Then
            With cmt.Shape.TextFrame.Characters.Font
                .Color = RGB(255, 0, 0)
                .Size = COMMENT_FONT_SIZE
                .Name = COMMENT_FONT_NAME
                .Bold = True
            End With
            cmt.Visible = False
            AutoSizeComment cmt
        End If
    Else
        cell.Interior.Color = COLOR_FILLED
        cell.Font.Color = COLOR_FILLED_FONT
        cell.Comment.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyGuideMergedCell(ByVal ws As Worksheet, ByVal anchorAddr As String, ByVal commentText As String, ByVal isEmpty As Boolean)
    Dim mergeArea As Range
    Dim anchorCell As Range
    Set anchorCell = ws.Range(anchorAddr)
    Set mergeArea = anchorCell.MergeArea

    On Error Resume Next
    If isEmpty Then
        mergeArea.Interior.Pattern = xlSolid
        mergeArea.Interior.Color = COLOR_INPUT_REQUIRED
        mergeArea.Font.Color = COLOR_INPUT_REQUIRED_FONT
        anchorCell.Comment.Delete
        Dim cmt As Comment
        Set cmt = anchorCell.AddComment(commentText)
        If Not cmt Is Nothing Then
            With cmt.Shape.TextFrame.Characters.Font
                .Color = RGB(255, 0, 0)
                .Size = COMMENT_FONT_SIZE
                .Name = COMMENT_FONT_NAME
                .Bold = True
            End With
            cmt.Visible = False
            AutoSizeComment cmt
        End If
    Else
        mergeArea.Interior.Pattern = xlSolid
        mergeArea.Interior.Color = COLOR_FILLED
        mergeArea.Font.Color = COLOR_FILLED_FONT
        anchorCell.Comment.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyPendingWorkTypeFill(ByVal cell As Range)
    On Error Resume Next
    RemoveDiagonalBorders cell
    cell.ClearContents
    cell.Comment.Delete
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = COLOR_PENDING_WORK_TYPE
    cell.Font.Color = COLOR_INPUT_REQUIRED_FONT
    On Error GoTo 0
End Sub

Private Sub ApplyRailPatternValidation(ByVal targetCell As Range)
    Dim validationCell As Range
    Set validationCell = targetCell
    If validationCell.MergeCells Then Set validationCell = validationCell.MergeArea.Cells(1, 1)

    On Error Resume Next
    With validationCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:=mod_WeldingUnitPrice.BasicInfoRailPatternValidationListText()
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyRow29And31(ByVal ws As Worksheet, ByVal colNum As Long, ByVal constructionType As String)
    If IsConstructionTypeUndetermined(constructionType) Then
        ApplyGuideCellByRowCol ws, 29, colNum, mod_BasicInfoGuideTexts.GetF29CommentText(), IsEmpty_Cell(ws.Cells(29, colNum))
        ApplyRow31UndeterminedState ws, colNum
    ElseIf IsRailConstructionType(constructionType) Then
        ' 軌道工事: 29=軌道工事会社外注比率コメント / 31=行30パターン依存
        ApplyGuideCellByRowCol ws, 29, colNum, mod_BasicInfoGuideTexts.GetF29KidoCommentText(), IsEmpty_Cell(ws.Cells(29, colNum))
        ApplyRow31ForRail ws, colNum

    ElseIf IsWeldingConstructionType(constructionType) Then
        ' 溶接工事: 29=入力不可（斜線×・白） / 31左列=外注比率(%)右詰め・値列=溶接会社外注比率コメント
        ApplyDisabledCell ws.Cells(29, colNum)
        ApplyRow31ForWelding ws, colNum

    Else
        ' その他: 29=施工会社外注比率コメント / 31=現状のまま
        ApplyGuideCellByRowCol ws, 29, colNum, mod_BasicInfoGuideTexts.GetF29CommentText(), IsEmpty_Cell(ws.Cells(29, colNum))
        ApplyGuideCellByRowCol ws, 31, colNum, mod_BasicInfoGuideTexts.GetF31CommentText(), IsEmpty_Cell(ws.Cells(31, colNum))
    End If
End Sub

Private Sub ApplyRow30(ByVal ws As Worksheet, ByVal colNum As Long, ByVal constructionType As String)
    Dim labelCol As Long
    labelCol = colNum - 1

    If IsConstructionTypeUndetermined(constructionType) Then
        ApplyRow30PendingState ws, colNum
    ElseIf IsRailConstructionType(constructionType) Then
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).Value = mod_BasicInfoGuideTexts.GetRailPatternRow30LabelText()
        On Error GoTo 0
        ApplyRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
        ApplyGuideCellByRowCol ws, ROW_RAIL_PATTERN, colNum, mod_BasicInfoGuideTexts.GetF30RailPatternCommentText(), _
                               IsEmpty_Cell(GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum))
    ElseIf IsWeldingConstructionType(constructionType) Then
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).value = mod_BasicInfoGuideTexts.GetWeldingRow30LabelText()
        On Error GoTo 0
        ClearRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
        ApplyRow30WeldingValueFill GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
    Else
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).Value = mod_BasicInfoGuideTexts.GetDefaultRow30LabelText()
        GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum).ClearContents
        On Error GoTo 0
        ClearRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
        ClearGuideByRowCol ws, ROW_RAIL_PATTERN, colNum
    End If
End Sub

Private Sub ApplyRow30PendingState(ByVal ws As Worksheet, ByVal colNum As Long)
    Dim labelCol As Long
    labelCol = colNum - 1

    ClearRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
    ApplyPendingWorkTypeFill GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol)
    ApplyPendingWorkTypeFill GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
End Sub

Private Sub ApplyRow30WeldingValueFill(ByVal cell As Range)
    On Error Resume Next
    RemoveDiagonalBorders cell
    cell.Comment.Delete
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = COLOR_PENDING_WORK_TYPE
    cell.Font.Color = COLOR_INPUT_REQUIRED_FONT
    On Error GoTo 0
End Sub

Private Sub ApplyRow31ForRail(ByVal ws As Worksheet, ByVal colNum As Long)
    Dim patternValue As String
    On Error Resume Next
    patternValue = NormalizeConstructionTypeText(GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum).value)
    On Error GoTo 0

    If IsOutsourceRatioPattern(patternValue) Then
        ApplyRow31OutsourceRatioLabel ws, colNum
        ApplyGuideCellByRowCol ws, 31, colNum, mod_BasicInfoGuideTexts.GetF31KidoTemotoCommentText(), IsEmpty_Cell(ws.Cells(31, colNum))
    Else
        ClearInactiveVendorCell ws, 31, colNum - 1
        ClearInactiveVendorCell ws, 31, colNum
    End If
End Sub

' 溶接工事の行31: 左列=外注比率(%)右詰め / 値列=黄色+「溶接会社の外注比率」コメント

Private Sub ApplyRow31ForWelding(ByVal ws As Worksheet, ByVal colNum As Long)
    ApplyRow31OutsourceRatioLabel ws, colNum
    ApplyGuideCellByRowCol ws, 31, colNum, mod_BasicInfoGuideTexts.GetF31YosetsuCommentText(), IsEmpty_Cell(ws.Cells(31, colNum))
End Sub

' 行31の左列(施工会社列の1つ左)に「外注比率(%)」を右詰めで表示（#06111D・白文字）

Private Sub ApplyRow31OutsourceRatioLabel(ByVal ws As Worksheet, ByVal colNum As Long)
    On Error Resume Next
    RemoveDiagonalBorders GuideWritableCell(ws, 31, colNum - 1)
    With GuideWritableCell(ws, 31, colNum - 1)
        .Comment.Delete
        .Interior.Pattern = xlSolid
        .Interior.Color = COLOR_FILLED
        .Font.Color = COLOR_FILLED_FONT
        .value = mod_BasicInfoGuideTexts.GetOutsourceRatioLabelText()
        .HorizontalAlignment = xlRight
    End With
    On Error GoTo 0
End Sub

' 行30の値が「外注比率適用パターン」かどうか

Private Sub ApplyRow31UndeterminedState(ByVal ws As Worksheet, ByVal colNum As Long)
    Dim labelCol As Long
    labelCol = colNum - 1
    ClearInactiveVendorCell ws, 31, labelCol
    ClearInactiveVendorCell ws, 31, colNum
End Sub

' 軌道工事の行31:
'   外注比率適用パターン -> 左列=外注比率(%)右詰め / 値列=黄色+「軌道工事会社の溶接手元の外注比率」コメント
'   それ以外(前年度単価適用/物価指数適用/未選択) -> 左列・値列とも #06111D（黄色なし・コメントなし）

Private Sub AutoSizeComment(ByVal cmt As Comment)
    On Error Resume Next
    With cmt.Shape
        .TextFrame.AutoSize = True
    End With
    On Error GoTo 0
End Sub

Private Sub ClearGuideByRowCol(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    RemoveDiagonalBorders ws.Cells(rowNum, colNum)
    With ws.Cells(rowNum, colNum)
        .Interior.Color = COLOR_FILLED
        .Font.Color = COLOR_FILLED_FONT
        .Comment.Delete
    End With
    On Error GoTo 0
End Sub

Private Sub ClearInactiveVendorCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    RemoveDiagonalBorders GuideWritableCell(ws, rowNum, colNum)
    With GuideWritableCell(ws, rowNum, colNum)
        .ClearContents
        .Comment.Delete
        .Interior.Pattern = xlSolid
        .Interior.Color = COLOR_FILLED
        .Font.Color = COLOR_FILLED_FONT
    End With
    On Error GoTo 0
End Sub

Private Sub ClearRailPatternValidation(ByVal targetCell As Range)
    Dim validationCell As Range
    Set validationCell = targetCell
    If validationCell.MergeCells Then Set validationCell = validationCell.MergeArea.Cells(1, 1)

    On Error Resume Next
    validationCell.Validation.Delete
    On Error GoTo 0
End Sub

Private Sub ClearRow29And31(ByVal ws As Worksheet, ByVal colNum As Long)
    Dim labelCol As Long
    labelCol = colNum - 1
    ClearGuideByRowCol ws, 29, colNum
    ClearInactiveVendorCell ws, 31, labelCol
    ClearInactiveVendorCell ws, 31, colNum
End Sub

' ----------------------------------------------------------------
' 行30 の工事種別別処理
'   未確定 -> 左列・値列とも文字なし #F1A983 塗りのみ
'   軌道工事 -> 左列=溶接手元単価ﾊﾟﾀｰﾝ / 値列=黄色+ドロップダウン
'   溶接工事 -> 左列=溶接工事外注比率 / 値列=#F1A983
' ----------------------------------------------------------------

Private Sub ClearRow30(ByVal ws As Worksheet, ByVal colNum As Long)
    ' 会社数外(F9件数超過)の列は #F1A983 にせず、未使用ブロックと同様に暗色へ戻す
    Dim labelCol As Long
    Dim spacerCol As Long
    labelCol = colNum - 1
    spacerCol = colNum + 1

    ClearRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
    ClearInactiveVendorCell ws, ROW_RAIL_PATTERN, labelCol
    ClearInactiveVendorCell ws, ROW_RAIL_PATTERN, colNum
    ClearInactiveVendorCell ws, ROW_RAIL_PATTERN, spacerCol
End Sub

Private Sub RefreshSingleVendorRowGuide(ByVal ws As Worksheet, ByVal companyIdx As Long)
    Dim colNum As Long
    colNum = VENDOR_COL_START + (companyIdx - 1) * VENDOR_COL_STEP

    Dim constructionType As String
    constructionType = Trim$(CStr(ws.Cells(ROW_CONSTRUCTION_TYPE, colNum).value))

    ApplyGuideCellByRowCol ws, 11, colNum, mod_BasicInfoGuideTexts.GetF11CommentText(), IsEmpty_Cell(ws.Cells(11, colNum))
    ApplyGuideCellByRowCol ws, 27, colNum, mod_BasicInfoGuideTexts.GetF27CommentText(), IsEmpty_Cell(ws.Cells(27, colNum))
    ApplyRow29And31 ws, colNum, constructionType
    ApplyRow30 ws, colNum, constructionType
End Sub

' 指定した1社分だけガイドを更新（F9増加後の枠線復元後など）

Private Sub RefreshVendorRowGuides(ByVal ws As Worksheet)
    Dim vendorCount As Long
    vendorCount = GetVendorCount(ws)
    If vendorCount > VENDOR_MAX_COUNT Then vendorCount = VENDOR_MAX_COUNT

    Dim companyIdx As Long
    For companyIdx = 1 To VENDOR_MAX_COUNT
        Dim colNum As Long
        colNum = VENDOR_COL_START + (companyIdx - 1) * VENDOR_COL_STEP

        If companyIdx <= vendorCount Then
            RefreshSingleVendorRowGuide ws, companyIdx
        Else
            ' 会社数外 → 全行を元の色に戻す（斜線も解除）
            ClearGuideByRowCol ws, 11, colNum
            ClearGuideByRowCol ws, 27, colNum
            ClearRow29And31 ws, colNum
            ClearRow30 ws, colNum
        End If
    Next companyIdx
End Sub

Private Sub RemoveDiagonalBorders(ByVal cell As Range)
    On Error Resume Next
    cell.Borders(xlDiagonalDown).LineStyle = xlNone
    cell.Borders(xlDiagonalUp).LineStyle = xlNone
    On Error GoTo 0
End Sub

Public Function GetVendorGuideMonitorRangePublic(ByVal ws As Worksheet) As Range
    Set GetVendorGuideMonitorRangePublic = GetVendorGuideMonitorRange(ws)
End Function

' ----------------------------------------------------------------
' 入力不可セル: 斜線×（右下がり＋左下がり）+ #06111D
' 斜線色は白（#FFFFFF）
' ----------------------------------------------------------------

Public Sub ClearAllGuides(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False

    Dim companyIdx As Long
    For companyIdx = 1 To VENDOR_MAX_COUNT
        Dim colNum As Long
        colNum = VENDOR_COL_START + (companyIdx - 1) * VENDOR_COL_STEP
        RemoveDiagonalBorders ws.Cells(29, colNum)
        RemoveDiagonalBorders ws.Cells(31, colNum)
        ClearRailPatternValidation ws.Cells(ROW_RAIL_PATTERN, colNum)
    Next companyIdx

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub InitBasicInfoGuide(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False

    ' 固定セル
    ApplyGuideCell ws, "C9",  mod_BasicInfoGuideTexts.GetC9CommentText(),  IsEmpty_Cell(ws.Range("C9"))
    ApplyGuideMergedCell ws, "C13", mod_BasicInfoGuideTexts.GetC13CommentText(), IsEmpty_Cell(ws.Range("C13").MergeArea.Cells(1, 1))
    ApplyGuideMergedCell ws, "C22", mod_BasicInfoGuideTexts.GetC22CommentText(), IsEmpty_Cell(ws.Range("C22").MergeArea.Cells(1, 1))
    ApplyGuideMergedCell ws, "C23", mod_BasicInfoGuideTexts.GetC23CommentText(), IsEmpty_Cell(ws.Range("C23").MergeArea.Cells(1, 1))
    ApplyGuideMergedCell ws, BASIC_INFO_IMPORTED_LINE_NAMES_CELL, mod_BasicInfoGuideTexts.GetC24CommentText(), IsImportedLineNamesEmpty(ws)

    ' F9（会社数）
    ApplyGuideCell ws, "F9", mod_BasicInfoGuideTexts.GetF9CommentText(), IsEmpty_Cell(ws.Range("F9"))

    ' ベンダー行（会社数・工事種別に応じて列数を決定）
    RefreshVendorRowGuides ws

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub OnCellChanged(ByVal ws As Worksheet, ByVal target As Range)
    If ws Is Nothing Or target Is Nothing Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False

    ' 固定セル判定
    If Not Intersect(target, ws.Range("C9"))  Is Nothing Then ApplyGuideCell ws, "C9",  mod_BasicInfoGuideTexts.GetC9CommentText(),  IsEmpty_Cell(ws.Range("C9"))
    If Not Intersect(target, ws.Range("C13").MergeArea) Is Nothing Then ApplyGuideMergedCell ws, "C13", mod_BasicInfoGuideTexts.GetC13CommentText(), IsEmpty_Cell(ws.Range("C13").MergeArea.Cells(1, 1))
    If Not Intersect(target, ws.Range("C22")) Is Nothing Then ApplyGuideMergedCell ws, "C22", mod_BasicInfoGuideTexts.GetC22CommentText(), IsEmpty_Cell(ws.Range("C22").MergeArea.Cells(1, 1))
    If Not Intersect(target, ws.Range("C23")) Is Nothing Then ApplyGuideMergedCell ws, "C23", mod_BasicInfoGuideTexts.GetC23CommentText(), IsEmpty_Cell(ws.Range("C23").MergeArea.Cells(1, 1))
    If Not Intersect(target, GetImportedLineNamesMergeArea(ws)) Is Nothing Then _
        ApplyGuideMergedCell ws, BASIC_INFO_IMPORTED_LINE_NAMES_CELL, mod_BasicInfoGuideTexts.GetC24CommentText(), IsImportedLineNamesEmpty(ws)
    If Not Intersect(target, ws.Range("F9"))  Is Nothing Then
        ApplyGuideCell ws, "F9", mod_BasicInfoGuideTexts.GetF9CommentText(), IsEmpty_Cell(ws.Range("F9"))
        RefreshVendorRowGuides ws
    End If

    ' ベンダー行セル（全10社分の列×4行 + F10等の工事種別行を監視）
    Dim vendorMonitorRange As Range
    Set vendorMonitorRange = GetVendorGuideMonitorRange(ws)
    If Not vendorMonitorRange Is Nothing Then
        If Not Intersect(target, vendorMonitorRange) Is Nothing Then
            RefreshVendorRowGuides ws
        End If
    End If

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub RefreshSingleVendorRowGuidePublic(ByVal ws As Worksheet, ByVal companyIdx As Long)
    If ws Is Nothing Then Exit Sub
    If companyIdx < 1 Or companyIdx > VENDOR_MAX_COUNT Then Exit Sub
    On Error Resume Next
    RefreshSingleVendorRowGuide ws, companyIdx
    On Error GoTo 0
End Sub

Public Sub RefreshVendorGuidesForBasicInfo(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshVendorRowGuides ws
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

