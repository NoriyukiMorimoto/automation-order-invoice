Option Explicit

' ========================================================
' mod_BasicInfoGuide
' 基本情報シートの入力ガイド（塗色 & コメント）管理
'
' ルール:
'   未入力 -> #FFFF00 塗色 + 黒文字 + 赤字12pt BizUDゴシックのコメント
'   入力済 -> #06111D 塗色 + 白文字 + コメント削除
'
' F10（工事種別）による F29/F30/F31 の動的切り替え:
'   未確定   -> F30左列・値列とも文字なし #F1A983 塗りのみ
'   軌道工事 -> F29: 軌道工事会社の外注比率コメント / F30: 溶接手元単価ﾊﾟﾀｰﾝ / F31: 現状のまま
'   溶接工事 -> F29: 入力不可（斜線×） / F30左列: 溶接工事外注比率 / F31: 溶接会社の外注比率コメント
'   その他   -> F29: 施工会社の外注比率コメント / F31: 現状のまま
'
' 会社数セル(F9)の値に応じてF/I/L/O/R/U/X/AA/AD/AG列のベンダー行を動的管理
' 最大10社対応（3列おき、F=6列目始まり）
' ========================================================

' --- 背景色定数 ---
Private Const COLOR_INPUT_REQUIRED As Long = &HFFFF&         ' #FFFF00 (BGR)
Private Const COLOR_FILLED         As Long = &H1D1106&        ' #06111D (BGR)
Private Const COLOR_PENDING_WORK_TYPE As Long = 8628721    ' #F1A983 RGB(241,169,131)
Private Const COLOR_INPUT_REQUIRED_FONT As Long = &H0&       ' 黒
Private Const COLOR_FILLED_FONT         As Long = &HFFFFFF&  ' 白

' --- コメントフォント設定 ---
Private Const COMMENT_FONT_NAME As String = "BIZ UDGothic"
Private Const COMMENT_FONT_SIZE As Long = 12

' --- ベンダー列: F=6始まり、3列おき、最大10社 ---
Private Const VENDOR_COL_START  As Long = 6
Private Const VENDOR_COL_STEP   As Long = 3
Private Const VENDOR_MAX_COUNT  As Long = 10

' --- 工事種別行 ---
Private Const ROW_CONSTRUCTION_TYPE As Long = 10   ' F10/I10/L10...
Private Const ROW_RAIL_PATTERN As Long = 30        ' 軌道工事: 溶接手元単価パターン

' --- 線区名（結合セル C24:C28、先頭セル C24 で MergeArea を取得）---
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_CELL As String = "C24"

' ----------------------------------------------------------------
' 公開API: シート全体のガイド状態を初期化（Activate時などに呼ぶ）
' ----------------------------------------------------------------
Public Sub InitBasicInfoGuide(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False

    ' 固定セル
    ApplyGuideCell ws, "C9",  GetC9CommentText(),  IsEmpty_Cell(ws.Range("C9"))
    ApplyGuideCell ws, "C22", GetC22CommentText(), IsEmpty_Cell(ws.Range("C22"))
    ApplyGuideCell ws, "C23", GetC23CommentText(), IsEmpty_Cell(ws.Range("C23"))
    ApplyGuideMergedCell ws, BASIC_INFO_IMPORTED_LINE_NAMES_CELL, GetC24CommentText(), IsImportedLineNamesEmpty(ws)

    ' F9（会社数）
    ApplyGuideCell ws, "F9", GetF9CommentText(), IsEmpty_Cell(ws.Range("F9"))

    ' ベンダー行（会社数・工事種別に応じて列数を決定）
    RefreshVendorRowGuides ws

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' 公開API: 特定セル変更時に呼ぶ（Worksheet_Changeから）
' ----------------------------------------------------------------
Public Sub OnCellChanged(ByVal ws As Worksheet, ByVal target As Range)
    If ws Is Nothing Or target Is Nothing Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False

    ' 固定セル判定
    If Not Intersect(target, ws.Range("C9"))  Is Nothing Then ApplyGuideCell ws, "C9",  GetC9CommentText(),  IsEmpty_Cell(ws.Range("C9"))
    If Not Intersect(target, ws.Range("C22")) Is Nothing Then ApplyGuideCell ws, "C22", GetC22CommentText(), IsEmpty_Cell(ws.Range("C22"))
    If Not Intersect(target, ws.Range("C23")) Is Nothing Then ApplyGuideCell ws, "C23", GetC23CommentText(), IsEmpty_Cell(ws.Range("C23"))
    If Not Intersect(target, GetImportedLineNamesMergeArea(ws)) Is Nothing Then _
        ApplyGuideMergedCell ws, BASIC_INFO_IMPORTED_LINE_NAMES_CELL, GetC24CommentText(), IsImportedLineNamesEmpty(ws)
    If Not Intersect(target, ws.Range("F9"))  Is Nothing Then
        ApplyGuideCell ws, "F9", GetF9CommentText(), IsEmpty_Cell(ws.Range("F9"))
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

' ----------------------------------------------------------------
' 公開API: ClearBasicInfo時に全斜線・ガイドをリセット
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

' ----------------------------------------------------------------
' ベンダー行ガイドを会社数・工事種別に応じてリフレッシュ（最大10社）
' ----------------------------------------------------------------
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

Private Sub RefreshSingleVendorRowGuide(ByVal ws As Worksheet, ByVal companyIdx As Long)
    Dim colNum As Long
    colNum = VENDOR_COL_START + (companyIdx - 1) * VENDOR_COL_STEP

    Dim constructionType As String
    constructionType = Trim$(CStr(ws.Cells(ROW_CONSTRUCTION_TYPE, colNum).value))

    ApplyGuideCellByRowCol ws, 11, colNum, GetF11CommentText(), IsEmpty_Cell(ws.Cells(11, colNum))
    ApplyGuideCellByRowCol ws, 27, colNum, GetF27CommentText(), IsEmpty_Cell(ws.Cells(27, colNum))
    ApplyRow29And31 ws, colNum, constructionType
    ApplyRow30 ws, colNum, constructionType
End Sub

' 指定した1社分だけガイドを更新（F9増加後の枠線復元後など）
Public Sub RefreshSingleVendorRowGuidePublic(ByVal ws As Worksheet, ByVal companyIdx As Long)
    If ws Is Nothing Then Exit Sub
    If companyIdx < 1 Or companyIdx > VENDOR_MAX_COUNT Then Exit Sub
    On Error Resume Next
    RefreshSingleVendorRowGuide ws, companyIdx
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' 行29・31 の工事種別別処理
' ----------------------------------------------------------------
Private Function NormalizeConstructionTypeText(ByVal value As Variant) As String
    NormalizeConstructionTypeText = Trim$(CStr(value))
End Function

Private Function IsRailConstructionType(ByVal constructionType As String) As Boolean
    IsRailConstructionType = (StrComp(NormalizeConstructionTypeText(constructionType), GetKidoKojiText(), vbTextCompare) = 0)
End Function

Private Function IsWeldingConstructionType(ByVal constructionType As String) As Boolean
    IsWeldingConstructionType = (StrComp(NormalizeConstructionTypeText(constructionType), GetYosetsuKojiText(), vbTextCompare) = 0)
End Function

Private Function IsConstructionTypeUndetermined(ByVal constructionType As String) As Boolean
    IsConstructionTypeUndetermined = (Len(NormalizeConstructionTypeText(constructionType)) = 0)
End Function

Private Function GuideWritableCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As Range
    Dim targetCell As Range
    Set targetCell = ws.Cells(rowNum, colNum)
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
    Set GuideWritableCell = targetCell
End Function

Private Sub ApplyRow29And31(ByVal ws As Worksheet, ByVal colNum As Long, ByVal constructionType As String)
    If IsRailConstructionType(constructionType) Then
        ' 軌道工事: 29=軌道工事会社外注比率コメント / 31=現状のまま
        ApplyGuideCellByRowCol ws, 29, colNum, GetF29KidoCommentText(), IsEmpty_Cell(ws.Cells(29, colNum))
        ApplyGuideCellByRowCol ws, 31, colNum, GetF31CommentText(), IsEmpty_Cell(ws.Cells(31, colNum))

    ElseIf IsWeldingConstructionType(constructionType) Then
        ' 溶接工事: 29=入力不可（斜線×） / 31=溶接会社外注比率コメント
        ApplyDisabledCell ws.Cells(29, colNum)
        ApplyGuideCellByRowCol ws, 31, colNum, GetF31YosetsuCommentText(), IsEmpty_Cell(ws.Cells(31, colNum))

    Else
        ' その他: 29=施工会社外注比率コメント / 31=現状のまま
        ApplyGuideCellByRowCol ws, 29, colNum, GetF29CommentText(), IsEmpty_Cell(ws.Cells(29, colNum))
        ApplyGuideCellByRowCol ws, 31, colNum, GetF31CommentText(), IsEmpty_Cell(ws.Cells(31, colNum))
    End If
End Sub

' ----------------------------------------------------------------
' 行29・31 を元の色に戻す（会社数外・クリア時）
' ----------------------------------------------------------------
Private Sub ClearRow29And31(ByVal ws As Worksheet, ByVal colNum As Long)
    ClearGuideByRowCol ws, 29, colNum
    ClearGuideByRowCol ws, 31, colNum
End Sub

' ----------------------------------------------------------------
' 行30 の工事種別別処理
'   未確定 -> 左列・値列とも文字なし #F1A983 塗りのみ
'   軌道工事 -> 左列=溶接手元単価ﾊﾟﾀｰﾝ / 値列=黄色+ドロップダウン
'   溶接工事 -> 左列=溶接工事外注比率 / 値列=従来どおり
' ----------------------------------------------------------------
Private Sub ApplyRow30(ByVal ws As Worksheet, ByVal colNum As Long, ByVal constructionType As String)
    Dim labelCol As Long
    labelCol = colNum - 1

    If IsConstructionTypeUndetermined(constructionType) Then
        ApplyRow30PendingState ws, colNum
    ElseIf IsRailConstructionType(constructionType) Then
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).Value = GetRailPatternRow30LabelText()
        On Error GoTo 0
        ApplyRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
        ApplyGuideCellByRowCol ws, ROW_RAIL_PATTERN, colNum, GetF30RailPatternCommentText(), _
                               IsEmpty_Cell(GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum))
    ElseIf IsWeldingConstructionType(constructionType) Then
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).Value = GetWeldingRow30LabelText()
        On Error GoTo 0
        ClearRailPatternValidation GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum)
        ApplyGuideCellByRowCol ws, ROW_RAIL_PATTERN, colNum, GetF31YosetsuCommentText(), _
                               IsEmpty_Cell(GuideWritableCell(ws, ROW_RAIL_PATTERN, colNum))
    Else
        On Error Resume Next
        GuideWritableCell(ws, ROW_RAIL_PATTERN, labelCol).Value = GetDefaultRow30LabelText()
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

Private Sub ClearRailPatternValidation(ByVal targetCell As Range)
    Dim validationCell As Range
    Set validationCell = targetCell
    If validationCell.MergeCells Then Set validationCell = validationCell.MergeArea.Cells(1, 1)

    On Error Resume Next
    validationCell.Validation.Delete
    On Error GoTo 0
End Sub

Private Function VendorRowLabelPrefixText() As String
    VendorRowLabelPrefixText = " " & ChrW$(&H25B8) & " "
End Function

Private Function GetDefaultRow30LabelText() As String
    GetDefaultRow30LabelText = GetWeldingRow30LabelText()
End Function

Private Function GetWeldingRow30LabelText() As String
    GetWeldingRow30LabelText = VendorRowLabelPrefixText() & _
        ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & _
        ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387)
End Function

Private Function GetRailPatternRow30LabelText() As String
    GetRailPatternRow30LabelText = VendorRowLabelPrefixText() & _
        ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & ChrW$(&H5143) & _
        ChrW$(&H5358) & ChrW$(&H4FA1) & _
        ChrW$(&HFF8A) & ChrW$(&HFF9F) & ChrW$(&HFF80) & ChrW$(&HFF70) & ChrW$(&HFF9D)
End Function

' F9(施工会社数)変更後など、ベンダー行ガイドを一括更新する。
Public Sub RefreshVendorGuidesForBasicInfo(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    RefreshVendorRowGuides ws
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' 公開API: ベンダー行の監視レンジ（Sheet1 Worksheet_Change 用）
' ----------------------------------------------------------------
Public Function GetVendorGuideMonitorRangePublic(ByVal ws As Worksheet) As Range
    Set GetVendorGuideMonitorRangePublic = GetVendorGuideMonitorRange(ws)
End Function

' ----------------------------------------------------------------
' 入力不可セル: 斜線×（右下がり＋左下がり）+ #06111D
' ----------------------------------------------------------------
Private Sub ApplyDisabledCell(ByVal cell As Range)
    On Error Resume Next
    cell.Comment.Delete
    cell.Interior.Color = COLOR_FILLED
    cell.Font.Color = COLOR_FILLED_FONT
    With cell.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(255, 0, 0)
    End With
    With cell.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(255, 0, 0)
    End With
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' 斜線を解除する
' ----------------------------------------------------------------
Private Sub RemoveDiagonalBorders(ByVal cell As Range)
    On Error Resume Next
    cell.Borders(xlDiagonalDown).LineStyle = xlNone
    cell.Borders(xlDiagonalUp).LineStyle = xlNone
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' 単一セルにガイドを適用（アドレス指定）
' ----------------------------------------------------------------
Private Sub ApplyGuideCell(ByVal ws As Worksheet, ByVal addr As String, ByVal commentText As String, ByVal isEmpty As Boolean)
    ApplyGuideCellToRange ws.Range(addr), commentText, isEmpty
End Sub

' ----------------------------------------------------------------
' 結合セルにガイドを適用（先頭セルから MergeArea を取得、塗色は結合範囲全体）
' ----------------------------------------------------------------
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

Private Function GetImportedLineNamesMergeArea(ByVal ws As Worksheet) As Range
    Set GetImportedLineNamesMergeArea = ws.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea
End Function

' ----------------------------------------------------------------
' 単一セルにガイドを適用（行列番号指定）
' ----------------------------------------------------------------
Private Sub ApplyGuideCellByRowCol(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal commentText As String, ByVal isEmpty As Boolean)
    ' 斜線が残っていれば先に解除してからガイドを適用
    RemoveDiagonalBorders ws.Cells(rowNum, colNum)
    ApplyGuideCellToRange ws.Cells(rowNum, colNum), commentText, isEmpty
End Sub

' ----------------------------------------------------------------
' ガイドを解除して元の色に戻す（斜線も解除）
' ----------------------------------------------------------------
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

' ----------------------------------------------------------------
' コアロジック: 塗色 + コメント付け外し
' ----------------------------------------------------------------
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

' ----------------------------------------------------------------
' コメントサイズ自動調整
' ----------------------------------------------------------------
Private Sub AutoSizeComment(ByVal cmt As Comment)
    On Error Resume Next
    With cmt.Shape
        .TextFrame.AutoSize = True
    End With
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' セルが空かどうか（Trim後で判定）
' ----------------------------------------------------------------
Private Function IsEmpty_Cell(ByVal cell As Range) As Boolean
    On Error Resume Next
    IsEmpty_Cell = (Len(Trim$(CStr(cell.value))) = 0)
    On Error GoTo 0
End Function

Private Function IsImportedLineNamesEmpty(ByVal ws As Worksheet) As Boolean
    IsImportedLineNamesEmpty = IsEmpty_Cell(ws.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.Cells(1, 1))
End Function

' ----------------------------------------------------------------
' 会社数取得（F9セル）
' ----------------------------------------------------------------
Private Function GetVendorCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    Dim v As Variant
    v = ws.Range("F9").value
    If IsNumeric(v) Then
        GetVendorCount = CLng(v)
    Else
        GetVendorCount = 0
    End If
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

' ================================================================
' コメントテキスト（ChrW$ で日本語構築）
' ================================================================

Private Function GetC9CommentText() As String
    ' 工事番号が工事現況表に登録されている場合はダブルクリックしてください。
    GetC9CommentText = _
        ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H756A) & ChrW$(&H53F7) & ChrW$(&H304C) & _
        ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H73FE) & ChrW$(&H6CC1) & ChrW$(&H8868) & _
        ChrW$(&H306B) & ChrW$(&H767B) & ChrW$(&H9332) & ChrW$(&H3055) & ChrW$(&H308C) & _
        ChrW$(&H3066) & ChrW$(&H3044) & ChrW$(&H308B) & ChrW$(&H5834) & ChrW$(&H5408) & _
        ChrW$(&H306F) & ChrW$(&H30C0) & ChrW$(&H30D6) & ChrW$(&H30EB) & ChrW$(&H30AF) & _
        ChrW$(&H30EA) & ChrW$(&H30C3) & ChrW$(&H30AF) & ChrW$(&H3057) & ChrW$(&H3066) & _
        ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetC22CommentText() As String
    ' 本工事に適用される単価を選択して下さい。
    GetC22CommentText = _
        ChrW$(&H672C) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H306B) & ChrW$(&H9069) & _
        ChrW$(&H7528) & ChrW$(&H3055) & ChrW$(&H308C) & ChrW$(&H308B) & ChrW$(&H5358) & _
        ChrW$(&H4FA1) & ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
        ChrW$(&H3066) & ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetC23CommentText() As String
    ' 本工事で溶接工事の有無を選択して下さい。
    GetC23CommentText = _
        ChrW$(&H672C) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H3067) & ChrW$(&H6EB6) & _
        ChrW$(&H63A5) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H306E) & ChrW$(&H6709) & _
        ChrW$(&H7121) & ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
        ChrW$(&H3066) & ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetC24CommentText() As String
    ' 本工事に適用される線区名をダブルクリックして選択してください。
    GetC24CommentText = _
        ChrW$(&H672C) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H306B) & ChrW$(&H9069) & _
        ChrW$(&H7528) & ChrW$(&H3055) & ChrW$(&H308C) & ChrW$(&H308B) & ChrW$(&H7DDA) & _
        ChrW$(&H533A) & ChrW$(&H540D) & ChrW$(&H3092) & ChrW$(&H30C0) & ChrW$(&H30D6) & _
        ChrW$(&H30EB) & ChrW$(&H30AF) & ChrW$(&H30EA) & ChrW$(&H30C3) & ChrW$(&H30AF) & _
        ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
        ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & _
        ChrW$(&H3002)
End Function

Private Function GetF9CommentText() As String
    ' 本工事を施工する施工会社数を入力して下さい。
    GetF9CommentText = _
        ChrW$(&H672C) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H3092) & ChrW$(&H65BD) & _
        ChrW$(&H5DE5) & ChrW$(&H3059) & ChrW$(&H308B) & ChrW$(&H65BD) & ChrW$(&H5DE5) & _
        ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H6570) & ChrW$(&H3092) & ChrW$(&H5165) & _
        ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H4E0B) & ChrW$(&H3055) & _
        ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF11CommentText() As String
    ' 施工会社名を選択して下さい。
    GetF11CommentText = _
        ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H540D) & _
        ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & ChrW$(&H3066) & _
        ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF27CommentText() As String
    ' 注文番号を入力してください。
    GetF27CommentText = _
        ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H756A) & ChrW$(&H53F7) & ChrW$(&H3092) & _
        ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & _
        ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF29CommentText() As String
    ' 施工会社の外注比率を入力して下さい。
    GetF29CommentText = _
        ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H306E) & _
        ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&H3092) & _
        ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H4E0B) & _
        ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF29KidoCommentText() As String
    ' 軌道工事会社の外注比率を入力して下さい。
    GetF29KidoCommentText = _
        ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H4F1A) & _
        ChrW$(&H793E) & ChrW$(&H306E) & ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & _
        ChrW$(&H7387) & ChrW$(&H3092) & ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & _
        ChrW$(&H3066) & ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF31CommentText() As String
    ' 軌道会社の溶接手元を外注比率で支払う場合は外注比率を入力して下さい。
    GetF31CommentText = _
        ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H306E) & _
        ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & ChrW$(&H5143) & ChrW$(&H3092) & _
        ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&H3067) & _
        ChrW$(&H652F) & ChrW$(&H6255) & ChrW$(&H3046) & ChrW$(&H5834) & ChrW$(&H5408) & _
        ChrW$(&H306F) & ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & _
        ChrW$(&H3092) & ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & _
        ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF31YosetsuCommentText() As String
    ' 溶接会社の外注比率を入力して下さい。
    GetF31YosetsuCommentText = _
        ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H4F1A) & ChrW$(&H793E) & ChrW$(&H306E) & _
        ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&H3092) & _
        ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H4E0B) & _
        ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

Private Function GetF30RailPatternCommentText() As String
    ' 溶接手元単価の算出パターンを選択して下さい。
    GetF30RailPatternCommentText = _
        ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & ChrW$(&H5143) & ChrW$(&H5358) & ChrW$(&H4FA1) & _
        ChrW$(&H306E) & ChrW$(&H7B97) & ChrW$(&H51FA) & ChrW$(&H30D1) & ChrW$(&H30BF) & ChrW$(&H30FC) & _
        ChrW$(&H30F3) & ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & ChrW$(&H3066) & _
        ChrW$(&H4E0B) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

' ================================================================
' 工事種別定数文字列（ChrW$ で構築）
' ================================================================

Private Function GetKidoKojiText() As String
    ' 軌道工事
    GetKidoKojiText = ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
End Function

Private Function GetYosetsuKojiText() As String
    ' 溶接工事
    GetYosetsuKojiText = ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
End Function
