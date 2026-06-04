Option Explicit

'==========================================================================
'  基本情報シート 日付ダブルクリック用カレンダー＆工事番号選択ディスパッチ
'    改修内容：
'      #9  : ExtractYear4Digits / 基本情報シート取得を mod_Common 経由に統一。
'            候補シート名配列の重複定義を削減。
'      #10 : DrawBasicInfoCalendar / DeleteBasicInfoCalendar の前後で
'            ScreenUpdating を OFF にし、Shape の追加・削除でちらつきが
'            出ないようにする（描画のもたつき緩和）。
'      #11 : C9 / 日付セルの対象判定で Target.MergeArea が 1004 を出す
'            場合に備え、MergeArea 左上セル取得を安全化。
'==========================================================================

Private Const CALENDAR_PREFIX As String = "BICal_"
Private Const LEGACY_CALENDAR_PREFIX As String = "__BasicInfoCalendar_"
Private Const CALENDAR_TARGET_SHEET_NAME As String = "__BasicInfoCalendarTargetSheet"
Private Const CALENDAR_TARGET_ADDRESS_NAME As String = "__BasicInfoCalendarTargetAddress"
Private Const CALENDAR_MONTH_NAME As String = "__BasicInfoCalendarMonth"
Private Const BASIC_INFO_DATE_CELLS As String = "C11,C15,C16,C2,F2,F3"

Public SharedMasterData As Variant
Public ProjectSelectionBasicInfoMode As Boolean
Public ProjectSelectionTargetSheetName As String
Public ProjectSelectionTargetAddress As String
Public ProjectSelectionYear As Long
Public ProjectSelectionBranchName As String
Public ProjectSelectionOfficeName As String

Private mCalendarTargetSheetName As String
Private mCalendarTargetAddress As String
Private mCalendarDisplayMonth As Date

Public Sub HandleBasicInfoDateDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> GetBasicInfoWorksheetName() Then Exit Sub

    Dim targetCell As Range
    Set targetCell = GetDateTargetCell(Target)
    If targetCell Is Nothing Then Exit Sub

    Cancel = True
    ShowBasicInfoDateCalendar targetCell
End Sub

Public Sub HandleBasicInfoProjectNameDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> GetBasicInfoWorksheetName() Then Exit Sub

    Dim targetCell As Range
    Set targetCell = GetBasicInfoProjectNameTargetCell(Target)
    If targetCell Is Nothing Then Exit Sub

    Cancel = True
    ShowBasicInfoProjectNameSelection targetCell.Worksheet, targetCell
End Sub

Private Sub ShowBasicInfoProjectNameSelection(ByVal wsInfo As Worksheet, ByVal targetCell As Range)
    ProjectSelectionBasicInfoMode = True
    ProjectSelectionTargetSheetName = wsInfo.Name
    ProjectSelectionTargetAddress = targetCell.Address(False, False)
    ProjectSelectionYear = Val(CommonExtractYear4Digits(CStr(wsInfo.Range("B4").value)))
    ProjectSelectionBranchName = Trim$(CStr(wsInfo.Range("B6").value))
    ProjectSelectionOfficeName = NormalizeProjectSelectionOfficeName(Trim$(CStr(wsInfo.Range("C6").value)))

    On Error Resume Next
    Project_Number_Selection.ClearSharedMasterData
    On Error GoTo 0

    Project_Number_Selection.Show vbModal
    ClearProjectSelectionState
End Sub

Public Sub ClearProjectSelectionState()
    ProjectSelectionBasicInfoMode = False
    ProjectSelectionTargetSheetName = ""
    ProjectSelectionTargetAddress = ""
    ProjectSelectionYear = 0
    ProjectSelectionBranchName = ""
    ProjectSelectionOfficeName = ""
End Sub

Public Sub HandleBasicInfoSelectionChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetStoredCalendarWorksheet()
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub

    Dim targetAddress As String
    targetAddress = GetCalendarTargetAddress()
    If targetAddress = "" Then Exit Sub

    Dim calendarTarget As Range
    Set calendarTarget = ws.Range(targetAddress)
    If Intersect(Target, calendarTarget.MergeArea) Is Nothing Then
        DeleteBasicInfoCalendar ws
    End If
End Sub

Public Sub ShowBasicInfoDateCalendar(ByVal targetCell As Range)
    Dim baseDate As Date
    baseDate = GetInitialCalendarDate(targetCell)

    StoreCalendarState targetCell, DateSerial(Year(baseDate), Month(baseDate), 1)
    DrawBasicInfoCalendar targetCell.Worksheet, targetCell, DateSerial(Year(baseDate), Month(baseDate), 1), baseDate
End Sub

Public Sub BasicInfoCalendarSelectDay()
    Dim callerName As String

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetStoredCalendarWorksheet()
    If ws Is Nothing Then Exit Sub

    callerName = CStr(Application.Caller)

    Dim selectedDate As Date
    If TryGetDateFromCallerName(callerName, selectedDate) Then
        ' shape 名に埋め込んだ日付を採用
    ElseIf TryGetDateFromSelection(selectedDate) Then
        ' 選択中の shape から日付取得
    Else
        selectedDate = DateFromIsoText(ws.Shapes(callerName).AlternativeText)
    End If

    Dim targetAddress As String
    targetAddress = GetCalendarTargetAddress()
    If targetAddress = "" Then Err.Raise vbObjectError + 514, , "Calendar target cell is not stored."

    Dim targetCell As Range
    Set targetCell = ws.Range(targetAddress)

    With targetCell
        .value = selectedDate
        .NumberFormatLocal = "yyyy年m月d日"
    End With

    DeleteBasicInfoCalendar ws
    targetCell.Select

    Exit Sub

ErrorHandler:
    MsgBox "日付を入力できませんでした。" & vbCrLf & _
           "Caller: " & callerName & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Public Sub BasicInfoCalendarPreviousMonth()
    MoveBasicInfoCalendarMonth -1
End Sub

Public Sub BasicInfoCalendarNextMonth()
    MoveBasicInfoCalendarMonth 1
End Sub

Public Sub BasicInfoCalendarClose()
    Dim ws As Worksheet
    Set ws = GetStoredCalendarWorksheet()
    If Not ws Is Nothing Then DeleteBasicInfoCalendar ws
End Sub

Private Sub MoveBasicInfoCalendarMonth(ByVal monthOffset As Long)
    On Error GoTo ExitHandler

    Dim ws As Worksheet
    Set ws = GetStoredCalendarWorksheet()
    If ws Is Nothing Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Range(GetCalendarTargetAddress())

    Dim currentMonth As Date
    currentMonth = GetCalendarDisplayMonth()
    currentMonth = DateAdd("m", monthOffset, currentMonth)

    StoreCalendarState targetCell, currentMonth
    DrawBasicInfoCalendar ws, targetCell, currentMonth, GetInitialCalendarDate(targetCell)

ExitHandler:
End Sub

'--------------------------------------------------------------------------
'  カレンダー本体描画
'  改修（#10）：Shape の生成・削除でちらつきが出るため ScreenUpdating
'                を OFF にして一括描画。元から呼び出される DeleteBasicInfoCalendar
'                とのネスト整合のため、状態を保存・復元する。
'--------------------------------------------------------------------------
Private Sub DrawBasicInfoCalendar(ByVal ws As Worksheet, ByVal targetCell As Range, ByVal displayMonth As Date, ByVal baseDate As Date)
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    DeleteBasicInfoCalendar ws

    Dim anchor As Range
    Set anchor = targetCell.MergeArea

    Dim leftPos As Double, topPos As Double
    leftPos = anchor.Left
    topPos = anchor.Top + anchor.Height + 4

    Dim cellW As Double, cellH As Double
    cellW = 32
    cellH = 24

    Dim calendarW As Double, calendarH As Double, footerTop As Double
    calendarW = cellW * 7 + 16
    footerTop = 64 + cellH * 6 + 6
    calendarH = footerTop + 30

    AddCalendarShape ws, "Back", msoShapeRectangle, leftPos, topPos, calendarW, calendarH, "", RGB(248, 250, 252), RGB(15, 23, 42), 1

    AddCalendarShape ws, "Prev", msoShapeRectangle, leftPos + 8, topPos + 8, 28, 24, "<", RGB(226, 232, 240), RGB(100, 116, 139), 1, "'" & ThisWorkbook.Name & "'!BasicInfoCalendarPreviousMonth"
    AddCalendarShape ws, "Next", msoShapeRectangle, leftPos + calendarW - 36, topPos + 8, 28, 24, ">", RGB(226, 232, 240), RGB(100, 116, 139), 1, "'" & ThisWorkbook.Name & "'!BasicInfoCalendarNextMonth"
    AddCalendarShape ws, "Close", msoShapeRectangle, leftPos + calendarW - 36, topPos + footerTop, 28, 22, "×", RGB(241, 245, 249), RGB(148, 163, 184), 1, "'" & ThisWorkbook.Name & "'!BasicInfoCalendarClose"
    AddCalendarShape ws, "Title", msoShapeRectangle, leftPos + 42, topPos + 8, calendarW - 84, 24, Format$(displayMonth, "yyyy年m月"), RGB(248, 250, 252), RGB(248, 250, 252), 0

    Dim weekNames As Variant
    weekNames = Array("日", "月", "火", "水", "木", "金", "土")

    Dim i As Long
    For i = 0 To 6
        AddCalendarShape ws, "Week" & CStr(i), msoShapeRectangle, leftPos + 8 + cellW * i, topPos + 40, cellW, 20, CStr(weekNames(i)), RGB(241, 245, 249), RGB(226, 232, 240), 1
    Next i

    Dim firstDay As Date, lastDay As Date
    firstDay = DateSerial(Year(displayMonth), Month(displayMonth), 1)
    lastDay = DateSerial(Year(displayMonth), Month(displayMonth) + 1, 0)

    Dim dayDate As Date, dayNumber As Long, rowIndex As Long, colIndex As Long
    Dim shapeText As String, fillColor As Long, lineColor As Long
    For dayNumber = CLng(firstDay) To CLng(lastDay)
        dayDate = CDate(dayNumber)
        colIndex = Weekday(dayDate, vbSunday) - 1
        rowIndex = ((Day(dayDate) + Weekday(firstDay, vbSunday) - 2) \ 7)

        fillColor = RGB(255, 255, 255)
        lineColor = RGB(226, 232, 240)
        Select Case Weekday(dayDate, vbSunday)
            Case vbSunday
                fillColor = RGB(252, 231, 243)
                lineColor = RGB(244, 114, 182)
            Case vbSaturday
                fillColor = RGB(224, 242, 254)
                lineColor = RGB(56, 189, 248)
        End Select
        If CLng(dayDate) = CLng(baseDate) Then
            fillColor = RGB(219, 234, 254)
            lineColor = RGB(37, 99, 235)
        End If
        If CLng(dayDate) = CLng(Date) Then
            fillColor = RGB(220, 252, 231)
            lineColor = RGB(22, 163, 74)
        End If

        shapeText = CStr(Day(dayDate))
        Dim dayShape As Shape
        Set dayShape = AddCalendarShape(ws, "Day" & Format$(dayDate, "yyyymmdd"), msoShapeRectangle, _
                                        leftPos + 8 + cellW * colIndex, topPos + 64 + cellH * rowIndex, _
                                        cellW, cellH, shapeText, fillColor, lineColor, 1, _
                                        "'" & ThisWorkbook.Name & "'!BasicInfoCalendarSelectDay")
        dayShape.AlternativeText = Format$(dayDate, "yyyy-mm-dd")
    Next dayNumber

    Dim fiscalToday As Date
    fiscalToday = GetFiscalToday(GetYearFromB4(ws))
    AddCalendarShape ws, "Footer", msoShapeRectangle, leftPos + 8, topPos + footerTop, calendarW - 48, 22, _
                     "基準日 " & Format$(fiscalToday, "yyyy年m月d日"), RGB(248, 250, 252), RGB(248, 250, 252), 0

    Application.ScreenUpdating = previousScreenUpdating
End Sub

Private Function AddCalendarShape(ByVal ws As Worksheet, ByVal suffix As String, ByVal shapeType As MsoAutoShapeType, _
                                  ByVal leftPos As Double, ByVal topPos As Double, ByVal shapeW As Double, ByVal shapeH As Double, _
                                  ByVal textValue As String, ByVal fillColor As Long, ByVal lineColor As Long, _
                                  ByVal lineWeight As Double, Optional ByVal macroName As String = "") As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(shapeType, leftPos, topPos, shapeW, shapeH)

    With shp
        .Name = CALENDAR_PREFIX & suffix
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = lineWeight
        .Placement = xlMove
        If macroName <> "" Then .OnAction = macroName
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.text = textValue
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            With .TextRange.Font
                .NameFarEast = "Meiryo UI"
                .Name = "Meiryo UI"
                .Size = 10
                .Bold = msoFalse
                .Fill.ForeColor.RGB = RGB(15, 23, 42)
            End With
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
        .ZOrder msoBringToFront
    End With

    Set AddCalendarShape = shp
End Function

Private Sub DeleteBasicInfoCalendar(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If IsCalendarShapeName(ws.Shapes(i).Name) Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function IsCalendarShapeName(ByVal shapeName As String) As Boolean
    IsCalendarShapeName = (Left$(shapeName, Len(CALENDAR_PREFIX)) = CALENDAR_PREFIX) Or _
                          (Left$(shapeName, Len(LEGACY_CALENDAR_PREFIX)) = LEGACY_CALENDAR_PREFIX)
End Function

Private Function GetBasicInfoProjectNameTargetCell(ByVal Target As Range) As Range
    Dim ws As Worksheet
    Set ws = Target.Worksheet

    Dim topLeftCell As Range
    Set topLeftCell = GetMergeAreaTopLeftCell(Target)
    If topLeftCell Is Nothing Then Exit Function

    Dim candidate As Range
    Set candidate = Intersect(topLeftCell, ws.Range("C9"))
    If Not candidate Is Nothing Then
        Set GetBasicInfoProjectNameTargetCell = candidate.Cells(1, 1)
        Exit Function
    End If

    Set candidate = Intersect(Target, ws.Range("C9"))
    If Not candidate Is Nothing Then Set GetBasicInfoProjectNameTargetCell = candidate.Cells(1, 1)
End Function

Private Function NormalizeProjectSelectionOfficeName(ByVal OfficeName As String) As String
    If StrComp(OfficeName, FukuchiyamaOfficeFullNameText(), vbTextCompare) = 0 Then
        NormalizeProjectSelectionOfficeName = FukuchiyamaOfficeSearchNameText()
    Else
        NormalizeProjectSelectionOfficeName = OfficeName
    End If
End Function

Private Function FukuchiyamaOfficeFullNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H798F) & ChrW$(&H77E5) & ChrW$(&H5C71) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240)
    End If
    FukuchiyamaOfficeFullNameText = cached
End Function

Private Function FukuchiyamaOfficeSearchNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H798F) & ChrW$(&H77E5) & ChrW$(&H5C71) & ChrW$(&H51FA) & ChrW$(&H5F35)
    End If
    FukuchiyamaOfficeSearchNameText = cached
End Function

Private Function GetDateTargetCell(ByVal Target As Range) As Range
    Dim ws As Worksheet
    Set ws = Target.Worksheet

    Dim topLeftCell As Range
    Set topLeftCell = GetMergeAreaTopLeftCell(Target)
    If topLeftCell Is Nothing Then Exit Function

    Dim candidate As Range
    Set candidate = Intersect(topLeftCell, ws.Range(BASIC_INFO_DATE_CELLS))
    If Not candidate Is Nothing Then
        Set GetDateTargetCell = candidate.Cells(1, 1)
        Exit Function
    End If

    Set candidate = Intersect(Target, ws.Range(BASIC_INFO_DATE_CELLS))
    If Not candidate Is Nothing Then Set GetDateTargetCell = candidate.Cells(1, 1)
End Function

Private Function GetMergeAreaTopLeftCell(ByVal targetRange As Range) As Range
    If targetRange Is Nothing Then Exit Function

    On Error Resume Next
    Set GetMergeAreaTopLeftCell = targetRange.MergeArea.Cells(1, 1)
    If GetMergeAreaTopLeftCell Is Nothing Then
        Set GetMergeAreaTopLeftCell = targetRange.Cells(1, 1)
    End If
    On Error GoTo 0
End Function

Private Function GetInitialCalendarDate(ByVal targetCell As Range) As Date
    Dim cellDate As Date
    If TryGetDate(targetCell.value, cellDate) Then
        GetInitialCalendarDate = cellDate
    Else
        GetInitialCalendarDate = GetFiscalToday(GetYearFromB4(targetCell.Worksheet))
    End If
End Function

Private Function TryGetDate(ByVal value As Variant, ByRef result As Date) As Boolean
    On Error GoTo ErrorHandler
    If IsDate(value) Then
        result = CDate(value)
        TryGetDate = True
    End If
    Exit Function

ErrorHandler:
    TryGetDate = False
End Function

Private Function DateFromIsoText(ByVal sourceText As String) As Date
    DateFromIsoText = DateSerial(CLng(Left$(sourceText, 4)), CLng(Mid$(sourceText, 6, 2)), CLng(Right$(sourceText, 2)))
End Function

Private Function TryGetDateFromCallerName(ByVal callerName As String, ByRef result As Date) As Boolean
    Dim markerPos As Long
    markerPos = InStrRev(callerName, "Day", -1, vbTextCompare)
    If markerPos = 0 Then Exit Function

    Dim dateText As String
    dateText = Mid$(callerName, markerPos + 3, 8)
    If Len(dateText) <> 8 Then Exit Function
    If Not IsNumeric(dateText) Then Exit Function

    result = DateSerial(CLng(Left$(dateText, 4)), CLng(Mid$(dateText, 5, 2)), CLng(Right$(dateText, 2)))
    TryGetDateFromCallerName = True
End Function

Private Function TryGetDateFromSelection(ByRef result As Date) As Boolean
    On Error GoTo ErrorHandler

    Dim shp As Shape
    Set shp = Selection.ShapeRange(1)

    If TryGetDateFromCallerName(shp.Name, result) Then
        TryGetDateFromSelection = True
        Exit Function
    End If

    If Len(shp.AlternativeText) > 0 Then
        result = DateFromIsoText(shp.AlternativeText)
        TryGetDateFromSelection = True
    End If
    Exit Function

ErrorHandler:
    TryGetDateFromSelection = False
End Function

Private Function GetFiscalToday(ByVal fiscalYear As Long) As Date
    If fiscalYear <= 0 Then fiscalYear = Year(Date)

    Dim calendarYear As Long
    calendarYear = fiscalYear
    If Month(Date) < 4 Then calendarYear = fiscalYear + 1

    GetFiscalToday = DateSerial(calendarYear, Month(Date), Day(Date))
End Function

Private Function GetYearFromB4(ByVal ws As Worksheet) As Long
    Dim yearText As String
    yearText = CommonExtractYear4Digits(CStr(ws.Range("B4").value))
    If yearText <> "" Then GetYearFromB4 = CLng(yearText)
End Function

' 基本情報シート名を返す（存在しない場合は空文字）。
Private Function GetBasicInfoWorksheetName() As String
    Dim ws As Worksheet
    Set ws = CommonGetBasicInfoWorksheet()
    If Not ws Is Nothing Then GetBasicInfoWorksheetName = ws.Name
End Function

Private Sub StoreCalendarState(ByVal targetCell As Range, ByVal displayMonth As Date)
    mCalendarTargetSheetName = targetCell.Worksheet.Name
    mCalendarTargetAddress = targetCell.Address(False, False)
    mCalendarDisplayMonth = displayMonth

    SetStoredText CALENDAR_TARGET_SHEET_NAME, targetCell.Worksheet.Name
    SetStoredText CALENDAR_TARGET_ADDRESS_NAME, targetCell.Address(False, False)
    SetStoredText CALENDAR_MONTH_NAME, Format$(displayMonth, "yyyy-mm-dd")
End Sub

Private Function GetCalendarTargetAddress() As String
    If mCalendarTargetAddress <> "" Then
        GetCalendarTargetAddress = mCalendarTargetAddress
    Else
        GetCalendarTargetAddress = GetStoredText(CALENDAR_TARGET_ADDRESS_NAME)
    End If
End Function

Private Function GetCalendarDisplayMonth() As Date
    If mCalendarDisplayMonth <> 0 Then
        GetCalendarDisplayMonth = mCalendarDisplayMonth
    Else
        GetCalendarDisplayMonth = DateFromIsoText(GetStoredText(CALENDAR_MONTH_NAME))
    End If
End Function

Private Sub SetStoredText(ByVal nameText As String, ByVal valueText As String)
    On Error Resume Next
    ThisWorkbook.names(nameText).Delete
    On Error GoTo 0

    ThisWorkbook.names.Add Name:=nameText, RefersTo:="=""" & Replace$(valueText, """", """""") & """", Visible:=False
End Sub

Private Function GetStoredText(ByVal nameText As String) As String
    On Error GoTo ErrorHandler
    GetStoredText = CStr(Application.Evaluate(ThisWorkbook.names(nameText).RefersTo))
    Exit Function

ErrorHandler:
    GetStoredText = ""
End Function

Private Function GetStoredCalendarWorksheet() As Worksheet
    On Error Resume Next
    If mCalendarTargetSheetName <> "" Then
        Set GetStoredCalendarWorksheet = ThisWorkbook.Worksheets(mCalendarTargetSheetName)
    Else
        Set GetStoredCalendarWorksheet = ThisWorkbook.Worksheets(GetStoredText(CALENDAR_TARGET_SHEET_NAME))
    End If
    On Error GoTo 0
End Function


