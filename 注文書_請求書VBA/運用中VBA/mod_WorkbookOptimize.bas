Option Explicit

' =============================================================
' 改修履歴: CHANGELOG.md 参照
'
' Excel の「ブックのパフォーマンス（未使用の書式設定とメタデータ）」警告の
' 原因となる、空セルに残った過剰な書式(配置・番号形式・テキストプロパティ)を
' クリアして UsedRange を縮小する。保存前(ThisWorkbook.Workbook_BeforeSave)から
' 自動実行される想定。手動実行用に OptimizeWorkbookNow も公開する。
'
' 方針:
'  - 基本情報シート … 業者ブロックは最大10社(値列 F..AG=33列目)まで使用する。
'    10社分の空ブロックには背景色・斜線ガイド等の意図した書式があるため消さない。
'    業者ブロック範囲より右(34列目以降)の空セル書式のみクリアする。
'  - 工事単価シート … 触らない。罫線・塗り・桁区切りは VBA で意図的に付与するため、
'    Find ベースの「値の外側」判定で ClearFormats すると誤って消える。
'  - それ以外のシート … 触らない(誤って装飾を消さないための安全側)。
' =============================================================

' 基本情報シートで業者ブロックが使用しうる最終列(AG=33)。この右からクリア対象。
Private Const BASIC_INFO_VENDOR_LAST_COL As Long = 33

Public Sub OptimizeWorkbookNow(Optional ByVal wb As Workbook)
    OptimizeWorkbookInternal wb, True
End Sub

' 保存前フックから呼ぶ静音版(メッセージ・確認なし)
Public Sub OptimizeWorkbookSilent(Optional ByVal wb As Workbook)
    OptimizeWorkbookInternal wb, False
End Sub

Private Sub OptimizeWorkbookInternal(ByVal wb As Workbook, ByVal showResult As Boolean)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.screenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo CleanExit
    Application.screenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim basicInfoName As String
    basicInfoName = CommonGetBasicInfoSheetName(wb)

    Dim clearedSheets As Long
    Dim ws As Worksheet
    For Each ws In wb.worksheets
        If StrComp(ws.Name, basicInfoName, vbTextCompare) = 0 Then
            If OptimizeBasicInfoSheet(ws) Then clearedSheets = clearedSheets + 1
        End If
    Next ws

    Application.Calculate

    mod_DebugLog.Log "[WorkbookOptimize] cleaned unused formats on " & clearedSheets & " sheet(s)"

    If showResult Then
        MsgBox OptimizeDoneMessageText(), vbInformation, OptimizeTitleText()
    End If

CleanExit:
    If Err.Number <> 0 Then
        mod_DebugLog.Log "[WorkbookOptimize] error " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.screenUpdating = prevScreen
End Sub

' 基本情報シート: 業者ブロック使用範囲(AG=33列)より右の空セル書式のみクリア
Private Function OptimizeBasicInfoSheet(ByVal ws As Worksheet) As Boolean
    Dim usedLastRow As Long
    Dim usedLastCol As Long
    usedLastRow = GetSheetUsedLastRow(ws)
    usedLastCol = GetSheetUsedLastCol(ws)
    If usedLastCol <= BASIC_INFO_VENDOR_LAST_COL Then Exit Function
    If usedLastRow < 1 Then Exit Function

    Dim clearRange As Range
    Set clearRange = ws.Range(ws.Cells(1, BASIC_INFO_VENDOR_LAST_COL + 1), _
                              ws.Cells(usedLastRow, usedLastCol))
    ResetRangeFormats clearRange
    OptimizeBasicInfoSheet = True
End Function

' 書式を既定へ戻す。結合セルがあれば解除してからクリアする。
Private Sub ResetRangeFormats(ByVal target As Range)
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    If target.MergeCells Then target.UnMerge
    target.ClearFormats
    target.NumberFormatLocal = "General"
    target.HorizontalAlignment = xlGeneral
    target.VerticalAlignment = xlBottom
    target.WrapText = False
    target.Orientation = 0
    On Error GoTo 0
End Sub

' UsedRange の末尾行(書式のみのセルも含む)
Private Function GetSheetUsedLastRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetSheetUsedLastRow = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1
    On Error GoTo 0
End Function

' UsedRange の末尾列(書式のみのセルも含む)
Private Function GetSheetUsedLastCol(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetSheetUsedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    On Error GoTo 0
End Function

Private Function OptimizeTitleText() As String
    OptimizeTitleText = ChrW$(&H30D6) & ChrW$(&H30C3) & ChrW$(&H30AF) & _
                        ChrW$(&H306E) & ChrW$(&H6700) & ChrW$(&H9069) & ChrW$(&H5316)
End Function

Private Function OptimizeDoneMessageText() As String
    OptimizeDoneMessageText = ChrW$(&H672A) & ChrW$(&H4F7F) & ChrW$(&H7528) & ChrW$(&H306E) & _
                              ChrW$(&H66F8) & ChrW$(&H5F0F) & ChrW$(&H3092) & ChrW$(&H6574) & _
                              ChrW$(&H7406) & ChrW$(&H3057) & ChrW$(&H307E) & ChrW$(&H3057) & _
                              ChrW$(&H305F) & ChrW$(&H3002)
End Function


