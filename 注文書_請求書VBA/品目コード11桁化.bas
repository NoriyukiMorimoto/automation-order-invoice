Attribute VB_Name = "品目コード11桁化"

Option Explicit

'==============================================================================
' 品目コード11桁化マクロ
'
' 【処理概要】
'   1. D列の数値に1000を加算（4桁化）
'   2. F列を2桁ゼロ埋め（例: 1 → "01"）
'   3. H列を5桁ゼロ埋め（例: 1 → "00001"）
'   4. D列(4桁) + F列(2桁) + H列(5桁) を結合し11桁コードをA列に挿入
'   5. A列の文字列コードを数値に変換
'   6. 元のC列をデータ最右端へ移動
'   7. 元のB・C・D・F・H・L・M列を削除
'   8. 全セルを上下中央揃え
'   9. データ範囲をテーブル設定（テーブル名=元C2の値、スタイル=オレンジ中間7）
'  10. ヘッダーが5行目になるように行を挿入（上に4行追加）
'  11. 5行目・A列・E列・G列・H列を左右中央揃え
'  12. F列を桁区切り・小数点なし書式に設定
'  13. E5・F5のヘッダー名を設定（単位・単価）
'  14. A1:G1を結合しタイトルを入力
'  15. 全セルのフォントをBIZ UDGothicに設定
'  16. 列幅を自動調整
'
' 【実行範囲】D列の2行目（データ開始行）から最終行まで
'
' 【タイトル生成ルール】
'   ファイル名例: "2024早期発注〇〇保線区購入充当単価表.xlsx"
'   → 先頭4桁の数値＋"年度"＋"購入充当単価表"＋ファイル名頭4文字
'   → "2024年度購入充当単価表早期発注"
'==============================================================================

Sub BuildCode11_New()

    Const FIRST_ROW  As Long = 2   ' データ開始行（1行目がヘッダー、2行目からデータ）
    Const HEADER_ROW As Long = 1   ' 現在のヘッダー行

    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim r           As Long

    Dim dVal        As Variant
    Dim fVal        As Variant
    Dim hVal        As Variant

    Dim d4          As String
    Dim f2          As String
    Dim h5          As String
    Dim code11      As String

    Dim tableRange  As Range
    Dim tbl         As ListObject
    Dim tableName   As String
    Dim lastCol     As Long
    Dim insertRows  As Long

    Set ws = ActiveSheet

    '--------------------------------------------------------------------------
    ' 0. テーブル名用に元C2の値を取得（処理前に保存）
    '--------------------------------------------------------------------------
    tableName = CStr(ws.Cells(2, "C").Value)
    tableName = CleanTableName(tableName)
    If tableName = "" Then tableName = "品目コードTable"

    '--------------------------------------------------------------------------
    ' 1. D列の最終行を取得
    '--------------------------------------------------------------------------
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    If lastRow < FIRST_ROW Then
        MsgBox "D列の2行目以降にデータがありません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    '--------------------------------------------------------------------------
    ' 2. A列を文字列フォーマットに設定（先頭ゼロ保持のため）
    '--------------------------------------------------------------------------
    ws.Range("A:A").NumberFormat = "@"

    '--------------------------------------------------------------------------
    ' 3. 各行について11桁コードを生成しA列に書き込む
    '--------------------------------------------------------------------------
    For r = FIRST_ROW To lastRow

        dVal = ws.Cells(r, "D").Value
        fVal = ws.Cells(r, "F").Value
        hVal = ws.Cells(r, "H").Value

        ' --- D列：数値 + 1000 → 4桁 ---
        If IsNumeric(dVal) And dVal <> "" Then
            d4 = Format$(CLng(dVal) + 1000, "0000")
        Else
            d4 = "0000"
        End If

        ' --- F列：数字のみ抽出 → 2桁ゼロ埋め ---
        f2 = PadZero(DigitsOnly(CStr(fVal)), 2)

        ' --- H列：数字のみ抽出 → 5桁ゼロ埋め ---
        h5 = PadZero(DigitsOnly(CStr(hVal)), 5)

        ' --- 結合して11桁コード ---
        code11 = d4 & f2 & h5

        ' --- A列に文字列として書き込む ---
        ws.Cells(r, "A").Value = code11

    Next r

    '--------------------------------------------------------------------------
    ' 4. A列の文字列コードを数値に変換
    '--------------------------------------------------------------------------
    ws.Range("A:A").NumberFormat = "0"
    Dim cellA As Range
    For r = FIRST_ROW To lastRow
        Set cellA = ws.Cells(r, "A")
        If cellA.Value <> "" Then
            cellA.Value = CDec(cellA.Value)
        End If
    Next r

    '--------------------------------------------------------------------------
    ' 5. 元のC列を最右端へ移動
    '    ※ 削除対象列を消す前にC列を右端へ退避
    '--------------------------------------------------------------------------
    lastCol = ws.Cells(FIRST_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column Then
        lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    End If

    ' C列全体をlastCol+1列へコピーして元C列をクリア
    ws.Columns("C").Copy Destination:=ws.Columns(lastCol + 1)
    ws.Columns("C").UnMerge  ' 結合セルがある場合に解除してからクリア
    ws.Columns("C").ClearContents

    '--------------------------------------------------------------------------
    ' 6. 不要列を削除（右から順に削除してアドレスのずれを防ぐ）
    '    削除対象：M, L, H, F, D, C（空）, B 列（元データ基準・列番号降順）
    '--------------------------------------------------------------------------
    ws.Columns("M").Delete Shift:=xlShiftToLeft   ' M列削除
    ws.Columns("L").Delete Shift:=xlShiftToLeft   ' L列削除（元データのL列）
    ws.Columns("H").Delete Shift:=xlShiftToLeft   ' H列削除
    ws.Columns("F").Delete Shift:=xlShiftToLeft   ' F列削除
    ws.Columns("D").Delete Shift:=xlShiftToLeft   ' D列削除
    ws.Columns("C").Delete Shift:=xlShiftToLeft   ' C列（空）削除
    ws.Columns("B").Delete Shift:=xlShiftToLeft   ' B列削除

    '--------------------------------------------------------------------------
    ' 7. 全セルを上下中央揃えに設定
    '--------------------------------------------------------------------------
    ws.Cells.VerticalAlignment = xlVAlignCenter

    '--------------------------------------------------------------------------
    ' 8. テーブル設定（列削除後の範囲で設定）
    '--------------------------------------------------------------------------
    Dim newLastCol As Long
    Dim newLastRow As Long
    newLastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    newLastRow  = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set tableRange = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(newLastRow, newLastCol))

    ' 既存テーブルがあれば解除
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    ' テーブル作成
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=tableRange, _
        XlListObjectHasHeaders:=xlYes)

    ' テーブル名設定
    tbl.Name = tableName

    ' スタイル設定：オレンジ・テーブルスタイル（中間）7
    tbl.TableStyle = "TableStyleMedium7"

    '--------------------------------------------------------------------------
    ' 9. ヘッダーが5行目になるように行を挿入（現在1行目 → 上に4行追加）
    '--------------------------------------------------------------------------
    insertRows = 4
    ws.Rows("1:" & insertRows).Insert Shift:=xlDown

    '--------------------------------------------------------------------------
    ' 10. 5行目（ヘッダー行）・A列・E列・G列・H列を左右中央揃えに設定
    '     ※ 行挿入後なのでヘッダーは5行目に移動済み
    '--------------------------------------------------------------------------
    ws.Rows(5).HorizontalAlignment = xlHAlignCenter
    ws.Columns("A").HorizontalAlignment = xlHAlignCenter
    ws.Columns("E").HorizontalAlignment = xlHAlignCenter
    ws.Columns("G").HorizontalAlignment = xlHAlignCenter
    ws.Columns("H").HorizontalAlignment = xlHAlignCenter

    '--------------------------------------------------------------------------
    ' 11. F列を桁区切り・小数点なし書式に設定（列削除・行挿入後）
    '--------------------------------------------------------------------------
    ws.Columns("F").NumberFormat = "#,##0"

    '--------------------------------------------------------------------------
    ' 12. E5・F5のヘッダー名を設定
    '--------------------------------------------------------------------------
    ws.Cells(5, "E").Value = "単位"
    ws.Cells(5, "F").Value = "単価"

    '--------------------------------------------------------------------------
    ' 13. A1:G1を結合してタイトルを入力
    '     タイトル形式：ファイル名先頭4桁数値 + "年度" + "購入充当単価表"
    '                   + ファイル名先頭4文字（早期発注 or 設計変更）
    '
    '     ファイル名例：「2024早期発注〇〇保線区購入充当単価表.xlsx」
    '     → 4桁数値部分 = "2024"（先頭から数字のみ抽出して4桁）
    '     → 先頭4文字   = "早期発注" or "設計変更"（数字除去後の先頭4文字）
    '--------------------------------------------------------------------------
    Dim wb          As Workbook
    Dim fileName    As String
    Dim yearStr     As String
    Dim prefixStr   As String
    Dim titleStr    As String
    Dim titleRange  As Range

    Set wb = ws.Parent
    ' 拡張子を除いたファイル名を取得
    fileName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)

    ' 【】の直前にある4桁数値を年度として抽出
    ' 例: "早期発注_1104材料購入充当工種一覧表兼(単価一覧表2026【京都保線区】"
    '     → 【の直前を遡って数字を収集 → "2026"
    ' ファイル名中の「20」で始まる4桁数字を年度として取得
    yearStr = ""
    Dim yIdx As Long
    Dim yCh  As String
    Dim yBuf As String
    For yIdx = 1 To Len(fileName) - 3
        yBuf = Mid(fileName, yIdx, 4)
        yBuf = StrConv(yBuf, vbNarrow)
        If Left(yBuf, 2) = "20" Then
            Dim allDigit As Boolean
            allDigit = True
            Dim dIdx As Long
            For dIdx = 1 To 4
                yCh = Mid(yBuf, dIdx, 1)
                If yCh < "0" Or yCh > "9" Then
                    allDigit = False
                    Exit For
                End If
            Next dIdx
            If allDigit Then
                yearStr = yBuf
                Exit For
            End If
        End If
    Next yIdx
    ' 見つからない場合はファイル名先頭の数字4桁にフォールバック
    If yearStr = "" Then
        yearStr = DigitsOnly(fileName)
        If Len(yearStr) >= 4 Then yearStr = Left(yearStr, 4)
    End If

    ' ファイル名先頭の4文字（「早期発注」or「設計変更」）を抽出
    ' ※ファイル名の先頭4文字を直接取得（アンダースコアより前の部分）
    Dim fLen As Long
    fLen = Len(fileName)
    Dim fIdx As Long
    Dim fCh  As String
    Dim nonDigitStr As String
    nonDigitStr = ""
    For fIdx = 1 To fLen
        fCh = Mid(fileName, fIdx, 1)
        fCh = StrConv(fCh, vbNarrow)
        If fCh < "0" Or fCh > "9" Then
            nonDigitStr = nonDigitStr & Mid(fileName, fIdx, 1)
        End If
    Next fIdx
    ' 非数字部分の先頭4文字を取得（「早期発注」or「設計変更」）
    If Len(nonDigitStr) >= 4 Then
        prefixStr = Left(nonDigitStr, 4)
    Else
        prefixStr = nonDigitStr
    End If

    ' G6セルの内容を取得（行挿入後なのでヘッダーは5行目、データ開始は6行目）
    Dim g6Val As String
    g6Val = CStr(ws.Cells(6, "G").Value)

    ' タイトル文字列を生成
    ' 形式：2026年度_購入充当単価表_早期発注_大阪保線区
    titleStr = yearStr & "年度" & "_" & "購入充当単価表" & "_" & prefixStr & "_" & g6Val

    ' A1:G1を結合
    Set titleRange = ws.Range("A1:G1")
    With titleRange
        .Merge
        .Value = titleStr
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Name = "BIZ UDGothic"
        .Font.Size = 14
    End With

    '--------------------------------------------------------------------------
    ' 14. 全セルのフォントをBIZ UDGothicに設定（サイズは変更しない）
    '     ※タイトルセルはフォントサイズ14を維持するため、データ範囲のみ適用
    '--------------------------------------------------------------------------
    ws.Cells.Font.Name = "BIZ UDGothic"
    ' タイトルセルのフォントサイズを14に再設定（全セル適用で上書きされる場合に備えて）
    ws.Range("A1").Font.Size = 14

    '--------------------------------------------------------------------------
    ' 15. 列幅を自動調整
    '--------------------------------------------------------------------------
    ws.Cells.EntireColumn.AutoFit

End Sub

'==============================================================================
' 【補助関数】数字のみ抽出（全角→半角変換してから抽出）
'==============================================================================
Private Function DigitsOnly(ByVal v As String) As String
    Dim t   As String
    Dim i   As Long
    Dim ch  As String
    Dim res As String

    t = StrConv(v, vbNarrow)    ' 全角数字を半角へ変換
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then res = res & ch
    Next i
    DigitsOnly = res
End Function

'==============================================================================
' 【補助関数】指定桁数に左ゼロ埋め（桁超過時は右端を優先）
'==============================================================================
Private Function PadZero(ByVal s As String, ByVal width As Long) As String
    Dim t As String
    t = DigitsOnly(s)
    If t = "" Then t = "0"

    If Len(t) >= width Then
        PadZero = Right$(t, width)      ' 桁超過時は末尾width桁を採用
    Else
        PadZero = String$(width - Len(t), "0") & t
    End If
End Function

'==============================================================================
' 【補助関数】テーブル名に使用できない文字を除去
'==============================================================================
Private Function CleanTableName(ByVal s As String) As String
    Dim res As String
    Dim i   As Long
    Dim ch  As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        ' 英数字・日本語・アンダースコアのみ許可、それ以外はアンダースコアに置換
        Select Case True
            Case ch >= "A" And ch <= "Z"
            Case ch >= "a" And ch <= "z"
            Case ch >= "0" And ch <= "9"
            Case ch >= "ぁ" And ch <= "ん"
            Case ch >= "ァ" And ch <= "ン"
            Case ch >= "亜" And ch <= "??"
            Case ch = "_"
            Case Else
                ch = "_"
        End Select
        res = res & ch
    Next i

    ' 先頭が数字の場合はアンダースコアを付加（テーブル名は数字始まり不可）
    If Len(res) > 0 Then
        If res Like "[0-9]*" Then res = "_" & res
    End If

    CleanTableName = res
End Function
