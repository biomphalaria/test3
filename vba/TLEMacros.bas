Attribute VB_Name = "TLEMacros"
Option Explicit

' 先頭と末尾の全角・半角スペースを除去するヘルパー関数
Private Function TrimEx(ByVal txt As String) As String
    txt = Replace(txt, ChrW(&H3000), " ") ' 全角スペースを半角に
    txt = Replace(txt, Chr(160), " ")    ' ノーブレークスペース
    TrimEx = Trim(txt)
End Function

' COSPAR ID の表記ゆれを統一するためのヘルパー関数
Private Function NormalizeCosparID(ByVal id As String) As String
    id = TrimEx(id)
    ' 一般的なハイフン・ダッシュ類を削除
    id = Replace(id, "-", "")
    id = Replace(id, "‑", "") ' U+2010
    id = Replace(id, "–", "") ' U+2013
    id = Replace(id, "—", "") ' U+2014
    id = Replace(id, "−", "") ' U+2212
    id = Replace(id, "－", "") ' U+FF0D
    id = Replace(id, " ", "")

    ' 4 桁年の場合は TLE 形式の 2 桁年に変換
    If Len(id) = 8 Then
        NormalizeCosparID = Right(Left(id, 4), 2) & Mid(id, 5)
    Else
        NormalizeCosparID = id
    End If
End Function

' インポート用マクロ - Sheet1 のボタンから呼び出す
Sub ImportTLE()
    Dim fd As Object 'FileDialog
    Dim FileName As String
    Dim FileNum As Integer
    Dim lineText As String
    Dim row As Long
    
    ' ファイル選択ダイアログ
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "TLE ファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "Text Files", "*.txt;*.tle", 1
    If fd.Show <> -1 Then Exit Sub
    FileName = fd.SelectedItems(1)
    
    ' Sheet2 に内容を貼り付け
    FileNum = FreeFile
    Open FileName For Input As #FileNum
    Sheets("Sheet2").Cells.Clear
    row = 1
    Do Until EOF(FileNum)
        Line Input #FileNum, lineText
        Sheets("Sheet2").Cells(row, 1).Value = lineText
        row = row + 1
    Loop
    Close #FileNum
End Sub

' Sheet5 のデータをテキストとして保存 - Sheet5 のボタンから呼び出す
Sub ExportTLE()
    Dim fd As Object 'FileDialog
    Dim FileName As String
    Dim FileNum As Integer
    Dim row As Long
    Dim lastRow As Long
    Dim sheetOut As Worksheet

    Set sheetOut = Sheets("Sheet5")

    lastRow = sheetOut.Cells(sheetOut.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        MsgBox "Sheet5 にデータがありません", vbExclamation
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "TLE を保存するファイルを指定してください"
    fd.InitialFileName = "export.tle"
    fd.Filters.Clear
    fd.Filters.Add "Text Files", "*.txt;*.tle", 1
    If fd.Show <> -1 Then Exit Sub
    FileName = fd.SelectedItems(1)

    FileNum = FreeFile
    Open FileName For Output As #FileNum
    For row = 1 To lastRow
        Print #FileNum, sheetOut.Cells(row, 1).Value
    Next row
    Close #FileNum

    MsgBox "保存が完了しました", vbInformation
End Sub
' 着目衛星のTLEを抽出 - Sheet3 のボタンから呼び出す
Sub ExtractFocusedSatellites()
    Dim lastRow As Long
    Dim tleLastRow As Long
    Dim outRow As Long
    Dim i As Long, j As Long
    Dim nameVal As String, noradVal As String, cosparVal As String
    Dim tleName As String, tleNorad As String, tleCospar As String
    Dim line2 As String
    Dim sheetTLE As Worksheet, sheetSat As Worksheet, sheetOut As Worksheet

    Set sheetTLE = Sheets("Sheet2")
    Set sheetSat = Sheets("Sheet3")
    Set sheetOut = Sheets("Sheet5")

    sheetOut.Cells.Clear

    lastRow = sheetSat.Cells(sheetSat.Rows.Count, 1).End(xlUp).Row
    tleLastRow = sheetTLE.Cells(sheetTLE.Rows.Count, 1).End(xlUp).Row
    outRow = 1

    For i = 2 To lastRow
        nameVal = TrimEx(CStr(sheetSat.Cells(i, 1).Value))
        noradVal = Replace(TrimEx(CStr(sheetSat.Cells(i, 2).Value)), " ", "")
        cosparVal = NormalizeCosparID(CStr(sheetSat.Cells(i, 3).Value))

        j = 1
        Do While j <= tleLastRow
            ' 空行をスキップ
            Do While j <= tleLastRow And TrimEx(CStr(sheetTLE.Cells(j, 1).Value)) = ""
                j = j + 1
            Loop
            If j + 2 > tleLastRow Then Exit Do

            tleName = TrimEx(CStr(sheetTLE.Cells(j, 1).Value))
            line2 = CStr(sheetTLE.Cells(j + 1, 1).Value)

            If Len(line2) >= 17 Then
                tleNorad = Trim(Mid(line2, 3, 5))
                tleCospar = NormalizeCosparID(Mid(line2, 10, 8))
            Else
                tleNorad = ""
                tleCospar = ""
            End If

            If (nameVal <> "" And StrComp(tleName, nameVal, vbTextCompare) = 0) _
                Or (noradVal <> "" And StrComp(tleNorad, noradVal, vbTextCompare) = 0) _
                Or (cosparVal <> "" And StrComp(tleCospar, cosparVal, vbTextCompare) = 0) Then
                sheetOut.Cells(outRow, 1).Value = sheetTLE.Cells(j, 1).Value
                sheetOut.Cells(outRow + 1, 1).Value = sheetTLE.Cells(j + 1, 1).Value
                sheetOut.Cells(outRow + 2, 1).Value = sheetTLE.Cells(j + 2, 1).Value
                outRow = outRow + 3
            End If

            j = j + 3
        Loop
    Next i
End Sub

