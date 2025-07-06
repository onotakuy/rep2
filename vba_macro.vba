Sub ImportCsv()
    Dim filePath As Variant
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim fileName As String

    ' ファイル選択ダイアログを表示
    filePath = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select CSV File")

    ' ファイルが選択されなかった場合は終了
    If filePath = False Then
        Exit Sub
    End If

    ' ファイル名（拡張子なし）を取得
    fileName = CreateObject("Scripting.FileSystemObject").GetBaseName(filePath)

    ' 同名のシートが存在するか確認し、存在すれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets(fileName)
    If Not ws Is Nothing Then
        ws.Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 新しいシートを追加してファイル名をシート名に設定
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = fileName

    ' QueryTablesを使ってCSVデータをインポート
    With newWs.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=newWs.Range("A1"))
        .TextFilePlatform = 932 ' 932 = Shift-JIS
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileConsecutiveDelimiter = False
        .RefreshOnFileOpen = False
        .Refresh BackgroundQuery:=False
        .Refresh
        .Delete ' 接続情報を削除してデータのみ残す
    End With

    MsgBox "CSVファイルのインポートが完了しました。" & vbCrLf & "シート名: " & fileName, vbInformation

End Sub
