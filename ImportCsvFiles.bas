'================================================================================
' Module: Module1
' Author: Gemini
' Date:   2025-07-05
'
' Description:
'   SQLの実行結果など、複数のCSVファイルを選択してExcelに取り込むためのマクロ。
'   選択されたCSVファイルごとに新しいワークシートを作成し、
'   シート名をCSVのファイル名（拡張子なし）に設定します。
'
' CSVの前提条件:
'   - 文字コード: UTF-8
'   - 区切り文字: カンマ (,)
'   - ヘッダー行: あり
'================================================================================

Option Explicit

' メインプロシージャ
Sub ImportCsvFiles()
    
    ' --- 変数宣言 ---
    Dim fileDialog As FileDialog
    Dim selectedFile As Variant
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim fullFilePath As String
    
    ' --- 画面更新を停止して処理を高速化 ---
    Application.ScreenUpdating = False
    
    ' --- ファイル選択ダイアログの準備 ---
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "インポートするCSVファイルを選択してください（複数選択可）"
        .AllowMultiSelect = True ' 複数ファイルの選択を許可
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        
        ' ダイアログを表示し、ファイルが選択されたかチェック
        If .Show = -1 Then '「開く」がクリックされた場合
            
            ' --- 選択されたファイルごとにループ処理 ---
            For Each selectedFile In .SelectedItems
                fullFilePath = CStr(selectedFile)
                
                ' --- 1. 新しいワークシートを作成 ---
                ' 最後のシートの後ろに追加
                Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
                
                ' --- 2. シート名をファイル名から生成 ---
                ' フルパスからファイル名部分のみを抽出 (例: "C:\data\test.csv" -> "test")
                sheetName = Mid(fullFilePath, InStrRev(fullFilePath, "\") + 1)
                sheetName = Replace(sheetName, ".csv", "", Compare:=vbTextCompare)
                
                ' Excelのシート名の禁則文字や長さを考慮（簡易版）
                sheetName = Left(ReplaceCharsForSheetName(sheetName), 31)
                
                ' 重複しないようにシート名を設定
                newSheet.Name = GetUniqueSheetName(sheetName)

                ' --- 3. CSVデータを新しいシートにインポート ---
                ' QueryTablesを使用して、文字化けや型変換エラーを防ぐ
                With newSheet.QueryTables.Add(Connection:="TEXT;" & fullFilePath, Destination:=newSheet.Range("A1"))
                    .Name = sheetName ' クエリ名
                    .FieldNames = True ' 1行目をヘッダーとする
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .RefreshStyle = xlInsertDeleteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .TextFilePromptOnRefresh = False
                    .TextFilePlatform = 65001 ' UTF-8 を指定
                    .TextFileStartRow = 1
                    .TextFileParseType = xlDelimited
                    .TextFileTextQualifier = xlTextQualifierDoubleQuote
                    .TextFileConsecutiveDelimiter = False
                    .TextFileTabDelimiter = False
                    .TextFileSemicolonDelimiter = False
                    .TextFileCommaDelimiter = True ' カンマ区切りを指定
                    .TextFileSpaceDelimiter = False
                    
                    ' 全ての列を文字列として取り込む（予期せぬ型変換を防ぐため）
                    ' ※列数が多い場合は、必要に応じてこの部分を調整してください。
                    Dim columnCount As Integer
                    columnCount = GetCsvColumnCount(fullFilePath)
                    If columnCount > 0 Then
                        Dim dataTypes() As Long
                        ReDim dataTypes(1 To columnCount)
                        Dim i As Integer
                        For i = 1 To columnCount
                            dataTypes(i) = xlTextFormat ' xlTextFormatは「2」
                        Next i
                        .TextFileColumnDataTypes = dataTypes
                    End If

                    .Refresh BackgroundQuery:=False ' 同期処理で実行
                    .Delete ' インポート後に接続情報を削除
                End With
            Next selectedFile
            
            MsgBox "CSVファイルのインポートが完了しました。", vbInformation
            
        Else ' 「キャンセル」がクリックされた場合
            MsgBox "処理はキャンセルされました。", vbInformation
        End If
    End With
    
    ' --- 後処理 ---
    Set fileDialog = Nothing
    Set newSheet = Nothing
    Application.ScreenUpdating = True ' 画面更新を再開
    
End Sub

' シート名に使えない文字を置換する補助関数
Private Function ReplaceCharsForSheetName(ByVal name As String) As String
    Dim invalidChars As String
    Dim i As Integer
    invalidChars = ":\/?*[]" ' シート名に使えない文字
    
    For i = 1 To Len(invalidChars)
        name = Replace(name, Mid(invalidChars, i, 1), "_")
    Next i
    ReplaceCharsForSheetName = name
End Function

' 重複しないユニークなシート名を返す補助関数
Private Function GetUniqueSheetName(ByVal baseName As String) As String
    Dim tempName As String
    Dim counter As Integer
    Dim ws As Worksheet
    Dim isUnique As Boolean
    
    tempName = baseName
    counter = 1
    
    Do
        isUnique = True
        For Each ws In ThisWorkbook.Worksheets
            If LCase(ws.Name) = LCase(tempName) Then
                isUnique = False
                Exit For
            End If
        Next ws
        
        If Not isUnique Then
            counter = counter + 1
            tempName = Left(baseName, 31 - Len(CStr(counter)) - 2) & " (" & CStr(counter) & ")"
        End If
    Loop While Not isUnique
    
    GetUniqueSheetName = tempName
End Function

' CSVファイルの列数を取得する補助関数
Private Function GetCsvColumnCount(ByVal filePath As String) As Integer
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim fileStream As Object
    Dim firstLine As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.OpenTextFile(filePath, 1) ' ForReading
    
    If Not fileStream.AtEndOfStream Then
        firstLine = fileStream.ReadLine
        GetCsvColumnCount = UBound(Split(firstLine, ",")) + 1
    Else
        GetCsvColumnCount = 0
    End If
    
    fileStream.Close
    Set fso = Nothing
    Set fileStream = Nothing
    Exit Function
    
ErrorHandler:
    GetCsvColumnCount = 0 ' エラー時は0を返す
End Function
