Attribute VB_Name = "Common"
Option Explicit


' 共通開始処理
Public Sub commonStart()
  Application.DisplayAlerts = False ' アラート表示OFF
  Application.ScreenUpdating = False ' 画面更新OFF
  Application.Calculation = xlCalculationManual ' 自動計算OFF
  Application.StatusBar = "処理を開始します"
End Sub


' 共通終了処理
Public Sub commonEnd()
  Application.DisplayAlerts = True ' アラート表示ON
  Application.ScreenUpdating = True ' 画面更新ON
  Application.Calculation = xlCalculationAutomatic ' 自動計算ON
End Sub


' シートを初期化する
Public Sub initializeSheet(ByRef ws As Worksheet)
  With ws
    .DrawingObjects.Delete  ' 図形・画像を全て削除
    .Cells.ClearContents    ' 値を全て削除
    .Cells.ClearComments    ' コメントを全て削除
    .Cells.ClearFormats     ' 書式を全て削除
  End With
End Sub


' 一覧のヘッダを設定する
Public Function setHeader(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "ヘッダ設定開始"
  Dim i As Integer
  
  For i = 1 To rs.Fields.Count
    With ws.Cells(row, col + i - 1)
      .Interior.Color = RGB(135, 206, 235)
      .Borders.LineStyle = xlContinuous
      .HorizontalAlignment = xlLeft
      .Font.Bold = True
      .ColumnWidth = 9
      .Value = rs.Fields(i - 1).Name
    End With
  Next
  ' フィルタの設定
  With ws
    .Range(.Cells(row, col), .Cells(row, col + i)).AutoFilter
  End With
  Application.StatusBar = ws.Name & "ヘッダ設定終了"
End Function


' 一覧のデータ部を設定する
Public Function setData(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "データ設定開始"
  ws.Activate
  ws.Cells(row, col).CopyFromRecordset rs
  ws.Cells(row, col).Select
  ws.Range(Selection, Selection.End(xlDown)).Select
  ws.Range(Selection, Selection.End(xlToRight)).Select
  Selection.Borders.LineStyle = xlContinuous
  Application.StatusBar = ws.Name & "データ設定終了"
End Function


'SQLシートからSQLを生成する
Public Function createSQL(ByRef sqlWs As Worksheet)
  Dim i As Integer
  i = 1
  Do
    ' コメント行はスキップする
    If Not (sqlWs.Cells(i, 1).Value Like "--*") Then
      createSQL = createSQL & " " & sqlWs.Cells(i, 1).Value
    End If
    i = i + 1
  Loop While sqlWs.Cells(i, 1).Value <> ""
End Function

