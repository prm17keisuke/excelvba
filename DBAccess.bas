Attribute VB_Name = "DBAccess"
Option Explicit


'Oracle に接続しハンドルを返す
'
Public Sub s_oraconnect(ByRef cn As Variant, data_source As String, USER_ID As String, PASSWORD As String)
    On Error GoTo ERR_HANDLER
       
    'ADOのConnectionオブジェクトを作成
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & data_source & ";" & "User ID=" & USER_ID & ";Password=" & PASSWORD
    cn.Open
  
    On Error GoTo 0
    
    Exit Sub
ERR_HANDLER:    'エラー処理
    'エラー番号とエラー内容の表示
    MsgBox Err.Number & vbLf & Err.Description
    Err.Clear
    End
End Sub

