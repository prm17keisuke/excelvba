Attribute VB_Name = "DBAccess"
Option Explicit


'Oracle �ɐڑ����n���h����Ԃ�
'
Public Sub s_oraconnect(ByRef cn As Variant, data_source As String, USER_ID As String, PASSWORD As String)
    On Error GoTo ERR_HANDLER
       
    'ADO��Connection�I�u�W�F�N�g���쐬
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & data_source & ";" & "User ID=" & USER_ID & ";Password=" & PASSWORD
    cn.Open
  
    On Error GoTo 0
    
    Exit Sub
ERR_HANDLER:    '�G���[����
    '�G���[�ԍ��ƃG���[���e�̕\��
    MsgBox Err.Number & vbLf & Err.Description
    Err.Clear
    End
End Sub

