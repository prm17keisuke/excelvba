Attribute VB_Name = "Common"
Option Explicit


Public Const CELL_NAME_SERVICE_NAME = "SERVICE_NAME"                            ' Oracle�T�[�r�X��
Public Const CELL_NAME_CONTRACT_ENTERPRISE_CD = "CONTRACT_ENTERPRISE_CD"        ' �_���ƃR�[�h
Public Const CELL_NAME_SCHOOLHOUSE_NAME = "SCHOOLHOUSE_NAME"                    ' �Z�ɖ�
Public Const CELL_NAME_KOMA_TYPE = "KOMA_TYPE"                                  ' �R�}���
Public Const CELL_NAME_KOMA_NUM = "KOMA_NUM"                                    ' �R�}��
Public Const CELL_NAME_APPLY_DATE = "APPLY_DATE"                                ' �K�p�J�n��
Public Const CELL_NAME_HOMEWORK_RESOLUTION_SALON = "HOMEWORK_RESOLUTION_SALON"  ' �h������T����

Public Const WS_NAME_CONTROL = "����V�[�g"
Public Const WS_NAME_EX_DATA = "�����E���ԑ�"
Public Const WS_NAME_COURSE_GROUP = "�R�[�X�O���[�v"
Public Const WS_NAME_COURSE_GROUP_SQL = "�R�[�X�O���[�v_SQL"
Public Const WS_NAME_EX_DATA_SQL = "�����E���ԑ�SQL"
Public Const WS_NAME_M_SCENE = "M_SCENE"
Public Const WS_NAME_M_SCENE_SQL = "M_SCENE_SQL"
Public Const WS_NAME_UPD_T_SCENE_HISTORY = "UPD_T_SCENE_HISTORY"

Public Const SQL_PARAM_SCHOOLHOUSE_NAME = ":schoolhouseName"
Public Const SQL_PARAM_APPLY_DATE = ":applyDate"


' ���ʊJ�n����
Public Sub commonStart()
  Application.DisplayAlerts = False ' �A���[�g�\��OFF
  Application.ScreenUpdating = False ' ��ʍX�VOFF
  Application.Calculation = xlCalculationManual ' �����v�ZOFF
  Application.StatusBar = "�������J�n���܂�"
End Sub


' ���ʏI������
Public Sub commonEnd()
  Application.DisplayAlerts = True ' �A���[�g�\��ON
  Application.ScreenUpdating = True ' ��ʍX�VON
  Application.Calculation = xlCalculationAutomatic ' �����v�ZON
End Sub


' �V�[�g������������
Public Sub initializeSheet(ByRef ws As Worksheet)
  With ws
    .DrawingObjects.Delete  ' �}�`�E�摜��S�č폜
    .Cells.ClearContents    ' �l��S�č폜
    .Cells.ClearComments    ' �R�����g��S�č폜
  End With
End Sub


' �ꗗ�̃w�b�_��ݒ肷��
Public Function setHeader(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "�w�b�_�ݒ�J�n"
  Dim i As Integer
  
  For i = 1 To rs.Fields.Count
    With ws.Cells(row, i)
      .Interior.Color = RGB(135, 206, 235)
      .Borders.LineStyle = xlContinuous
      .HorizontalAlignment = xlLeft
      .Font.Bold = True
      .ColumnWidth = 9
      .Value = rs.Fields(i - 1).Name
    End With
  Next
  ' �t�B���^�̐ݒ�
  With ws
    .Range(.Cells(row, 1), .Cells(row, i - 1)).AutoFilter
  End With
  Application.StatusBar = ws.Name & "�w�b�_�ݒ�I��"
End Function


' �ꗗ�̃f�[�^����ݒ肷��
Public Function setData(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "�f�[�^�ݒ�J�n"
  ws.Activate
  ws.Cells(row, col).CopyFromRecordset rs
  ws.Cells(row, col).Select
  ws.Range(Selection, Selection.End(xlToRight)).Select
  ws.Range(Selection, Selection.End(xlDown)).Select
  Selection.Borders.LineStyle = xlContinuous
  Application.StatusBar = ws.Name & "�f�[�^�ݒ�I��"
End Function


'SQL�V�[�g����SQL�𐶐�����
Public Function createSQL(ByRef sqlWs As Worksheet)
  Dim i As Integer
  i = 1
  Do
    ' �R�����g�s�̓X�L�b�v����
    If Not (sqlWs.Cells(i, 1).Value Like "--*") Then
      createSQL = createSQL & " " & sqlWs.Cells(i, 1).Value
    End If
    i = i + 1
  Loop While sqlWs.Cells(i, 1).Value <> ""
End Function


