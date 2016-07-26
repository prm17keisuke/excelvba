Attribute VB_Name = "Common"
Option Explicit


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
    .Cells.ClearFormats     ' ������S�č폜
  End With
End Sub


' �ꗗ�̃w�b�_��ݒ肷��
Public Function setHeader(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "�w�b�_�ݒ�J�n"
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
  ' �t�B���^�̐ݒ�
  With ws
    .Range(.Cells(row, col), .Cells(row, col + i)).AutoFilter
  End With
  Application.StatusBar = ws.Name & "�w�b�_�ݒ�I��"
End Function


' �ꗗ�̃f�[�^����ݒ肷��
Public Function setData(ByRef ws As Worksheet, ByRef rs As ADODB.Recordset, ByVal row As Integer, ByVal col As Integer)
  Application.StatusBar = ws.Name & "�f�[�^�ݒ�J�n"
  ws.Activate
  ws.Cells(row, col).CopyFromRecordset rs
  ws.Cells(row, col).Select
  ws.Range(Selection, Selection.End(xlDown)).Select
  ws.Range(Selection, Selection.End(xlToRight)).Select
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

