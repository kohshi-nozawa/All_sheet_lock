Attribute VB_Name = "Module1"
Sub Lock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excel�u�b�N,*.xlsx", Title:="�V�[�g�����b�N����G�N�Z���t�@�C����I��")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "�L�����Z������܂���"
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass1 As String, pass2 As String
  pass1 = InputBox("�ݒ肷��p�X���[�h�����", "�p�X���[�h�ݒ�", "Passw0rd")
  pass2 = InputBox("�p�X���[�h�̊m�F", "�p�X���[�h�m�F", "Passw0rd")
  If pass1 <> pass2 Then
    MsgBox "�p�X���[�h����v���܂���"
    Exit Sub
  End If
  For Each s In wb1.Sheets
    s.Protect Password:=pass1,
    AllowFiltering:=True
  Next s
End Sub