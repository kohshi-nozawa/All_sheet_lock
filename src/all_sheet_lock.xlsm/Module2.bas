Attribute VB_Name = "Module2"
Sub UnLock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excel�u�b�N,*.xlsx", Title:="�V�[�g�̃��b�N����������G�N�Z���t�@�C����I��")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "�L�����Z������܂���", Buttons:=vbExclamation
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass As String
  pass = InputBox("�����p�̃p�X���[�h�����", "�p�X���[�h����", "Passw0rd")
  For Each s In wb1.Sheets
    s.UnProtect Password:=pass
    If s.ProtectContents Then
      MsgBox "�p�X���[�h���Ⴂ�܂�"
      Exit Sub
    End If
  Next s
  MsgBox strFilePath & "�̃��b�N���������܂���", Buttons:=vbOKOnly + vbInformation
End Sub
