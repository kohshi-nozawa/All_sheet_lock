Attribute VB_Name = "Module1"
Sub Lock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excel�u�b�N,*.xlsx", Title:="�V�[�g�����b�N����G�N�Z���t�@�C����I��")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "�L�����Z������܂���", Buttons:=vbExclamation
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass1 As String, pass2 As String
  pass1 = InputBox("�ݒ肷��p�X���[�h�����" & vbCrLf & "���p�X���[�h�͑啶���E�������͋�ʂ���܂��B", "�p�X���[�h�ݒ�", "Passw0rd")
  pass2 = InputBox("�p�X���[�h��������x���͂��Ă�������" & vbCrLf & "���ӁF�Y��Ă��܂����p�X���[�h�͉񕜂��邱�Ƃ��ł��܂���B�p�X���[�h�Ƃ���ɑΉ�����u�b�N�ƃV�[�g�̖��O�����S�ȏꏊ�ɕۊǂ��邱�Ƃ������߂��܂��B", "�p�X���[�h�m�F", "Passw0rd")
  If pass1 <> pass2 Then
    MsgBox "�p�X���[�h����v���܂���", Buttons:=vbCritical
    Exit Sub
  End If
  For Each s In wb1.Sheets
    s.Protect Password:=pass1, _
    AllowFiltering:=True
  Next s
  MsgBox strFilePath & "�����b�N���܂���", Buttons:=vbOKOnly + vbInformation
End Sub
