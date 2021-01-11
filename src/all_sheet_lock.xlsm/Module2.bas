Attribute VB_Name = "Module2"
Sub UnLock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excelブック,*.xlsx", Title:="シートのロックを解除するエクセルファイルを選択")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "キャンセルされました", Buttons:=vbExclamation
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass As String
  pass = InputBox("解除用のパスワードを入力", "パスワード入力", "Passw0rd")
  For Each s In wb1.Sheets
    s.UnProtect Password:=pass
    If s.ProtectContents Then
      MsgBox "パスワードが違います"
      Exit Sub
    End If
  Next s
  MsgBox strFilePath & "のロックを解除しました", Buttons:=vbOKOnly + vbInformation
End Sub
