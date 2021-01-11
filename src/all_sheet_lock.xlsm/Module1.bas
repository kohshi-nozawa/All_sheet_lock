Attribute VB_Name = "Module1"
Sub Lock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excelブック,*.xlsx", Title:="シートをロックするエクセルファイルを選択")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "キャンセルされました"
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass1 As String, pass2 As String
  pass1 = InputBox("設定するパスワードを入力", "パスワード設定", "Passw0rd")
  pass2 = InputBox("パスワードの確認", "パスワード確認", "Passw0rd")
  If pass1 <> pass2 Then
    MsgBox "パスワードが一致しません"
    Exit Sub
  End If
  For Each s In wb1.Sheets
    s.Protect Password:=pass1,
    AllowFiltering:=True
  Next s
End Sub