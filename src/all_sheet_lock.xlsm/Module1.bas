Attribute VB_Name = "Module1"
Sub Lock_sheet()
  Dim strFilePath As String, wb1 As Workbook
  strFilePath = Application.GetOpenFilename(Filefilter:="Excelブック,*.xlsx", Title:="シートをロックするエクセルファイルを選択")
  If strFilePath <> "False" Then
    Set wb1 = Workbooks.Open(strFilePath)
  Else
    MsgBox "キャンセルされました", Buttons:=vbExclamation
    Exit Sub
  End If
  wb1.Activate
  On Error Resume Next

  Dim s As Worksheet
  Dim pass1 As String, pass2 As String
  pass1 = InputBox("設定するパスワードを入力" & vbCrLf & "※パスワードは大文字・小文字は区別されます。", "パスワード設定", "Passw0rd")
  pass2 = InputBox("パスワードをもう一度入力してください" & vbCrLf & "注意：忘れてしまったパスワードは回復することができません。パスワードとそれに対応するブックとシートの名前を安全な場所に保管することをお勧めします。", "パスワード確認", "Passw0rd")
  If pass1 <> pass2 Then
    MsgBox "パスワードが一致しません", Buttons:=vbCritical
    Exit Sub
  End If
  For Each s In wb1.Sheets
    s.Protect Password:=pass1, _
    AllowFiltering:=True
  Next s
  MsgBox strFilePath & "をロックしました", Buttons:=vbOKOnly + vbInformation
End Sub
