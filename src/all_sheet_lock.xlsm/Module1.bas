Attribute VB_Name = "Module1"
Sub Lock_sheet()
  Dim strFilePath As String
  strFilePath = Application.GetOpenFilename(Filefilter:="Excelブック,*.xlsx", Title:="シートをロックするエクセルファイルを選択")
  
End Sub

