Attribute VB_Name = "DialogCTL"
Public Function GetFileName(ByVal FilePath As String) As String
On Error Resume Next
  Dim temp As Long
  Dim TempName As String
  temp = InStrRev(FilePath, "\")
  TempName = Mid(FilePath, temp + 1)
  GetFileName = Left(TempName, Len(TempName) - 4)
End Function
