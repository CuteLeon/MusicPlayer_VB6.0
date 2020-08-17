Attribute VB_Name = "PublicData"
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  TimeOut As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
Public TuoPan As PublicData.NOTIFYICONDATA

Public Function GetTime(MyTime As Long) As String
  Dim H As String
  Dim M As String
  Dim S As String
  
  If MyTime >= 3600 And MyTime <= 86400 Then
    H = Fix(MyTime / 3600)
    M = Fix((MyTime - H * 3600) / 60)
    S = MyTime - H * 3600 - M * 60
    
    If Len(H) = 1 Then H = "0" & H
    If Len(M) = 1 Then M = "0" & M
    If Len(S) = 1 Then S = "0" & S
    
    GetTime = H & ":" & M & ":" & S
  ElseIf MyTime < 3600 And MyTime >= 60 Then
    M = Fix(MyTime / 60)
    S = MyTime - M * 60
    
    If Len(M) = 1 Then M = "0" & M
    If Len(S) = 1 Then S = "0" & S
    
    GetTime = M & ":" & S
  ElseIf MyTime < 60 And MyTime >= 0 Then
    M = "00"
    
    S = MyTime
    
    If Len(S) = 1 Then S = "0" & S
    
    GetTime = M & ":" & S
  End If
End Function
