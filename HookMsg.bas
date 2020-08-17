Attribute VB_Name = "HookMsg"
Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

Public Const KEY_Shift = &H4
Public Const KEY_Ctrl = &H2
Public Const KEY_Alt = &H1

Public preWinProc As Long

Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If Msg = WM_HOTKEY Then
    If wParam = 1 Then
      Form1.PlayLast_Click
    ElseIf wParam = 2 Then
      Form1.PlayNext_Click
    ElseIf wParam = 3 Then
      If Form1.Music.State = stPlaying Then
        Form1.PlayPause.GoPause
        MenuForm.PlayPause.GoPause
        Form1.PlayPause_Pause
      Else
        Form1.PlayPause.GoPause
        MenuForm.PlayPause.GoPause
        Form1.PlayPause_Play
      End If
    ElseIf wParam = 4 Then
      Form1.UP
    ElseIf wParam = 5 Then
      Form1.DOWN
    ElseIf wParam = 6 Then
      MenuForm.FormVisible_Click
    End If
  ElseIf Msg = 6 Then
    If wParam = 0 Then MenuForm.Hide
  End If

  WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function
