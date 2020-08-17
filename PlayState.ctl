VERSION 5.00
Begin VB.UserControl PlayState 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   420
      Top             =   120
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   7
      Left            =   0
      Picture         =   "PlayState.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   6
      Left            =   0
      Picture         =   "PlayState.ctx":5AE6
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   5
      Left            =   0
      Picture         =   "PlayState.ctx":B1FD
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   4
      Left            =   0
      Picture         =   "PlayState.ctx":10BA0
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   3
      Left            =   0
      Picture         =   "PlayState.ctx":1640F
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   2
      Left            =   0
      Picture         =   "PlayState.ctx":1BD4A
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   1
      Left            =   0
      Picture         =   "PlayState.ctx":215F8
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ImageCtl 
      Height          =   750
      Index           =   8
      Left            =   0
      Picture         =   "PlayState.ctx":26F63
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "PlayState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINT_API
  X As Long
  Y As Long
End Type

Public Event Change()
Public StateData As Long

Private Sub UserControl_Click()
  StateData = IIf(StateData = 4, 1, StateData + 1)
  RaiseEvent Change
  UserControl.PaintPicture ImageCtl(StateData * 2).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_InitProperties()
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 50, 50)
  SetWindowRgn UserControl.hWnd, lonRect, True
  StateData = 1
  UserControl.PaintPicture ImageCtl(1).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture ImageCtl(StateData * 2).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
    Timer1.Enabled = True
    UserControl.PaintPicture ImageCtl(StateData * 2).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    UserControl.PaintPicture ImageCtl(StateData * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    Timer1.Enabled = False
  End If
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 750
  UserControl.Height = 750
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 50, 50)
  SetWindowRgn UserControl.hWnd, lonRect, True
End Sub

Private Sub UserControl_Show()
  If StateData = 0 Then StateData = 1
  UserControl.PaintPicture ImageCtl(StateData * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
End Sub
