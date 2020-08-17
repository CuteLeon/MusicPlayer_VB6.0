VERSION 5.00
Begin VB.UserControl FButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ClipBehavior    =   0  'нч
   FillStyle       =   0  'Solid
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3480
      Top             =   780
   End
   Begin VB.Image Image4 
      Height          =   900
      Left            =   2700
      Picture         =   "FButton.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   1800
      Picture         =   "FButton.ctx":5019
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   900
      Picture         =   "FButton.ctx":A111
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "FButton.ctx":EF43
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "FButton"
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

Public Event Play()
Public Event Pause()
Public Enabled As Boolean
Dim State As Boolean

Public Sub GoPause()
  If Enabled = False Then Exit Sub
  
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Else
    UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  End If
  State = False
End Sub

Public Sub GoPlay()
  If Enabled = False Then Exit Sub
  
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Else
    UserControl.PaintPicture Image4.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  End If
  State = True
End Sub

Private Sub UserControl_Click()
  If Enabled = False Then Exit Sub
  
  If State = True Then
    RaiseEvent Pause
    GoPause
  Else
    RaiseEvent Play
    GoPlay
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
    Timer1.Enabled = True
    If State = False Then
      UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Else
      UserControl.PaintPicture Image4.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hWnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    If State = False Then
      UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Else
      UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
    Timer1.Enabled = False
  End If
End Sub

Private Sub UserControl_InitProperties()
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 60, 60)
  SetWindowRgn UserControl.hWnd, lonRect, True
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 900
  UserControl.Height = 900
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 60, 60)
  SetWindowRgn UserControl.hWnd, lonRect, True
End Sub

Private Sub UserControl_Show()
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
