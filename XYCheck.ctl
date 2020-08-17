VERSION 5.00
Begin VB.UserControl XYCheck 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   91
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   60
   End
   Begin VB.Image ImageCtl 
      Height          =   630
      Index           =   4
      Left            =   0
      Picture         =   "XYCheck.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image ImageCtl 
      Height          =   630
      Index           =   3
      Left            =   0
      Picture         =   "XYCheck.ctx":5A77
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image ImageCtl 
      Height          =   630
      Index           =   2
      Left            =   0
      Picture         =   "XYCheck.ctx":B18E
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image ImageCtl 
      Height          =   630
      Index           =   1
      Left            =   0
      Picture         =   "XYCheck.ctx":10CC1
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "XYCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINT_API
  X As Long
  Y As Long
End Type

Public Event Click()
Dim m_Val As Long

Private Sub UserControl_InitProperties()
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 8, 8)
  SetWindowRgn UserControl.hwnd, lonRect, True
  UserControl.PaintPicture ImageCtl(1).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
  m_Val = 1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture ImageCtl(m_Val * 2).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
    Timer1.Enabled = True
    UserControl.PaintPicture ImageCtl(m_Val * 2).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    UserControl.PaintPicture ImageCtl(m_Val * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
    Timer1.Enabled = False
  End If
End Sub

Private Sub UserControl_Resize()
  If m_Val = 0 Then m_Val = 1
  lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 8, 8)
  SetWindowRgn UserControl.hwnd, lonRect, True
  UserControl.PaintPicture ImageCtl(m_Val * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_Show()
  If StateData = 0 Then StateData = 1
  UserControl.PaintPicture ImageCtl(m_Val * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
End Sub

Private Sub UserControl_Click()
  If m_Val = 1 Then m_Val = 2 Else m_Val = 1
  RaiseEvent Click
  UserControl.PaintPicture ImageCtl(m_Val * 2).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Val = PropBag.ReadProperty("Val", 1)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Val", m_Val, 1)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,1
Public Property Get Val() As Long
  Val = m_Val
End Property

Public Property Let Val(ByVal New_Val As Long)
  m_Val = New_Val
  UserControl.PaintPicture ImageCtl(m_Val * 2 - 1).Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
  PropertyChanged "Val"
End Property
