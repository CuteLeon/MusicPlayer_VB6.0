VERSION 5.00
Begin VB.UserControl CloseButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ClipBehavior    =   0  '无
   FillStyle       =   0  'Solid
   ScaleHeight     =   300
   ScaleWidth      =   420
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   120
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   0
      Picture         =   "CloseButton.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   0
      Picture         =   "CloseButton.ctx":4C84
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "CloseButton.ctx":9872
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "CloseButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINT_API
  X As Long
  Y As Long
End Type

Public Event Click()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOut()  '鼠标经过控件
Public Event MouseIn()   '鼠标离开控件


Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  RaiseEvent MouseUp(Button, Shift, X, Y)
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
  Timer1.Enabled = True
    RaiseEvent MouseIn
    UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    RaiseEvent MouseOut
    Timer1.Enabled = False
    UserControl.Picture = Image1.Picture
  End If
End Sub

Private Sub UserControl_InitProperties()
  UserControl.Picture = Image1.Picture
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = Image1.Width
  UserControl.Height = Image1.Height
End Sub

Private Sub UserControl_Show()
  UserControl.Picture = Image1.Picture
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

