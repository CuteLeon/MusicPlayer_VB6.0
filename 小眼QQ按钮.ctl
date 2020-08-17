VERSION 5.00
Begin VB.UserControl XYQQButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ClipBehavior    =   0  '无
   FillStyle       =   0  'Solid
   ScaleHeight     =   450
   ScaleWidth      =   1290
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1980
      Top             =   240
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XYButton"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   0
      Picture         =   "小眼QQ按钮.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "小眼QQ按钮.ctx":349B
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   0
      Picture         =   "小眼QQ按钮.ctx":63E4
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "XYQQButton"
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

Private Sub Image4_Click()
  RaiseEvent Click
End Sub

Private Sub Label1_Click()
  RaiseEvent Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
  Timer1.Enabled = True
    UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
  Timer1.Enabled = True
    UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  End If
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl.PaintPicture Image3.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then Exit Sub
  
  If Timer1.Enabled = False Then
  Timer1.Enabled = True
    UserControl.PaintPicture Image2.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  End If
End Sub

Private Sub Timer1_Timer()
  Dim MousePos As POINT_API
  GetCursorPos MousePos
  If hwnd <> WindowFromPoint(MousePos.X, MousePos.Y) Then
    Timer1.Enabled = False
    UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  End If
End Sub

Private Sub UserControl_InitProperties()
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  Image4.Move 180, (UserControl.ScaleHeight - Image4.Height) / 2
  Label1.Move (UserControl.ScaleWidth + Image4.Left + Image4.Width - Label1.Width) / 2, (UserControl.ScaleHeight - Label1.Height) / 2
End Sub

Private Sub UserControl_Resize()
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
  Image4.Move 180, (UserControl.ScaleHeight - Image4.Height) / 2
  Label1.Move (UserControl.ScaleWidth + Image4.Left + Image4.Width - Label1.Width) / 2, (UserControl.ScaleHeight - Label1.Height) / 2
End Sub

Private Sub UserControl_Show()
  UserControl.PaintPicture Image1.Picture, 0, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
  UserControl.Enabled = NewEnabled
  Label1.ForeColor = IIf(NewEnabled = False, &H8000000B, vbBlack)
  PropertyChanged "Enabled"
End Property

Public Property Get Caption() As String
  Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
  Label1.Caption = NewCaption
  PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
  Label1.ForeColor = NewForeColor
  PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
  Set Font = Label1.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
  Set Label1.Font = NewFont
  PropertyChanged "Font"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
  Caption = PropBag.ReadProperty("Caption", "XYButton")
  Enabled = PropBag.ReadProperty("Enabled", True)
  ForeColor = PropBag.ReadProperty("ForeColor", Label1.ForeColor)
  Set Font = PropBag.ReadProperty("Font", Label1.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Picture", Picture, Nothing)
  Call PropBag.WriteProperty("Caption", Label1.Caption, UserControl.Name)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, UserControl.Ambient.ForeColor)
  Call PropBag.WriteProperty("Font", Label1.Font, UserControl.Ambient.Font)
End Sub

'MappingInfo=Image4,Image4,-1,Picture
Public Property Get Picture() As Picture
  Set Picture = Image4.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  Set Image4.Picture = New_Picture
  PropertyChanged "Picture"
End Property

