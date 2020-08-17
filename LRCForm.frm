VERSION 5.00
Begin VB.Form LRCForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "小眼音乐"
   ClientHeight    =   1410
   ClientLeft      =   2505
   ClientTop       =   1425
   ClientWidth     =   14175
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LRCForm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label LRCLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "小眼音乐"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002EF4FA&
      Height          =   570
      Index           =   2
      Left            =   5970
      TabIndex        =   1
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label LRCLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "小眼软件"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   720
      Index           =   1
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   2880
   End
End
Attribute VB_Name = "LRCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_Activate()
  SetWindowLong Me.hWnd, -20, GetWindowLong(Me.hWnd, -20) Or &H8000000
  SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_Load()
  SetWindowLong hWnd, (-20), GetWindowLong(Me.hWnd, (-20)) Or &H80000
  SetLayeredWindowAttributes hWnd, Me.BackColor, 230, 1 Or 2
  
  SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = -1
  Form1.LRCCheck.Val = 1
  Me.Hide
End Sub

Private Sub LRCLabel_Change(Index As Integer)
If LRCLabel(1).Width > 14200 Or LRCLabel(2).Width > 14200 Then
  If LRCLabel(1).Width > LRCLabel(2).Width Then
    LRCLabel(1).Left = 100
    Me.Left = Me.Left + (Me.Width - (LRCLabel(1).Width + 290)) / 2
    Me.Width = LRCLabel(1).Width + 290
    LRCLabel(2).Left = (Me.ScaleWidth - LRCLabel(2).Width) / 2
  Else
    LRCLabel(2).Left = 100
    Me.Left = Me.Left + (Me.Width - (LRCLabel(2).Width + 290)) / 2
    Me.Width = LRCLabel(2).Width + 290
    LRCLabel(1).Left = (Me.ScaleWidth - LRCLabel(1).Width) / 2
  End If
Else
  If Me.Width = 14260 Then Exit Sub
  Me.Left = Me.Left - (14260 - Me.Width) / 2
  Me.Width = 14260
  LRCLabel(1).Left = (Me.ScaleWidth - LRCLabel(1).Width) / 2
  LRCLabel(2).Left = (Me.ScaleWidth - LRCLabel(2).Width) / 2
End If
End Sub
