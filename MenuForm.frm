VERSION 5.00
Begin VB.Form MenuForm 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   -2040
   ClientTop       =   -2520
   ClientWidth     =   2985
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin С�����ֲ�����.XYQQButton FormVisible 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   767
      Picture         =   "MenuForm.frx":0000
      Caption         =   "���ز�����"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin С�����ֲ�����.XYQQButton ExitPlayer 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   1375
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   767
      Picture         =   "MenuForm.frx":4E85
      Caption         =   "�˳�������"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin С�����ֲ�����.FButton PlayPause 
      Height          =   900
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "����/��ͣ [Ctrl+�ո�]"
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
   End
   Begin С�����ֲ�����.SButton PlayLast 
      Height          =   750
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "��һ�� [Ctrl+Left]"
      Top             =   60
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "MenuForm.frx":9ADB
      PictureNormal   =   "MenuForm.frx":E8CD
   End
   Begin С�����ֲ�����.SButton PlayNext 
      Height          =   750
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "��һ�� [Ctrl+Right]"
      Top             =   60
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "MenuForm.frx":13737
      PictureNormal   =   "MenuForm.frx":18556
   End
   Begin С�����ֲ�����.SButton StopPlay 
      Height          =   750
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "ֹͣ����"
      Top             =   60
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "MenuForm.frx":1D3D7
      PictureNormal   =   "MenuForm.frx":21EF4
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   3000
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Sub ExitPlayer_Click()
On Error Resume Next
  Form1.ModelShown = True
  Me.Hide
  Dim Playing As Boolean
  If Form1.Music.State = stPlaying Then
    PlayPause.GoPause
    Form1.PlayPause.GoPause
    Form1.PlayPause_Pause
    Playing = True
  End If
  
  Call PlaySound("MYSOUND", App.hInstance, &H40004 + &H1 + &H2)

  If Form1.LRCCheck.Val = 2 Then
    SetWindowPos LRCForm.hwnd, 1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  End If
  
  SetWindowPos Form1.hwnd, 1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  If Form1.Check1.Val = 2 Then SetWindowPos Form2.hwnd, 1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  If MsgBox("���Ҫ�˳�ô��" & vbCrLf, vbOKCancel + vbDefaultButton2 + vbQuestion, "��Ҫ��...") = vbOK Then
    Shell_NotifyIcon &H2, PublicData.TuoPan
    CreateObject("wscript.shell").regdelete "HKEY_CLASSES_ROOT\.MP3\XYMusicCommand"
    UnregisterHotKey MenuForm.hwnd, 1
    UnregisterHotKey MenuForm.hwnd, 2
    UnregisterHotKey MenuForm.hwnd, 3
    UnregisterHotKey MenuForm.hwnd, 4
    UnregisterHotKey MenuForm.hwnd, 5
    UnregisterHotKey MenuForm.hwnd, 6
    SetWindowLong MenuForm.hwnd, GWL_WNDPROC, preWinProc
    
    If Form1.Check1.Val = 2 Then
      Form2.Hide
      Form1.Check1.Val = 1
    End If
    If Form1.LRCCheck.Val = 2 Then
      LRCForm.Hide
      Form1.LRCCheck.Val = 1
    End If
    
    SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    Form1.Timer2.Enabled = True
  Else
    Call PlaySound("MYSOUND1", App.hInstance, &H40004 + &H1 + &H2)
    If Playing = True Then
      PlayPause.GoPlay
      Form1.PlayPause.GoPlay
      Form1.PlayPause_Play
      Playing = False
    End If
    MenuForm.FormVisible.Caption = "���ز�����"
    If Form1.LRCCheck.Val = 2 Then
      SetWindowPos LRCForm.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    End If
    Form1.ModelShown = False
    SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    If Form1.Check1.Val = 2 Then SetWindowPos Form2.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  End If
End Sub

Private Sub Form_Load()
  SetWindowLong hwnd, (-20), GetWindowLong(Me.hwnd, (-20)) Or &H80000
  SetLayeredWindowAttributes hwnd, &H808080, 215, 1 Or 2
  
  preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
  SetWindowLong MenuForm.hwnd, GWL_WNDPROC, AddressOf WndProc
  
  RegisterHotKey Me.hwnd, 1, KEY_Ctrl, vbKeyLeft
  RegisterHotKey Me.hwnd, 2, KEY_Ctrl, vbKeyRight
  RegisterHotKey Me.hwnd, 3, KEY_Ctrl, vbKeySpace
  RegisterHotKey Me.hwnd, 4, KEY_Ctrl, vbKeyUp
  RegisterHotKey Me.hwnd, 5, KEY_Ctrl, vbKeyDown
  RegisterHotKey Me.hwnd, 6, KEY_Alt, vbKeyX
End Sub

Public Sub FormVisible_Click()
  If FormVisible.Caption = "���ز�����" Then
    If Form1.Check1.Val = 2 Then Form2.Hide
    PublicData.Shell_NotifyIcon &H0, PublicData.TuoPan
    PublicData.TuoPan.szInfoTitle = "С�����֣�" & Chr(0)     '����
    PublicData.TuoPan.szInfo = "    ��ã����������ϣ�   O(��_��)O����~    " & vbCrLf & vbCrLf & "           ���������һ���������...    " & Chr(0)              '����
    PublicData.TuoPan.dwInfoFlags = &H4                '����ͼ��
    PublicData.Shell_NotifyIcon &H1, PublicData.TuoPan
    Form1.Hide
    FormVisible.Caption = "��ʾ������"
  Else
    Form1.Show
    
    If Form1.Check1.Val = 2 Then
      If Screen.Width - Form1.Width - Form1.Left < Form2.Width Then
        Form2.Left = Form1.Left - Form2.Width
      Else
        Form2.Left = Form1.Left + Form1.Width
      End If
      Form2.Top = Form1.Top
      Form2.List1.Height = Form2.ScaleHeight - Form2.List1.Top
      Form2.Show , Form1
      SetWindowPos Form2.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    End If
    SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    FormVisible.Caption = "���ز�����"
  End If
  Me.Hide
End Sub

Private Sub PlayLast_Click()
  Me.Hide
  Form1.PlayLast_Click
End Sub

Private Sub PlayNext_Click()
  Me.Hide
  Form1.PlayNext_Click
End Sub

Private Sub PlayPause_Pause()
  Me.Hide
  Form1.PlayPause.GoPause
  PlayPause.GoPause
  Form1.PlayPause_Pause
End Sub

Private Sub PlayPause_Play()
  Me.Hide
  Form1.PlayPause.GoPlay
  PlayPause.GoPlay
  Form1.PlayPause_Play
End Sub

Private Sub StopPlay_Click()
  Me.Hide
  Form1.StopPlay_Click
End Sub
