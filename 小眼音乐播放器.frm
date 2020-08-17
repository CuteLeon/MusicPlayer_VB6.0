VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5085
   ClientLeft      =   2820
   ClientTop       =   1380
   ClientWidth     =   6750
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   Icon            =   "小眼音乐播放器.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "小眼音乐播放器.frx":57E2
   ScaleHeight     =   5085
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 小眼音乐播放器.XYCheck LRCCheck 
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin 小眼音乐播放器.XYCheck Check1 
      Height          =   495
      Left            =   4860
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Val             =   2
   End
   Begin 小眼音乐播放器.PlayState PlayState 
      Height          =   750
      Left            =   5640
      TabIndex        =   11
      Top             =   4020
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
   End
   Begin VB.Timer LRCTimer 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   4080
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   5760
      Top             =   0
   End
   Begin 小眼音乐播放器.CloseButton CloseButton 
      Height          =   300
      Left            =   6340
      TabIndex        =   9
      Top             =   -10
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   529
   End
   Begin VB.TextBox TextCommand 
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   5100
      Visible         =   0   'False
      Width           =   3615
   End
   Begin 小眼音乐播放器.FButton PlayPause 
      Height          =   900
      Left            =   2580
      TabIndex        =   7
      ToolTipText     =   "播放/暂停 [Ctrl+空格]"
      Top             =   3480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
   End
   Begin 小眼音乐播放器.SButton PlayLast 
      Height          =   750
      Left            =   1740
      TabIndex        =   6
      ToolTipText     =   "上一曲 [Ctrl+Left]"
      Top             =   3600
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "小眼音乐播放器.frx":1B9EF
      PictureNormal   =   "小眼音乐播放器.frx":207E1
   End
   Begin 小眼音乐播放器.SButton PlayNext 
      Height          =   750
      Left            =   3540
      TabIndex        =   5
      ToolTipText     =   "下一曲 [Ctrl+Right]"
      Top             =   3600
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "小眼音乐播放器.frx":2564B
      PictureNormal   =   "小眼音乐播放器.frx":2A46A
   End
   Begin 小眼音乐播放器.SButton StopPlay 
      Height          =   750
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "停止播放"
      Top             =   3600
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      PictureMove     =   "小眼音乐播放器.frx":2F2EB
      PictureNormal   =   "小眼音乐播放器.frx":33E08
   End
   Begin 小眼音乐播放器.SliderBar SliderBar1 
      Height          =   300
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "调整音量 [Ctrl+Up/Down]"
      Top             =   3120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      Value           =   25
      MyMax           =   25
      MyStyle         =   8
   End
   Begin 小眼音乐播放器.ProgButton Bar 
      Height          =   375
      Left            =   480
      Top             =   2700
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Value           =   100
      MyCaption       =   "小眼音乐播放器"
      MyStyle         =   8
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   6240
      Top             =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "音量:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2130
      TabIndex        =   0
      Top             =   3140
      Width           =   525
   End
   Begin VB.Label LRCLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "小眼软件 @ 软贱你的生活"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2100
      Width           =   6735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置为默认播放器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "把小眼音乐设置为默认的播放器"
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Label Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "小眼音乐"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   2595
      TabIndex        =   1
      ToolTipText     =   "双击可打开文件"
      Top             =   1620
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'图片尺寸：450*339
Private Const WM_SETTEXT = &HC
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY_CLASSES_ROOT As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal HKEY_CLASSES_ROOT As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
 X As Long
 Y As Long
End Type
Dim scrPT As POINTAPI
'前端显示
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'鼠标事件
Const WM_MOUSEMOVE = &H200
'左键
Const WM_LBUTTONUP = &H202
'右键
Const WM_RBUTTONUP = &H205
'变量
Public Music As clsMusic
Public MyTime As Single
Dim TempPath As String
Public ModelShown As Boolean
Public IndexN
Public MP3FilePath As String
Public oMagneticWnd As New cMagneticWnd
Dim LRCTime() As Single, LRCText() As String, LRCIndex As Long

Private Sub CloseButton_Click()
  MenuForm.FormVisible_Click
End Sub

'—————————————————————窗体及控件
Private Sub Form_Load()
On Error Resume Next
  If App.PrevInstance = True Then
    Dim TohWnd As Long, DataStr As String
    DataStr = IIf(Command = "", "显示" & Timer, Command)
    TohWnd = CLng(CreateObject("Wscript.Shell").RegRead("HKEY_CLASSES_ROOT\.MP3\XYMusicCommand"))
    SendMessage TohWnd, WM_SETTEXT, 0, ByVal DataStr
    End
  Else
    CreateObject("wscript.shell").regwrite "HKEY_CLASSES_ROOT\.MP3\XYMusicCommand", TextCommand.hwnd
    If Command = "" Then CreateObject("sapi.spvoice").speak "小眼音乐"
  End If
  
  Set Music = New clsMusic
  
  App.Title = ""
  Load MenuForm
  SetWindowPos MenuForm.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    
  With PublicData.TuoPan
    .cbSize = Len(PublicData.TuoPan)
    .hwnd = Me.hwnd
    .uID = 0
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = &H2 Or &H10 Or &H1 Or &H4
    .hIcon = Me.Icon
    .szTip = "小眼音乐 @ 音乐你的生活" & vbNullChar
  End With

  PublicData.Shell_NotifyIcon &H0, PublicData.TuoPan
  
  DoEvents
  
  PublicData.TuoPan.szInfoTitle = "小眼音乐：" & Chr(0)     '标题
  PublicData.TuoPan.szInfo = vbCrLf & "     小眼音乐 @ 音乐你的生活      " & vbCrLf & "    Happy Everyday!!!   " & Chr(0)            '内容
  PublicData.TuoPan.dwInfoFlags = &H4                '气泡图标
  PublicData.Shell_NotifyIcon &H1, PublicData.TuoPan
  SetWindowLong hwnd, (-20), GetWindowLong(Form1.hwnd, (-20)) Or &H80000
  SetLayeredWindowAttributes hwnd, 0, 210, 2
  
  ReDim LRCTime(0): ReDim LRCText(0)
  
  Me.Show
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  Check1.Val = 2
  Check1_Click
End Sub

Private Sub Check1_Click()
  If Check1.Val = 2 Then
    Call oMagneticWnd.AddWindow(Me.hwnd)
    Call Form2.Show(vbModeless, Me)
    Call Form1.oMagneticWnd.AddWindow(Form2.hwnd, Form1.hwnd)
    If Screen.Width - Me.Width - Me.Left < Form2.Width Then
      Form2.Left = Me.Left - Form2.Width
    Else
      Form2.Left = Me.Left + Me.Width
    End If
    Form2.Top = Me.Top
    Form2.List1.Height = Form2.ScaleHeight - Form2.List1.Top
    Form2.Show , Form1
    Form2.Text1.SetFocus
    SetWindowPos Form2.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
    If LRCCheck.Val = 2 Then SetWindowPos LRCForm.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  Else
    Form2.Hide
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call MoveForm(hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Label3.ForeColor = vbRed Then Label3.ForeColor = vbBlack
  If Label3.FontUnderline = True Then Label3.FontUnderline = False

  Dim lMsg As Single
  If ModelShown = True Then Exit Sub
  lMsg = X / Screen.TwipsPerPixelX
  If lMsg = WM_RBUTTONUP Then
    Dim MenuLT As POINTAPI
    GetCursorPos MenuLT
    MenuForm.Move MenuLT.X * 15 - MenuForm.Width, MenuLT.Y * 15 - MenuForm.Height
    MenuForm.Show , Me
    Exit Sub
  ElseIf lMsg = WM_LBUTTONUP Then
    MenuForm.FormVisible_Click
    Exit Sub
  End If
End Sub

Private Sub Label3_Click()
  RegSetValue &H80000000, ".mp3", 1, "小眼Music", 7
  RegSetValue &H80000000, ".wav", 1, "小眼Music", 7
  RegSetValue &H80000000, ".wma", 1, "小眼Music", 7
  RegSetValue &H80000000, "小眼Music", 1, "小眼 @ Music", 9
  RegSetValue &H80000000, "小眼Music" & "\shell", 1, "Open", 5
  RegSetValue &H80000000, "小眼Music" & "\shell\Open", 1, "小眼音乐", 3
  RegSetValue &H80000000, "小眼Music" & "\shell\open\command", 1, Replace(App.Path & "\" & App.EXEName & ".exe", "\\", "\") & " %1", LenB(StrConv(Replace(App.Path & "\" & App.EXEName & ".exe", "\\", "\") & " %1", vbFromUnicode)) + 1
  RegSetValue &H80000000, "小眼Music" & "\DefaultIcon", 1, Replace(App.Path & "\" & App.EXEName & ".exe ,1", "\\", "\"), LenB(StrConv(Replace(App.Path & "\" & App.EXEName & ".exe ,1", "\\", "\"), vbFromUnicode)) + 1
  RegCloseKey &H80000000
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Label3.ForeColor <> vbRed Then Label3.ForeColor = vbRed
  If Label3.FontUnderline = False Then Label3.FontUnderline = True
End Sub

Public Sub LRCCheck_Click()
On Error Resume Next
  If LRCCheck.Val = 2 Then
    LRCForm.LRCLabel(1).Caption = LRCLabel.Caption
    If LRCIndex <= UBound(LRCText) Then LRCForm.LRCLabel(2) = LRCText(LRCIndex) Else LRCForm.LRCLabel(2) = ""
    
    If Check1.Val = 1 Then
      Check1.Val = 2
      Check1_Click
    End If
    
    DoEvents
    LRCForm.Show , Form2
    SetWindowPos LRCForm.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  Else
    LRCForm.Hide
  End If
End Sub

Private Sub LRCTimer_Timer()
On Error Resume Next
  If LRCIndex > UBound(LRCTime) Then Exit Sub
  If LRCTime(LRCIndex) <= Format(Music.Position, "####.##") Then
    LRCLabel.Caption = LRCText(LRCIndex)
    
    If LRCCheck.Val = 2 Then
      LRCForm.LRCLabel(1).Caption = LRCLabel.Caption
      If LRCIndex + 1 <= UBound(LRCText) Then LRCForm.LRCLabel(2) = LRCText(LRCIndex + 1) Else LRCForm.LRCLabel(2) = ""
    End If
    LRCIndex = LRCIndex + 1
  End If
End Sub

Private Sub SliderBar1_Change()
  Music.Volume = (SliderBar1.Value - 25) * 100
  SliderBar1.ToolTipText = "音量：" & SliderBar1.Value / SliderBar1.Max * 100 & "%"
End Sub

Private Sub Text2_DblClick()
  If Text2 <> "小眼音乐" Then Shell "explorer.exe /select," & Form2.List2.List(Form2.List1.ListIndex), vbNormalFocus
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Text2.ToolTipText = Text2.Caption
End Sub

Private Sub TextCommand_Change()
  If TextCommand = "" Then Exit Sub
  If Left(TextCommand, 2) <> "显示" Then
    Form2.MP3Command = Trim(TextCommand)
    Form2.SplitCommand
  End If
  If MenuForm.FormVisible.Caption = "显示播放器" Then MenuForm.FormVisible_Click
  TextCommand = ""
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
  Bar.Value = Music.Position
  Bar.Caption = GetTime(Fix(Music.Position)) & " / " & GetTime(Fix(Music.Duration))
  
  If Music.Position = Music.Duration Then
    If PlayState.StateData = 2 Then '单曲循环
      MyTime = 0
      PlayPause_Play
    ElseIf PlayState.StateData = 3 Then '单曲播放
      StopPlay_Click
    ElseIf PlayState.StateData = 1 Then '所有循环
      If Val(IndexN) = Form2.List1.ListCount - 1 Then
        IndexN = 0
      Else
        IndexN = Val(IndexN) + 1
      End If
      MP3FilePath = Form2.List2.List(Val(IndexN))
      Text2 = GetFileName(MP3FilePath)
      Music.FileName = MP3FilePath
      GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
      
      Bar.Max = Fix(Music.Duration) + 1
      Bar.Value = 0
      Music.Position = 0
      Music.Volume = (SliderBar1.Value - 25) * 100
      Form2.List1.ListIndex = Val(IndexN)
      Music.Play
    ElseIf PlayState.StateData = 4 Then  '随机播放
      IndexN = Rnd() * Form2.List1.ListCount - 1
      MP3FilePath = Form2.List2.List(Val(IndexN))
      Text2 = GetFileName(MP3FilePath)
      Music.FileName = MP3FilePath
      
      GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
      
      Bar.Max = Fix(Music.Duration) + 1
      Bar.Value = 0
      Music.Position = 0
      Music.Volume = (SliderBar1.Value - 25) * 100
      Form2.List1.ListIndex = Val(IndexN)
      Music.Play
    End If
  End If
End Sub

Private Sub Bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  If MP3FilePath <> "" And Music.State <> stStopped Then
    Bar.Value = Bar.Max * X / Bar.Width * 15
    Music.Position = Bar.Value
    Music.Volume = (SliderBar1.Value - 25) * 100
    Music.Play
    PlayPause.GoPlay
    MenuForm.PlayPause.GoPlay
    Timer1_Timer
    Timer1.Enabled = True
    Bar.Max = Fix(Music.Duration) + 1
    If LRCTimer.Enabled = True Then
      Dim NTemp As Long
      For NTemp = UBound(LRCTime) To LBound(LRCTime) Step -1
        If Val(LRCTime(NTemp)) < Music.Position Then
          LRCIndex = NTemp
          LRCLabel.Caption = LRCText(LRCIndex)
          If LRCCheck.Val = 2 Then
            LRCForm.LRCLabel(1).Caption = LRCLabel.Caption
            If LRCIndex + 1 <= UBound(LRCText) Then LRCForm.LRCLabel(2) = LRCText(LRCIndex + 1) Else LRCForm.LRCLabel(2) = ""
          End If
          Exit For
        End If
      Next
    End If
  End If
End Sub

Private Sub Timer2_Timer()
  Static N As Integer
  SetWindowLong hwnd, (-20), GetWindowLong(Form1.hwnd, (-20)) Or &H80000
  SetLayeredWindowAttributes Me.hwnd, vbBlue, 225 - N, 1 Or 2
  N = N + 15
  If N = 225 Then
    Unload Me
    End
  End If
End Sub

'—————————————————————操作按钮

Public Sub PlayLast_Click()
On Error Resume Next
If MP3FilePath = "" Then Exit Sub
  If Form2.List1.ListCount > 1 Then
    If Val(IndexN) = 0 Then
      IndexN = Form2.List1.ListCount - 1
    Else
      IndexN = Val(IndexN) - 1
    End If
    MP3FilePath = Form2.List2.List(Val(IndexN))
    Text2 = GetFileName(Form2.List2.List(Val(IndexN)))
    Music.FileName = MP3FilePath
    
    GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
    
    Music.Position = 0
    Bar.Max = Fix(Music.Duration) + 1
    Bar.Value = 0
    Music.Volume = (SliderBar1.Value - 25) * 100
    Form2.List1.ListIndex = Val(IndexN)
    PlayPause.GoPlay
    MenuForm.PlayPause.GoPlay
    Music.Play
    Bar.Max = Fix(Music.Duration) + 1
  End If
End Sub

Public Sub PlayNext_Click()
On Error Resume Next
If MP3FilePath = "" Then Exit Sub
  If Form2.List1.ListCount > 1 Then
    If PlayState.StateData = 4 Then  '随机播放
      IndexN = Rnd() * Form2.List1.ListCount - 1
      MP3FilePath = Form2.List2.List(Val(IndexN))
      Text2 = GetFileName(MP3FilePath)
      Music.FileName = MP3FilePath
      
      GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
      
      Bar.Max = Fix(Music.Duration) + 1
      Bar.Value = 0
      Music.Position = 0
      Music.Volume = (SliderBar1.Value - 25) * 100
      Form2.List1.ListIndex = Val(IndexN)
      Music.Play
    Else
      If Val(IndexN) = Form2.List1.ListCount - 1 Then
        IndexN = 0
      Else
        IndexN = Val(IndexN) + 1
      End If
    
      MP3FilePath = Form2.List2.List(Val(IndexN))
      Text2.Caption = GetFileName(MP3FilePath)
      Music.FileName = MP3FilePath
      
      GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
      
      Bar.Max = Fix(Music.Duration) + 1
      Bar.Value = 0
      Music.Position = 0
      Music.Volume = (SliderBar1.Value - 25) * 100
      Form2.List1.ListIndex = Val(IndexN)
      PlayPause.GoPlay
      MenuForm.PlayPause.GoPlay
      Music.Play
    End If
  End If
End Sub

Public Sub PlayPause_Pause()
On Error Resume Next
  If Trim(MP3FilePath) = "" Then Exit Sub
  Music.Pause
  MenuForm.PlayPause.GoPause
  PlayPause.GoPause
  MyTime = Music.Position
  Timer1.Enabled = False
  LRCTimer.Enabled = False
  Timer1_Timer
End Sub

Public Sub PlayPause_Play()
On Error Resume Next
  If Trim(MP3FilePath) = "" Then Exit Sub
  
  If Dir(MP3FilePath) = "" Then
    Bar.Caption = "[" & Text2 & "] 文件不存在!"
    Text2 = "小眼音乐"
    LRCLabel = "小眼软件 @ 软贱你的生活"
    If Form2.List1.ListCount > 1 Then PlayNext_Click
  Else
    If Music.State = stStopped Then
      Music.FileName = MP3FilePath
      GetLRC Left(MP3FilePath, Len(MP3FilePath) - 3) & "lrc"
    
      Text2 = GetFileName(MP3FilePath)
    End If
    Music.Position = MyTime
    Music.Volume = (SliderBar1.Value - 25) * 100
    Form2.List1.ListIndex = Val(IndexN)
    Music.Play
    MenuForm.PlayPause.GoPlay
    PlayPause.GoPlay
    Timer1_Timer
    Timer1.Enabled = True
    If UBound(LRCText) > 0 Then LRCTimer.Enabled = True
    Bar.Max = Fix(Music.Duration) + 1
  End If
End Sub

Public Sub StopPlay_Click()
On Error Resume Next
If MP3FilePath = "" Then Exit Sub
  Music.StopPlaying
  Music.Position = 0
  MenuForm.PlayPause.GoPause
  PlayPause.GoPause
  MyTime = Music.Position
  Timer1.Enabled = False
  If LRCTimer.Enabled = True Then LRCTimer.Enabled = False
  Text2 = "小眼音乐"
  Bar.Caption = "小眼音乐播放器"
  LRCLabel = "小眼软件 @ 软贱你的生活"
  Bar.Value = Bar.Max
End Sub


'—————————————————————函数
Public Sub UP()
  If Music.Volume < 0 Then
    SliderBar1.Value = SliderBar1.Value + 1
    SliderBar1.ToolTipText = "音量：" & SliderBar1.Value / SliderBar1.Max * 100 & "%"
    Music.Volume = (SliderBar1.Value - 25) * 100
  End If
End Sub

Public Sub DOWN()
  If Music.Volume > -2500 Then
    SliderBar1.Value = SliderBar1.Value - 1
    SliderBar1.ToolTipText = "音量：" & SliderBar1.Value / SliderBar1.Max * 100 & "%"
    Music.Volume = (SliderBar1.Value - 25) * 100
  End If
End Sub

Public Sub MoveForm(hwnd As Long)
  ReleaseCapture
  SendMessageA hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'获取歌词
Public Sub GetLRC(ByVal lrcPath As String)
On Error Resume Next
  LRCIndex = 0
  
  If Dir(lrcPath) <> "" Then
    If LRCCheck.Val = 2 Then
      LRCLabel = ""
      LRCForm.LRCLabel(1).Caption = LRCLabel.Caption
      If 1 <= UBound(LRCText) Then LRCForm.LRCLabel(2) = LRCText(1) Else LRCForm.LRCLabel(2) = ""
    End If
    
    Dim Data As String
    ReDim LRCTime(0): ReDim LRCText(0)
    
    Open lrcPath For Input As #1
    Do While Not EOF(1)
      Line Input #1, Data
      Data = LTrim(Data)
  
      LRCTime(UBound(LRCTime)) = Val(Format(Str(Val(Mid(Mid(Data, 2, InStr(Data, "]") - 2), 1, 2)) * 60 + Val(Right(Mid(Data, 2, InStr(Data, "]") - 2), 5))), "####.00"))
      ReDim Preserve LRCTime(UBound(LRCTime) + 1)
  
      LRCText(UBound(LRCText)) = Right(Data, Len(Data) - Val(InStr(Data, "]")))
      ReDim Preserve LRCText(UBound(LRCText) + 1)
    Loop
  
    ReDim Preserve LRCText(UBound(LRCText) - 1)
    ReDim Preserve LRCTime(UBound(LRCTime) - 1)
  
    Close #1
    
    LRCTimer.Enabled = True
  Else
    LRCTimer.Enabled = False
    LRCLabel.Caption = "O(∩_∩)O~"
    If LRCCheck.Val = 2 Then
      LRCForm.LRCLabel(1).Caption = "O(∩_∩)O~"
      LRCForm.LRCLabel(2).Caption = ""
    End If
    ReDim LRCTime(0): ReDim LRCText(0)
  End If
End Sub
