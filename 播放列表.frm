VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "播放列表:"
   ClientHeight    =   4545
   ClientLeft      =   13320
   ClientTop       =   3900
   ClientWidth     =   3705
   Icon            =   "播放列表.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "查找"
      Default         =   -1  'True
      Height          =   315
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "开始查找"
      Top             =   0
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "输入歌曲名称关键字[支持通配]"
      Top             =   0
      Width           =   2805
   End
   Begin VB.ListBox List2 
      Height          =   2760
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   4230
      ItemData        =   "播放列表.frx":000C
      Left            =   0
      List            =   "播放列表.frx":000E
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   315
      Width           =   3705
   End
   Begin VB.Menu Menu 
      Caption         =   "列表操作菜单"
      Begin VB.Menu PlayIt 
         Caption         =   "播放"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu C 
         Caption         =   "-"
      End
      Begin VB.Menu ReMoveIt 
         Caption         =   "移除曲目"
         Enabled         =   0   'False
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu DelFile 
         Caption         =   "删除文件"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu A 
         Caption         =   "-"
      End
      Begin VB.Menu GetIt 
         Caption         =   "提取到文件夹"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu OpenAdress 
         Caption         =   "打开所在文件夹"
         Enabled         =   0   'False
         Shortcut        =   ^L
      End
      Begin VB.Menu B 
         Caption         =   "-"
      End
      Begin VB.Menu Cls 
         Caption         =   "清空列表"
         Shortcut        =   ^C
      End
      Begin VB.Menu ReMyList 
         Caption         =   "刷新列表"
         Shortcut        =   ^R
      End
      Begin VB.Menu REList 
         Caption         =   "扫描音乐文件"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHFileOperation Lib "shell32" (lpFileOp As SHFILEOPSTRUCT) As Long
Const FO_COPY = &H2
Const FOF_ALLOWUNDO = &H40
Const FOF_NOCONFIRMATION = &H10
Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type
Dim xFile As SHFILEOPSTRUCT

Dim MyPath As String
Dim K, J As Long
Dim Added As Boolean
Public MP3Command As String
Dim KuoZhanMing As String

Private Sub DelFile_Click()
  On Error Resume Next
  Dim DelN As Long
  DelN = List1.ListIndex
  Kill List2.List(DelN)
  Kill Left(List2.List(DelN), Len(List2.List(DelN)) - 3) & "lrc"
  ReMoveMe DelN
  F5
End Sub

Public Sub Form_Load()
On Error Resume Next
  Call Form1.oMagneticWnd.AddWindow(Me.hwnd, Form1.hwnd)
  
  List1.Height = Me.ScaleHeight - List1.Top
  List1.Width = Me.ScaleWidth
  MyPath = Replace(App.Path & "\", "\\", "\")
  
  If Command = "" Then
    FindFiles MyPath
    Me.Caption = "播放列表：     (总数：" & List1.ListCount & ")"
  Else
    MP3Command = Command
    SplitCommand
  End If
End Sub

Public Sub SplitCommand()
  If InStr(InStr(MP3Command, ":\") + 2, MP3Command, ":\") = 0 Then
    If (GetAttr(Replace(MP3Command, Chr(34), "")) And vbDirectory) = vbDirectory Then
      FindFiles Replace(Replace(MP3Command, Chr(34), "") & "\", "\\", "\")
    Else
      Dim DataTemp As String
      DataTemp = Replace(MP3Command, Chr(34), "")
      KuoZhanMing = Right(UCase(DataTemp), 3)
      If KuoZhanMing = "MP3" Or KuoZhanMing = "WAV" Or KuoZhanMing = "WMA" Then
        If List2.List(List2.ListCount - 1) <> DataTemp Then
          List2.AddItem Replace(MP3Command, Chr(34), "")
          List1.AddItem GetFileName(Replace(MP3Command, Chr(34), ""))
        End If
        Added = True
      End If
    End If
  Else
    Dim SplitStr() As String, N As Long, X As Long, L As Long
    SplitStr = Split(MP3Command)
    For N = LBound(SplitStr) To UBound(SplitStr)
      DoEvents
      If Left(SplitStr(N), 1) = Chr(34) Then
        For L = N + 1 To UBound(SplitStr)
          SplitStr(N) = SplitStr(N) & " " & SplitStr(L)
          If Right(SplitStr(L), 1) = Chr(34) Then
            If (GetAttr(Replace(SplitStr(N), Chr(34), "")) And vbDirectory) = vbDirectory Then
              FindFiles Replace(Replace(SplitStr(N), Chr(34), "") & "\", "\\", "\")
            Else
              KuoZhanMing = Right(UCase(Replace(MP3Command, Chr(34), "")), 3)
              If KuoZhanMing = "MP3" Or KuoZhanMing = "WAV" Or KuoZhanMing = "WMA" Then
                List2.AddItem Replace(SplitStr(N), Chr(34), "")
                List1.AddItem GetFileName(Replace(SplitStr(N), Chr(34), ""))
                Added = True
              End If
            End If
            X = L
            Exit For
          End If
        Next L
        N = X
      Else
        If (GetAttr(Replace(SplitStr(N), Chr(34), "")) And vbDirectory) = vbDirectory Then
          FindFiles Replace(SplitStr(N) & "\", "\\", "\")
        Else
          KuoZhanMing = Right(UCase(Replace(MP3Command, Chr(34), "")), 3)
          If KuoZhanMing = "MP3" Or KuoZhanMing = "WAV" Or KuoZhanMing = "WMA" Then
            List2.AddItem SplitStr(N)
            List1.AddItem GetFileName(SplitStr(N))
            Added = True
          End If
        End If
      End If
    Next N
  End If
  
  PlayLast
  MP3Command = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
  List1.Height = Me.ScaleHeight - List1.Top
  List1.Width = Me.ScaleWidth
  Command1.Left = Me.ScaleWidth - Command1.Width - 60
  Text1.Width = Command1.Left - 90
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = -1
  Call Form1.oMagneticWnd.RemoveWindow(Me.hwnd)
  Me.Hide
  Form1.Check1.Val = 1
End Sub

Private Sub Command1_Click()
  Dim Done As Boolean
  
Again:
  For K = J To List1.ListCount - 1
    DoEvents
    If InStr(UCase(Trim(List1.List(K))), UCase(Text1)) <> 0 Then
      List1.ListIndex = K
      PlayIt.Enabled = True
      ReMoveIt.Enabled = True
      GetIt.Enabled = True
      DelFile.Enabled = True
      OpenAdress.Enabled = True
      J = K + 1
      List1.SetFocus
      Exit Sub
    End If
  Next
  J = -1
  If Done = True Then Exit Sub Else GoTo Again
  Done = True
End Sub

Private Sub GetIt_Click()
On Error Resume Next
  Static ToPath As String
  If ToPath = "" Or Dir(ToPath, vbDirectory) = "" Then
    Dim A As Object
    Set A = CreateObject("shell.application")
    Dim B As Object
    Set B = A.BrowseForFolder(Me.hwnd, "请设置要提取到的文件夹路径：", 0)
    ToPath = Replace(B.Self.Path & "\", "\\", "\")
  End If
  Dim ThePath(1 To 4) As String
  ThePath(1) = List2.List(List1.ListIndex)
  If Dir(ThePath(1)) <> "" Then
    ThePath(3) = ToPath & GetFileName(ThePath(1)) & ".mp3"
    xFile.pFrom = ThePath(1)
    xFile.pTo = ThePath(3)
    xFile.fFlags = &H40 Or &H10
    xFile.wFunc = FO_COPY
    xFile.hwnd = Me.hwnd
    If SHFileOperation(xFile) Then
    End If
    ThePath(2) = Left(ThePath(1), Len(ThePath(1)) - 3) & "lrc"
    If Dir(ThePath(2)) <> "" Then
      ThePath(4) = ToPath & GetFileName(ThePath(2)) & ".lrc"
      xFile.pFrom = ThePath(2)
      xFile.pTo = ThePath(4)
      xFile.fFlags = &H40 Or &H10
      xFile.wFunc = FO_COPY
      xFile.hwnd = Me.hwnd
      If SHFileOperation(xFile) Then
      End If
    End If
  End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PlayIt.Enabled = False
  ReMoveIt.Enabled = False
  GetIt.Enabled = False
  DelFile.Enabled = False
  OpenAdress.Enabled = False
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReMoveIt.Enabled = True
  If Dir(List2.List(List1.ListIndex)) <> "" Then
    PlayIt.Enabled = True
    DelFile.Enabled = True
    GetIt.Enabled = True
    OpenAdress.Enabled = True
  End If
  
  If Button = 2 Then Me.PopupMenu Menu
End Sub

Private Sub List1_DblClick()
  PlayIt_Click
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  Dim i As Long
  For i = 1 To Data.Files.Count
    If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
      FindFiles Data.Files(i) & "\"
    Else
      KuoZhanMing = Right(UCase(Data.Files(i)), 3)
      If KuoZhanMing = "MP3" Or KuoZhanMing = "WAV" Or KuoZhanMing = "WMA" Then
        List2.AddItem Data.Files(i)
        List1.AddItem GetFileName(Data.Files(i))
        Added = True
        PlayLast
      End If
    End If
  Next
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 32 Then             '空格
    PlayIt.Enabled = True
    ReMoveIt.Enabled = True
    GetIt.Enabled = True
    DelFile.Enabled = True
    OpenAdress.Enabled = True
    Me.PopupMenu Menu
  ElseIf KeyAscii = 13 Then         '回车
    If Trim(Text1.Text) = "" Then
      If List1.ListIndex = -1 Then
        Command1_Click
      Else
        PlayIt.Enabled = True
        ReMoveIt.Enabled = True
        GetIt.Enabled = True
        DelFile.Enabled = True
        OpenAdress.Enabled = True
        J = K + 1
        PlayIt_Click
      End If
    Else
      Command1_Click
    End If
  ElseIf KeyAscii = 8 Then          'BackSpace
    Dim TempN As Long
    TempN = List1.ListIndex
    ReMoveIt_Click
    Me.Caption = "播放列表：     (总数：" & List1.ListCount & ")"
    If TempN = List1.ListCount Then TempN = 0
    List1.ListIndex = TempN
  End If
End Sub


'————————————————————菜单：

Private Sub Menu_Click()
  If List1.ListIndex > -1 Then
    ReMoveIt.Enabled = True
    If Dir(List2.List(List1.ListIndex)) <> "" Then
      OpenAdress.Enabled = True
      DelFile.Enabled = True
      GetIt.Enabled = True
      PlayIt.Enabled = True
    End If
  End If
End Sub

Private Sub OpenAdress_Click()
On Error Resume Next
  Shell "explorer.exe /select," & List2.List(List1.ListIndex), vbNormalFocus
End Sub

Private Sub Cls_Click()
On Error Resume Next
  Form1.StopPlay_Click
  List1.Clear
  List2.Clear
  
  F5
End Sub

Public Sub REList_Click()
On Error Resume Next
  
  Dim Path As String
  
  Dim A As Object
  Set A = CreateObject("shell.application")
  Dim B As Object
  Set B = A.BrowseForFolder(Me.hwnd, "请选择音乐文件存放目录：", 0)
  Path = Replace(B.Self.Path & "\", "\\", "\")   '得到完整路径
  If Trim(Path) = "" Then Exit Sub
  MyPath = Path
  
  Cls_Click
  FindFiles Path
  
  F5
End Sub

Private Sub ReMoveIt_Click()
On Error Resume Next
  ReMoveMe List1.ListIndex
  F5
End Sub

Private Sub ReMyList_Click()
On Error Resume Next
  
  Cls_Click
  FindFiles MyPath

  F5
End Sub

Private Sub PlayIt_Click()
On Error Resume Next
  Form1.MP3FilePath = List2.List(List1.ListIndex)
  Form1.Music.FileName = Form1.MP3FilePath
  Form1.GetLRC Left(Form1.MP3FilePath, Len(Form1.MP3FilePath) - 3) & "lrc"
  
  Form1.Text2 = GetFileName(List2.List(List1.ListIndex))
  PlayIt.Enabled = False
  ReMoveIt.Enabled = False
  GetIt.Enabled = False
  DelFile.Enabled = False
  OpenAdress.Enabled = False
  
  Form1.MyTime = 0
  Form1.IndexN = List1.ListIndex
  Form1.PlayPause.Enabled = True
  MenuForm.PlayPause.Enabled = True
  Form1.PlayPause.GoPlay
  MenuForm.PlayPause.GoPlay
  Form1.PlayPause_Play
End Sub

'————————————————————过程及函数：
Private Sub ReMoveMe(nIndex As Long)
On Error Resume Next
  List1.RemoveItem nIndex
  List2.RemoveItem nIndex
  
  If Val(nIndex) < Form1.IndexN Then
    Form1.IndexN = Form1.IndexN - 1
  ElseIf Val(nIndex) = Form1.IndexN Then
    Form1.IndexN = Form1.IndexN - 1
    If Form1.Music.State = stPlaying Then
      If List1.ListCount > 0 Then
        Form1.PlayNext_Click
      Else
        Form1.Text2 = "小眼音乐"
        Form1.StopPlay_Click
      End If
    End If
  End If
  
  List1.ListIndex = nIndex
End Sub

Public Sub FindFiles(ThePath As String)
On Error Resume Next
  Dim Intt() As String
  Dim FilePage() As String
  Dim DirString As String
  Dim K, J, T As Integer
  ReDim Intt(0)
  Intt(0) = ThePath
  Do While K <= J
    DirString = Dir(Intt(K), vbDirectory)
  
    Do While DirString <> ""
      If DirString <> "." And DirString <> ".." Then
        If (GetAttr(Intt(K) & DirString) And vbDirectory) = vbDirectory Then
          J = J + 1
          ReDim Preserve Intt(J)
          Intt(J) = Intt(K) + DirString + "\"
        Else
          If UCase(DirString) Like "*.MP3" Or UCase(DirString) Like "*.WAV" Or UCase(DirString) Like "*.WMA" Then
            ReDim Preserve FilePage(T)
            FilePage(T) = Intt(K) + DirString
            T = T + 1
          End If
        End If
      End If
      DirString = Dir
    Loop
    K = K + 1
    
    DoEvents
  Loop
  
  If K <> 0 Then
    For K = 0 To UBound(FilePage())
      DoEvents
      List2.AddItem FilePage(K)
      List1.AddItem GetFileName(FilePage(K))
    Next K
    Added = True
  End If
End Sub

Private Sub F5()
  PlayIt.Enabled = False
  ReMoveIt.Enabled = False
  DelFile.Enabled = False
  GetIt.Enabled = False
  OpenAdress.Enabled = False
  
  Me.Caption = "播放列表：     (总数：" & List1.ListCount & ")"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub PlayLast()
  If Added = True Then
    Form1.IndexN = List1.NewIndex
    Form1.MP3FilePath = List2.List(Val(Form1.IndexN))
    Form1.GetLRC Left(Form1.MP3FilePath, Len(Form1.MP3FilePath) - 3) & "lrc"
    
    Form1.Music.FileName = Form1.MP3FilePath
    Form1.Text2.Caption = List1.List(Form2.List1.ListCount - 1)
    List1.ListIndex = List1.NewIndex
    Form1.MyTime = 0
    Form1.PlayPause.Enabled = True
    MenuForm.PlayPause.Enabled = True
    Form1.PlayPause.GoPlay
    MenuForm.PlayPause.GoPlay
    Form1.PlayPause_Play
    Me.Caption = "播放列表：     (总数：" & List1.ListCount & ")"
    Added = False
  End If
End Sub
