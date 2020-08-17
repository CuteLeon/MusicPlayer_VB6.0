VERSION 5.00
Begin VB.UserControl SliderBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   DefaultCancel   =   -1  'True
   ForeColor       =   &H80000007&
   MaskColor       =   &H80000007&
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   127
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "SliderBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MyHeight      As Long                                                        '空间宽度
Dim MyWidth     As Long                                                        '控件高度
Dim MyValue      As Long                                                        '进度条的值
Dim Value_Width  As Long                                                        '计算出的进度条的宽度
Dim MyMax        As Long                                                        '最大值
Dim MyMin        As Long                                                        '最小值
Private Type RECT
    Left         As Long
    Top          As Long
    Right        As Long
    Bottom       As Long
End Type
Private Type RECT2
    Left         As Long
    Top          As Long
    Width        As Long
    Height       As Long
End Type
Private Type TRIVERTEX
    X            As Long
    Y            As Long
    Red          As Integer
    Green        As Integer
    Blue         As Integer
    Alpha        As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft    As Long
    LowerRight   As Long
End Type
'\\\\\控件事件
Public Event Change()                                                           '滚动事件
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标按下
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标弹起
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标移动
Public Event Click()                                                            '单击事件
Public Event DblClick()                                                         '双击事件
'\\\\\Api
'创建内存位图的API
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
'释放DC,对象的API
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'\\\\\绘制图片
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal dwRop As Long = vbSrcCopy) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, Optional ByVal dwRop As Long = vbSrcCopy) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'\\\\绘制虚线Api
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'\\\\获取窗口客户区坐标的API
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'\\\\绘制颜色块Api
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
'\\\\创建圆角矩形区域的API
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'\\\\设置窗口区域的API
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'\\\\创建边框的APi
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'\\\\内存赋值Api
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'\\\\设置指定矩形的内容
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'在设备场景描点的API
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim MeRect       As RECT
Dim RGBR(16)     As Long                                                        '进度条各位置的颜色
Dim MyStyle      As qStyle_Type
Dim TmRect       As Integer
Dim Focus        As Boolean
Dim FocuRect     As RECT
Dim MouseMove    As Boolean
Dim MouseDown    As Boolean
Dim MouseUp      As Boolean
Dim MouseX       As Integer

Public Enum qStyle_Type
    [橙色] = 1
    [黄色] = 2
    [浅绿] = 3
    [浅蓝] = 4
    [紫色] = 5
    [洋红] = 6
    [红色] = 7
    [蓝色] = 8
    [灰色] = 9
End Enum

Dim MoveButton   As RECT2
Dim LargeChang   As Integer
Const DuBug_Value = 0
Const DuBug_Max = 100
Const DuBug_Min = 0
Const DeBug_LargeChang = 1
Const DeBug_Style = 橙色

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyValue = PropBag.ReadProperty("Value", DuBug_Value)
    MyMin = PropBag.ReadProperty("MyMin", DuBug_Min)
    MyMax = PropBag.ReadProperty("MyMax", DuBug_Max)
    LargeChang = PropBag.ReadProperty("LargeChang", DeBug_LargeChang)
    MyStyle = PropBag.ReadProperty("MyStyle", DeBug_Style)
    MoveButton.Width = 18
    SetValue MyValue                                                            '计算滚动条位置
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Value", MyValue, DuBug_Value
    PropBag.WriteProperty "MyMin", MyMin, DuBug_Min
    PropBag.WriteProperty "MyMax", MyMax, DuBug_Max
    PropBag.WriteProperty "LargeChang", LargeChang, DeBug_LargeChang
    PropBag.WriteProperty "MyStyle", MyStyle, DeBug_Style
End Sub
                                                                         
Private Sub Timer_Timer()
    DoEvents
    If MouseDown Then
        MyValue = MyValue + LargeChang
        If MyValue < MyMin Then MyValue = MyMin
        If MyValue > MyMax Then MyValue = MyMax
        SetValue MyValue
        RaiseEvent Change
    End If
    If MouseUp Then
        MyValue = MyValue - LargeChang
        If MyValue < MyMin Then MyValue = MyMin
        If MyValue > MyMax Then MyValue = MyMax
        SetValue MyValue
        RaiseEvent Change
    End If
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_Initialize()                                            '初始化控件
    UserControl.AutoRedraw = True
    MyMax = 100
    MyMin = 0
    MyStyle = 洋红
    TmRect = 2                                                                  '设置透明范围
    LargeChang = 1
End Sub
                                                                         
Private Sub UserControl_Click()                                                 '控件单击事件
    RaiseEvent Click
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_DblClick()                                              '控件双击事件
    RaiseEvent DblClick
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Timer.Enabled = True
    '滑动块是否按下
    If X > MoveButton.Left And X < MoveButton.Left + MoveButton.Width And Y > MoveButton.Top And Y < MoveButton.Top + MoveButton.Height Then MouseMove = True: MouseX = X
    Call DrawBack
End Sub
                                                                         
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If MouseMove And Button = 1 Then
        Dim Nt As Integer
        If X > MouseX Then
            MoveButton.Left = MoveButton.Left + (X - MouseX)
        Else
            MoveButton.Left = MoveButton.Left - (MouseX - X)
        End If
        MouseX = X
        If MoveButton.Left < 2 Then MoveButton.Left = 2
        If MoveButton.Left + MoveButton.Width > UserControl.ScaleWidth - 3 Then MoveButton.Left = UserControl.ScaleWidth - MoveButton.Width - 3
        Call GetValue(MoveButton.Left)
    End If
    Call DrawBack
End Sub
                                                                         
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Timer.Enabled = False
    MouseMove = False
    If MouseDown Then
        MyValue = MyValue + LargeChang
        If MyValue < MyMin Then MyValue = MyMin
        If MyValue >= MyMax Then MyValue = MyMax
        SetValue MyValue
        RaiseEvent Change
    End If
    If MouseUp Then
        MyValue = MyValue - LargeChang
        If MyValue < MyMin Then MyValue = MyMin
        If MyValue >= MyMax Then MyValue = MyMax
        SetValue MyValue
        RaiseEvent Change
    End If
    MouseDown = False
    MouseUp = False
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Public Function GetValue(Top As Long)                                           '根据滑动条计算Value
    DoEvents
    Dim UpX As Integer, M_Value As Single
    Dim DownX As Integer, N_Value As Long
    UpX = 2
    DownX = UserControl.ScaleWidth - MoveButton.Width - 3
    M_Value = ((MoveButton.Left - UpX) / (DownX - UpX) * 100)
    M_Value = ((MyMax - MyMin) * M_Value / 100) + MyMin
    If M_Value <> Value Then
        MyValue = M_Value
        RaiseEvent Change
    End If
End Function
                                                                    
Public Function SetValue(Values As Long)                                        '根据Value计算滑动条位置
    Dim UpX As Integer, M_Value As Long
    Dim DownX As Integer, N_Value As Long
    UpX = 2
    DownX = UserControl.ScaleWidth - MoveButton.Width - 3
    M_Value = ((Values - MyMin) / (MyMax - MyMin) * 100)
    MoveButton.Left = ((DownX - UpX) * M_Value / 100) + UpX
    RaiseEvent Change
End Function
                                                                         
Private Sub UserControl_EnterFocus()                                            '获得焦点
    Focus = True
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_ExitFocus()                                             '失去焦点
    Focus = False
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_Paint()                                                 '响应控件重绘事件
    Call DrawBack                                                               '界面重绘
End Sub
                                                                         
Private Sub UserControl_Resize()                                                '响应控件重绘事件
    Dim U_LRECT As Long
    UserControl.Height = 20 * 15
    If MyWidth <> UserControl.ScaleWidth - 1 Or MyHeight <> UserControl.ScaleHeight - 1 Then '检查控件大小有没有改变
        MoveButton.Left = 2
        MoveButton.Top = 2
        MoveButton.Width = 30
        MoveButton.Height = (UserControl.Height / 15) - 5
        MyWidth = UserControl.ScaleWidth - 1                                    '取控件的宽度
        MyHeight = UserControl.ScaleHeight - 1                                '取控件高度
        U_LRECT = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, TmRect, TmRect) ''设置透明边角
        SetWindowRgn UserControl.hwnd, U_LRECT, True                            '设置透明边角
    End If
    Call DrawBack
End Sub
                                                                         
Private Sub DrawBack()
    UserControl.Cls                                                             '清除画板
    CheckColor (MyStyle)                                                        '根据效果选择颜色
    UserControl.BackColor = RGBR(0)                                             '设置背景色
    If MouseMove Then                                                           '可移动块按下  绘出 可移动块颜色
        DrawColor UserControl.hdc, MoveButton.Left + 1, MoveButton.Top + 1, MoveButton.Left + MoveButton.Width - 1, MoveButton.Top + MoveButton.Height - 1, RGBR(1), RGBR(8)
    Else
        DrawColor UserControl.hdc, MoveButton.Left + 1, MoveButton.Top + 1, MoveButton.Left + MoveButton.Width - 1, MoveButton.Top + MoveButton.Height - 1, RGBR(1), RGBR(1)
    End If
    UserControl.ForeColor = RGBR(8)                                             '设置背景颜色 = 边框颜色
    RoundRect UserControl.hdc, MoveButton.Left, MoveButton.Top, MoveButton.Left + MoveButton.Width, MoveButton.Top + MoveButton.Height, TmRect, TmRect '绘出滚动边框
    RoundRect UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, TmRect, TmRect '绘出圆角边框
    If Focus = False Then                                                       '绘制焦点虚线
    End If
End Sub
                                                                         
Private Function CheckColor(Color As Integer) As Byte()
    If Color = 1 Then
        RGBR(0) = RGB(248, 246, 242)
        RGBR(1) = RGB(233, 227, 211)                                            '背景上部颜色 结束
        RGBR(8) = RGB(204, 168, 62)                                             '边框颜色
    ElseIf Color = 2 Then
        RGBR(0) = RGB(247, 248, 242)
        RGBR(1) = RGB(231, 233, 211)
        RGBR(8) = RGB(192, 204, 62)
    ElseIf Color = 3 Then
        RGBR(0) = RGB(242, 248, 243)
        RGBR(1) = RGB(211, 233, 213)
        RGBR(8) = RGB(62, 204, 74)
    ElseIf Color = 4 Then
        RGBR(0) = RGB(242, 248, 247)
        RGBR(1) = RGB(211, 233, 231)
        RGBR(8) = RGB(62, 204, 192)
    ElseIf Color = 5 Then
        RGBR(0) = RGB(243, 242, 248)
        RGBR(1) = RGB(213, 211, 233)
        RGBR(8) = RGB(74, 62, 204)
    ElseIf Color = 6 Then
        RGBR(0) = RGB(248, 242, 247)
        RGBR(1) = RGB(233, 211, 231)
        RGBR(8) = RGB(204, 62, 192)
    ElseIf Color = 7 Then
        RGBR(0) = RGB(248, 242, 242)
        RGBR(1) = RGB(233, 211, 211)
        RGBR(8) = RGB(204, 62, 62)
    ElseIf Color = 8 Then
        RGBR(0) = RGB(250, 253, 254)
        RGBR(1) = RGB(228, 243, 252)
        RGBR(8) = RGB(23, 139, 211)
    ElseIf Color = 9 Then
        RGBR(0) = RGB(231, 243, 232)
        RGBR(1) = RGB(225, 219, 225)
        RGBR(8) = RGB(188, 184, 188)
    End If
End Function
                                                                    
'\\\区域绘制颜色
Public Sub DrawColor(ByVal hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, ByVal StartColor As Long, ByVal EndColor As Long)
    Dim PropVert(1) As TRIVERTEX, PropRECT As GRADIENT_RECT
    Dim GetRECT As RECT
    SetRect GetRECT, Left, Top, Width, Height
    With PropVert(0)
        .X = GetRECT.Left
        .Y = GetRECT.Top
        .Red = LongToShort(CLng((StartColor And &HFF&) * 256))
        .Green = LongToShort(CLng(((StartColor And &HFF00&) \ &H100&) * 256))
        .Blue = LongToShort(CLng(((StartColor And &HFF0000) \ &H10000) * 256))
        .Alpha = 0&
    End With
    With PropVert(1)
        .X = GetRECT.Right
        .Y = GetRECT.Bottom
        .Red = LongToShort(CLng((EndColor And &HFF&) * 256))
        .Green = LongToShort(CLng(((EndColor And &HFF00&) \ &H100&) * 256))
        .Blue = LongToShort(CLng(((EndColor And &HFF0000) \ &H10000) * 256))
        .Alpha = 0&
    End With
    PropRECT.UpperLeft = 1
    PropRECT.LowerRight = 0
    GradientFillRect hdc, PropVert(0), 2, PropRECT, 1, &H1
End Sub
                                                                         
Private Function LongToShort(ByVal Unsigned As Long) As Integer
    If Unsigned < 32768 Then
        LongToShort = CInt(Unsigned)
    Else
        LongToShort = CInt(Unsigned - &H10000)
    End If
End Function
                                                                                                                     
Property Get Color() As qStyle_Type                                              '获取颜色
    Color = MyStyle
End Property
                                                                    
Property Let Color(ByVal New_Style As qStyle_Type)                               '设置颜色
    MyStyle = New_Style
    Call DrawBack
End Property
                                                                    
Property Get Min() As Long                                                      '获取最小值
    Min = MyMin
End Property
                                                                    
Property Let Min(ByVal New_Min As Long)                                         '设置最小值
    MyMin = New_Min
    SetValue MyValue                                                            '计算滚动条位置
    Call DrawBack
End Property
                                                                    
Property Get Max() As Long                                                      '获取最大值
    Max = MyMax
End Property
                                                                    
Property Let Max(ByVal New_Max As Long)                                         '设置最大值
    MyMax = New_Max
    SetValue MyValue                                                            '计算滚动条位置
    Call DrawBack
End Property
                                                                    
Property Get Value() As Long                                                    '获取当前值
    Value = MyValue
End Property
                                                                    
Property Let Value(ByVal New_Value As Long)                                     '设置当前值
    If New_Value > MyMax Then New_Value = MyMax
    If New_Value < MyMin Then New_Value = MyMin
    MyValue = New_Value
    Call SetValue(MyValue)
    Call DrawBack
End Property
                                                                    
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
                                                                    
Property Get LargeChange() As Long                                              '获取当前值
    LargeChange = LargeChang
End Property
                                                                    
Property Let LargeChange(ByVal New_LargeChange As Long)                         '设置当前值
    If New_LargeChange > MyMax Then New_LargeChange = MyMax
    If New_LargeChange < 0 Then New_LargeChange = 0
    LargeChang = New_LargeChange
End Property
                                                                    
