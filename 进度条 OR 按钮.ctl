VERSION 5.00
Begin VB.UserControl ProgButton 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   DefaultCancel   =   -1  'True
   ForeColor       =   &H80000007&
   MaskColor       =   &H80000007&
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
End
Attribute VB_Name = "ProgButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'进度条的值在鼠标点击的地方
'Private Sub ProgButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  ProgButton1.Value = ProgButton1.Max * X / ProgButton1.Width * 15
'End Sub


Dim MyWidth      As Long                                                        '空间宽度
Dim MyHeight     As Long                                                        '控件高度
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
Public Enum Style_Type
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
Public Enum My_Type
    [Porgress] = 0
    [Button] = 1
End Enum
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
'\\\\获取窗口客户区坐标的API
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'\\\\绘制颜色块Api
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
'\\\\创建圆角矩形区域的API
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'\\\\设置窗口区域的API
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
'\\\\创建边框的APi
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'\\\\内存赋值Api
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'\\\\设置指定矩形的内容
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'在设备场景描点的API
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'\\\\画字用的Api
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Dim MeRect       As RECT
Dim RGBR(16)     As Long                                                        '进度条各位置的颜色
Dim MyStyle      As Style_Type
Dim TmRect       As Integer
Dim MyType       As My_Type
Dim MyImage      As StdPicture
Dim MyCaption    As String
Dim MouseDown    As Boolean
Dim MyDefault    As Boolean                                                     '是否为缺省命令按钮
'\\\\缺省变量
Const DuBug_Value = 0
Const DuBug_Max = 100
Const DuBug_Min = 0
Const DeBug_Caption = "Command"
Const DeBug_Type = Porgress
Const DeBug_Style = 橙色

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyValue = PropBag.ReadProperty("Value", DuBug_Value)
    MyMin = PropBag.ReadProperty("MyMin", DuBug_Min)
    MyMax = PropBag.ReadProperty("MyMax", DuBug_Max)
    MyCaption = PropBag.ReadProperty("MyCaption", DeBug_Caption)
    MyType = PropBag.ReadProperty("MyType", DeBug_Type)
    MyStyle = PropBag.ReadProperty("MyStyle", DeBug_Style)
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Value", MyValue, DuBug_Value
    PropBag.WriteProperty "MyMin", MyMin, DuBug_Min
    PropBag.WriteProperty "MyMax", MyMax, DuBug_Max
    PropBag.WriteProperty "MyCaption", MyCaption, DeBug_Caption
    PropBag.WriteProperty "MyType", MyType, DeBug_Type
    PropBag.WriteProperty "MyStyle", MyStyle, DeBug_Style
End Sub
                                                                         
Private Sub UserControl_Initialize()                                            '初始化控件
    UserControl.AutoRedraw = True
    MyValue = 0
    MyMin = 0
    MyMax = 100
    ' MyType = Porgress
    MyStyle = 橙色
    'UserControl.Extender.ZOrder 0
    TmRect = 5                                                                  '设置透明范围
End Sub
                                                                         
Private Sub UserControl_Click()                                                 '控件单击事件
    RaiseEvent Click
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_DblClick()                                              '控件双击事件
    RaiseEvent DblClick
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If MyType = Button Then                                                     '缺省命令按钮   KeyAscii = 13 And
        RaiseEvent Click
    End If
End Sub
                                                                         
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If MyType = Button And KeyCode = 13 Then                                    '缺省命令按钮   KeyAscii = 13 And
        RaiseEvent Click
        UserControl.Extender.SetFocus
    End If
End Sub
                                                                         
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    RaiseEvent MouseDown(Button, Shift, X, Y)
    MouseDown = True
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
                                                                         
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    RaiseEvent MouseUp(Button, Shift, X, Y)
    MouseDown = False
    UserControl_Resize                                                          '界面重绘
End Sub

Private Sub UserControl_Paint()                                                 '响应控件重绘事件
    UserControl_Resize                                                          '界面重绘
End Sub
                                                                         
Private Sub UserControl_Resize()                                                '响应控件重绘事件
    UserControl.Cls                                                             '清除画板
    Dim U_LRECT As Long
    U_LRECT = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, TmRect, TmRect) ''设置透明边角
    SetWindowRgn UserControl.hWnd, U_LRECT, True                                '设置透明边角
    MyWidth = UserControl.ScaleWidth - 1                                        '取控件的宽度
    MyHeight = UserControl.ScaleHeight - 1                                      '取控件高度
    CheckColor (MyStyle)                                                        '根据效果选择颜色
    If MyType = Porgress Then                                                   '当前为进度条样式
        DrawColor UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2, RGBR(0), RGBR(1) '绘出背景颜色上部
        DrawColor UserControl.hdc, 0, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight, RGBR(2), RGBR(3) '绘出背景颜色下部
        Dim M_Value As Integer                                                  '设置进度条宽度
        M_Value = ((MyValue - MyMin) / (MyMax - MyMin) * 100)                   '设置进度条宽度
        Value_Width = (M_Value * MyWidth / 100)                                 '设置进度条宽度
        DrawColor UserControl.hdc, 0, 0, Val(Value_Width), UserControl.ScaleHeight / 2, RGBR(4), RGBR(5) '绘出进度条颜色上部
        DrawColor UserControl.hdc, 0, UserControl.ScaleHeight / 2, Val(Value_Width), UserControl.ScaleHeight, RGBR(6), RGBR(7) '绘出进度条颜色下部
        UserControl.ForeColor = RGBR(8)                                         '设置背景颜色 = 边框颜色
        RoundRect UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, TmRect, TmRect '绘出圆角边框
        UserControl.ForeColor = &H0&                                            '设置背景颜色 = 字体颜色
        DrawTextR MyCaption
    Else                                                                        '当前为按钮样式
        DrawColor UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2, RGBR(0), RGBR(1) '绘出背景颜色上部
        DrawColor UserControl.hdc, 0, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight, RGBR(2), RGBR(3) '绘出背景颜色下部
        If MouseDown Then
            DrawColor UserControl.hdc, 0, 0, MyWidth, UserControl.ScaleHeight / 2, RGBR(4), RGBR(5) '绘出进度条颜色上部
            DrawColor UserControl.hdc, 0, UserControl.ScaleHeight / 2, MyWidth, UserControl.ScaleHeight, RGBR(6), RGBR(7) '绘出进度条颜色下部
        Else
        End If
        UserControl.ForeColor = RGBR(8)                                         '设置背景颜色 = 边框颜色
        RoundRect UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, TmRect, TmRect '绘出圆角边框
        UserControl.ForeColor = &H0&                                            '设置背景颜色 = 字体颜色
        DrawTextR MyCaption                                                     '绘出字
    End If
End Sub
                                                                         
Private Function DrawImage()
    Dim hdc As Long
    On Error Resume Next
    hdc = CreateCompatibleDC(0)                                                 '创建新的Hdc
    SelectObject hdc, MyImage.Handle                                            '设置Hdc
    'StretchBlt UserControl.hdc, 0, 0, UserControl.Width * 1.5, UserControl.Height * 1.5, hdc, 0, 0, MyImage.Width, MyImage.Height, vbSrcCopy
    BitBlt UserControl.hdc, 2, 0, UserControl.Width, UserControl.Height, hdc, 0, 0
    DeleteDC hdc
End Function
                                                                    
Private Function CheckColor(Color As Integer) As Byte()
    If Color = 1 Then
        'BACK
        RGBR(0) = RGB(248, 246, 242)                                            '背景上部颜色 起始
        RGBR(1) = RGB(233, 227, 211)                                            '背景上部颜色 结束
        '\
        RGBR(2) = RGB(226, 215, 182)                                            '背景下部颜色 起始
        RGBR(3) = RGB(239, 233, 215)                                            '背景下部颜色 结束
        'FRONT
        RGBR(4) = RGB(251, 244, 223)                                            '进度上部颜色 起始
        RGBR(5) = RGB(239, 213, 133)                                            '进度上部颜色 结束
        '\
        RGBR(6) = RGB(203, 166, 57)                                             '进度下部颜色 起始
        RGBR(7) = RGB(237, 224, 187)                                            '进度下部颜色 结束
        'FORE COLOUR
        RGBR(8) = RGB(204, 168, 62)                                             '边框颜色
    ElseIf Color = 2 Then
        'BACK
        RGBR(0) = RGB(247, 248, 242)
        RGBR(1) = RGB(231, 233, 211)
        RGBR(2) = RGB(222, 226, 182)
        RGBR(3) = RGB(237, 239, 215)
        'FRONT
        RGBR(4) = RGB(249, 251, 223)
        RGBR(5) = RGB(230, 239, 133)
        '\
        RGBR(6) = RGB(190, 203, 57)
        RGBR(7) = RGB(233, 237, 187)
        'FORE COLOUR
        RGBR(8) = RGB(192, 204, 62)
    ElseIf Color = 3 Then
        'BACK
        RGBR(0) = RGB(242, 248, 243)
        RGBR(1) = RGB(211, 233, 213)
        '\
        RGBR(2) = RGB(182, 226, 186)
        RGBR(3) = RGB(215, 239, 217)
        'FRONT
        RGBR(4) = RGB(223, 251, 225)
        RGBR(5) = RGB(133, 239, 142)
        '\
        RGBR(6) = RGB(57, 203, 70)
        RGBR(7) = RGB(187, 237, 191)
        'FORE COLOUR
        RGBR(8) = RGB(62, 204, 74)
    ElseIf Color = 4 Then
        'BACK
        RGBR(0) = RGB(242, 248, 247)
        RGBR(1) = RGB(211, 233, 231)
        '\
        RGBR(2) = RGB(182, 226, 222)
        RGBR(3) = RGB(215, 239, 237)
        'FRONT
        RGBR(4) = RGB(223, 251, 249)
        RGBR(5) = RGB(133, 239, 230)
        '\
        RGBR(6) = RGB(57, 203, 190)
        RGBR(7) = RGB(187, 237, 233)
        'FORE COLOUR
        RGBR(8) = RGB(62, 204, 192)
    ElseIf Color = 5 Then
        'BACK
        RGBR(0) = RGB(243, 242, 248)
        RGBR(1) = RGB(213, 211, 233)
        '\
        RGBR(2) = RGB(186, 182, 226)
        RGBR(3) = RGB(217, 215, 239)
        'FRONT
        RGBR(4) = RGB(225, 223, 251)
        RGBR(5) = RGB(142, 133, 239)
        '\
        RGBR(6) = RGB(70, 57, 203)
        RGBR(7) = RGB(191, 187, 237)
        'FORE COLOUR
        RGBR(8) = RGB(74, 62, 204)
    ElseIf Color = 6 Then
        'BACK
        RGBR(0) = RGB(248, 242, 247)
        RGBR(1) = RGB(233, 211, 231)
        '\
        RGBR(2) = RGB(226, 182, 222)
        RGBR(3) = RGB(239, 215, 237)
        'FRONT
        RGBR(4) = RGB(251, 223, 249)
        RGBR(5) = RGB(239, 133, 230)
        '\
        RGBR(6) = RGB(203, 57, 190)
        RGBR(7) = RGB(237, 187, 233)
        'FORE COLOUR
        RGBR(8) = RGB(204, 62, 192)
    ElseIf Color = 7 Then
        'BACK
        RGBR(0) = RGB(248, 242, 242)
        RGBR(1) = RGB(233, 211, 211)
        '\
        RGBR(2) = RGB(226, 182, 182)
        RGBR(3) = RGB(239, 215, 215)
        'FRONT
        RGBR(4) = RGB(251, 223, 223)
        RGBR(5) = RGB(239, 133, 133)
        '\
        RGBR(6) = RGB(203, 57, 57)
        RGBR(7) = RGB(237, 187, 187)
        'FORE COLOUR
        RGBR(8) = RGB(204, 62, 62)
    ElseIf Color = 8 Then
        'BACK
        RGBR(0) = RGB(250, 253, 254)
        RGBR(1) = RGB(228, 243, 252)
        '\
        RGBR(2) = RGB(199, 230, 249)
        RGBR(3) = RGB(237, 247, 253)
        'FRONT
        RGBR(4) = RGB(225, 247, 255)
        RGBR(5) = RGB(67, 208, 255)
        '\
        RGBR(6) = RGB(63, 112, 233)
        RGBR(7) = RGB(63, 226, 246)
        'FORE COLOUR
        RGBR(8) = RGB(23, 139, 211)
    ElseIf Color = 9 Then
        'BACK
        RGBR(0) = RGB(231, 243, 232)
        RGBR(1) = RGB(225, 219, 225)
        '\
        RGBR(2) = RGB(179, 189, 179)
        RGBR(3) = RGB(226, 238, 226)
        'FRONT
        RGBR(4) = RGB(223, 251, 223)
        RGBR(5) = RGB(108, 255, 108)
        '\
        RGBR(6) = RGB(26, 228, 26)
        RGBR(7) = RGB(217, 244, 217)
        'FORE COLOUR
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
    PropRECT.UpperLeft = 0
    PropRECT.LowerRight = 1
    GradientFillRect hdc, PropVert(0), 2, PropRECT, 1, &H1
End Sub
                                                                         
Private Function LongToShort(ByVal Unsigned As Long) As Integer
    If Unsigned < 32768 Then
        LongToShort = CInt(Unsigned)
    Else
        LongToShort = CInt(Unsigned - &H10000)
    End If
End Function
                                                                    
Public Function DrawTextR(Text As String)
    MeRect.Left = (MyWidth / 2) - GetFontWidth(Text) / 2                        '设置画字左坐标
    MeRect.Top = (MyHeight / 2) - GetFontHeight(Text) / 2                       '设置画字上坐标
    MeRect.Right = MeRect.Left + GetFontWidth(Text)                             '设置画字右坐标
    MeRect.Bottom = MeRect.Top + GetFontHeight(Text)                            '设置画字下坐标
    DrawText UserControl.hdc, Text, -1, MeRect, dwFlag                          '绘出字体
End Function
                                                                    
Public Function GetFontWidth(Tmp As String) As Long                             '取字符宽度
    GetFontWidth = UserControl.TextWidth(Tmp)
End Function
                                                                    
Public Function GetFontHeight(Tmp As String) As Long                            '取字符高度
    GetFontHeight = UserControl.TextHeight(Tmp)
End Function
                                                                    
Property Get Color() As Style_Type                                              '获取颜色
    Color = MyStyle
End Property
                                                                    
Property Let Color(ByVal New_Style As Style_Type)                               '设置颜色
    MyStyle = New_Style
    Call UserControl_Resize
End Property
                                                                    
Property Get Min() As Long                                                      '获取最小值
    Min = MyMin
End Property
                                                                    
Property Let Min(ByVal New_Min As Long)                                         '设置最小值
    MyMin = New_Min
    Call UserControl_Resize
End Property
                                                                    
Property Get Max() As Long                                                      '获取最大值
    Max = MyMax
End Property
                                                                    
Property Let Max(ByVal New_Max As Long)                                         '设置最大值
    MyMax = New_Max
    Call UserControl_Resize
End Property
                                                                    
Property Get Style() As My_Type                                                 '获取样式 进度条 还是按钮
    Style = MyType
End Property
                                                                    
Property Let Style(ByVal New_Type As My_Type)                                   '设置样式 进度条 还是按钮
    MyType = New_Type
    Call UserControl_Resize
End Property
                                                                    
Property Get Caption() As String                                                '获取控件标题
    Caption = MyCaption
End Property
                                                                    
Property Let Caption(ByVal New_Caption As String)                               '设置控件标题
    Dim i As Integer, J As Integer
    i = InStrRev(New_Caption, "&")
    Do While i
        If Mid$(New_Caption, i, 2) = "&&" Then
            i = InStrRev(i - 1, New_Caption, "&")
        Else
            J = i + 1: i = 0
        End If
    Loop
    If J Then AccessKeys = Mid$(New_Caption, J, 1)
    MyCaption = New_Caption
    Call UserControl_Resize
End Property
                                                                    
Property Get Value() As Long                                                    '获取当前值
    Value = MyValue
End Property
                                                                    
Property Let Value(ByVal New_Value As Long)                                     '设置当前值
    If New_Value > MyMax Then New_Value = MyMax
    If New_Value < MyMin Then New_Value = MyMin
    MyValue = New_Value
    UserControl_Resize
End Property
                                                                    
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
                                                                    
Public Property Get Picture() As StdPicture
    Set Picture = MyImage
End Property
                                                                    
Public Property Set Picture(xPic As StdPicture)
    Set MyImage = xPic
    Call UserControl_Resize
End Property
                                                                    
