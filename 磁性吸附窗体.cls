VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMagneticWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'甘來簾現完悶
'！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！
'！！！！銚完悶:
'Public oMagneticWnd As New cMagneticWnd

'Private Sub Form_Load()
'    Call oMagneticWnd.AddWindow(Me.hWnd)
'    Call Form2.Show(vbModeless, Me)
'End Sub
'！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！
'！！！！徨完悶:
'Private Sub Form_Load()
'    Call Form1.oMagneticWnd.AddWindow(Me.hWnd, Form1.hWnd)
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Cancel = -1
'    Call Form1.oMagneticWnd.RemoveWindow(Me.hWnd)
'End Sub

'==============================================================
'窃庁翠��Name��   cMagneticWnd

Private Enum eMsgWhen
    [MSG_AFTER] = 1
    [MSG_BEFORE] = 2
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE
End Enum
Private Const ALL_MESSAGES     As Long = -1
Private Const CODE_LEN         As Long = 197
Private Const GWL_WNDPROC      As Long = -4
Private Const PATCH_04         As Long = 88
Private Const PATCH_05         As Long = 93
Private Const PATCH_08         As Long = 132
Private Const PATCH_09         As Long = 137

Private Type tSubData
    hWnd                       As Long
    nAddrSub                   As Long
    nAddrOrig                  As Long
    nMsgCntA                   As Long
    nMsgCntB                   As Long
    aMsgTblA()                 As Long
    aMsgTblB()                 As Long
End Type

Private sc_aSubData()          As tSubData
Private sc_aBuf(1 To CODE_LEN) As Byte
Private sc_pCWP                As Long
Private sc_pEbMode             As Long
Private sc_pSWL                As Long
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Type POINTAPI
    x1 As Long
    y1 As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const SPI_GETWORKAREA  As Long = 48

Private Const WM_SIZING        As Long = &H214
Private Const WM_MOVING        As Long = &H216
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE  As Long = &H232
Private Const WM_SYSCOMMAND    As Long = &H112
Private Const WM_COMMAND       As Long = &H111

Private Const WMSZ_LEFT        As Long = 1
Private Const WMSZ_RIGHT       As Long = 2
Private Const WMSZ_TOP         As Long = 3
Private Const WMSZ_TOPLEFT     As Long = 4
Private Const WMSZ_TOPRIGHT    As Long = 5
Private Const WMSZ_BOTTOM      As Long = 6
Private Const WMSZ_BOTTOMLEFT  As Long = 7
Private Const WMSZ_BOTTOMRIGHT As Long = 8

Private Const SC_MINIMIZE      As Long = &HF020&
Private Const SC_RESTORE       As Long = &HF120&

Private Const SWP_NOSIZE       As Long = &H1
Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_NOACTIVATE   As Long = &H10

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function UnionRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type WND_INFO
    hWnd       As Long
    hWndParent As Long
    Glue       As Boolean
End Type

Private Const LB_RECT As Long = 16
Private m_uWndInfo()  As WND_INFO
Private m_lWndCount   As Long
Private m_rcWnd()     As RECT2
Private m_ptAnchor    As POINTAPI
Private m_ptOffset    As POINTAPI
Private m_ptCurr      As POINTAPI
Private m_ptLast      As POINTAPI
Private m_lSnapWidth As Long

Private Sub Class_Initialize()
    m_lSnapWidth = 10
    ReDim m_uWndInfo(0) As WND_INFO
    m_lWndCount = 0
End Sub

Private Sub Class_Terminate()
    If (m_lWndCount) Then
        Call Subclass_StopAll
    End If
End Sub

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
  Dim rcWnd As RECT2
  Dim lc    As Long
    Select Case uMsg
        Case WM_ENTERSIZEMOVE
            Call SystemParametersInfo(SPI_GETWORKAREA, 0, m_rcWnd(0), 0)
            For lc = 1 To m_lWndCount
                If (IsZoomed(m_uWndInfo(lc).hWnd)) Then
                    Call CopyMemory(m_rcWnd(lc), m_rcWnd(0), LB_RECT)
                  Else
                    Call GetWindowRect(m_uWndInfo(lc).hWnd, m_rcWnd(lc))
                End If
                If (m_uWndInfo(lc).hWnd = lng_hWnd) Then
                    Call GetCursorPos(m_ptAnchor)
                    Call GetCursorPos(m_ptLast)
                    m_ptOffset.x1 = m_rcWnd(lc).x1 - m_ptLast.x1
                    m_ptOffset.y1 = m_rcWnd(lc).y1 - m_ptLast.y1
                End If
            Next lc
        Case WM_SIZING
            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
            Call pvSizeRect(lng_hWnd, rcWnd, wParam)
            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            bHandled = True
            lReturn = 1
        Case WM_MOVING
            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
            Call pvMoveRect(lng_hWnd, rcWnd)
            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            bHandled = True
            lReturn = 1
        Case WM_EXITSIZEMOVE
            Call pvCheckGlueing
        Case WM_SYSCOMMAND
            If (wParam = SC_MINIMIZE Or wParam = SC_RESTORE) Then
                Call pvCheckGlueing
            End If
        Case WM_COMMAND
            Call pvCheckGlueing
    End Select
End Sub

Public Function AddWindow(ByVal hWnd As Long, Optional ByVal hWndParent As Long = 0) As Boolean
  Dim lc As Long
    For lc = 1 To m_lWndCount
        If (hWnd = m_uWndInfo(lc).hWnd) Then Exit Function
    Next lc
    If (IsWindow(hWnd) And (IsWindow(hWndParent) Or hWndParent = 0)) Then
        m_lWndCount = m_lWndCount + 1
        ReDim Preserve m_uWndInfo(0 To m_lWndCount)
        ReDim Preserve m_rcWnd(0 To m_lWndCount)
        With m_uWndInfo(m_lWndCount)
            .hWnd = hWnd
            .hWndParent = hWndParent
        End With
        Call pvCheckGlueing
        Call Subclass_Start(hWnd)
        Call Subclass_AddMsg(hWnd, WM_ENTERSIZEMOVE)
        Call Subclass_AddMsg(hWnd, WM_SIZING, [MSG_BEFORE])
        Call Subclass_AddMsg(hWnd, WM_MOVING, [MSG_BEFORE])
        Call Subclass_AddMsg(hWnd, WM_EXITSIZEMOVE)
        Call Subclass_AddMsg(hWnd, WM_SYSCOMMAND)
        Call Subclass_AddMsg(hWnd, WM_COMMAND)
        AddWindow = True
    End If
End Function

Public Function RemoveWindow(ByVal hWnd As Long) As Boolean
  Dim lc1 As Long
  Dim lc2 As Long
    For lc1 = 1 To m_lWndCount
        If (hWnd = m_uWndInfo(lc1).hWnd) Then
            For lc2 = lc1 To m_lWndCount - 1
                m_uWndInfo(lc2) = m_uWndInfo(lc2 + 1)
            Next lc2
            m_lWndCount = m_lWndCount - 1
            ReDim Preserve m_uWndInfo(m_lWndCount)
            ReDim Preserve m_rcWnd(m_lWndCount)
            For lc2 = 1 To m_lWndCount
                If (m_uWndInfo(lc2).hWndParent = hWnd) Then
                    m_uWndInfo(lc2).hWndParent = 0
                End If
            Next lc2
            Call Subclass_Stop(hWnd)
            Call pvCheckGlueing
            RemoveWindow = True
            Exit For
        End If
    Next lc1
End Function

Public Sub CheckGlueing()
    Call pvCheckGlueing
End Sub

Public Property Get SnapWidth() As Long
    SnapWidth = m_lSnapWidth
End Property

Public Property Let SnapWidth(ByVal New_SnapWidth As Long)
    m_lSnapWidth = New_SnapWidth
End Property

Private Sub pvSizeRect(ByVal hWnd As Long, rcWnd As RECT2, ByVal lfEdge As Long)
  Dim rcTmp As RECT2
  Dim lc    As Long
    Call CopyMemory(rcTmp, rcWnd, LB_RECT)
    For lc = 0 To m_lWndCount
        With m_rcWnd(lc)
            If (m_uWndInfo(lc).hWnd <> hWnd) Then
                If (rcWnd.y1 < .y2 + m_lSnapWidth And rcWnd.y2 > .y1 - m_lSnapWidth) Then
                    Select Case lfEdge
                      Case WMSZ_LEFT, WMSZ_TOPLEFT, WMSZ_BOTTOMLEFT
                        Select Case True
                          Case Abs(rcTmp.x1 - .x1) < m_lSnapWidth: rcWnd.x1 = .x1
                          Case Abs(rcTmp.x1 - .x2) < m_lSnapWidth: rcWnd.x1 = .x2
                        End Select
                      Case WMSZ_RIGHT, WMSZ_TOPRIGHT, WMSZ_BOTTOMRIGHT
                        Select Case True
                          Case Abs(rcTmp.x2 - .x1) < m_lSnapWidth: rcWnd.x2 = .x1
                          Case Abs(rcTmp.x2 - .x2) < m_lSnapWidth: rcWnd.x2 = .x2
                        End Select
                    End Select
                End If
                If (rcWnd.x1 < .x2 + m_lSnapWidth And rcWnd.x2 > .x1 - m_lSnapWidth) Then
                    Select Case lfEdge
                      Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        Select Case True
                          Case Abs(rcTmp.y1 - .y1) < m_lSnapWidth: rcWnd.y1 = .y1
                          Case Abs(rcTmp.y1 - .y2) < m_lSnapWidth: rcWnd.y1 = .y2
                        End Select
                      Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        Select Case True
                          Case Abs(rcTmp.y2 - .y1) < m_lSnapWidth: rcWnd.y2 = .y1
                          Case Abs(rcTmp.y2 - .y2) < m_lSnapWidth: rcWnd.y2 = .y2
                        End Select
                    End Select
                End If
            End If
        End With
    Next lc
End Sub

Private Sub pvMoveRect(ByVal hWnd As Long, rcWnd As RECT2)
  Dim lc1   As Long
  Dim lc2   As Long
  Dim lWId  As Long
  Dim rcTmp As RECT2
  Dim lOffx As Long
  Dim lOffy As Long
  Dim hDWP  As Long
    Call GetCursorPos(m_ptCurr)
    Call OffsetRect(rcWnd, (m_ptCurr.x1 - rcWnd.x1) + m_ptOffset.x1, 0)
    Call OffsetRect(rcWnd, 0, (m_ptCurr.y1 - rcWnd.y1) + m_ptOffset.y1)
    For lc1 = 0 To m_lWndCount
        If (m_uWndInfo(lc1).hWnd <> hWnd) Then
            If (m_uWndInfo(lc1).Glue = False Or m_uWndInfo(lc1).hWndParent <> hWnd) Then
                With m_rcWnd(lc1)
                    If (rcWnd.y1 < .y2 + m_lSnapWidth And rcWnd.y2 > .y1 - m_lSnapWidth) Then
                        Select Case True
                          Case Abs(rcWnd.x1 - .x1) < m_lSnapWidth: lOffx = .x1 - rcWnd.x1
                          Case Abs(rcWnd.x1 - .x2) < m_lSnapWidth: lOffx = .x2 - rcWnd.x1
                          Case Abs(rcWnd.x2 - .x1) < m_lSnapWidth: lOffx = .x1 - rcWnd.x2
                          Case Abs(rcWnd.x2 - .x2) < m_lSnapWidth: lOffx = .x2 - rcWnd.x2
                        End Select
                    End If
                    If (rcWnd.x1 < .x2 + m_lSnapWidth And rcWnd.x2 > .x1 - m_lSnapWidth) Then
                        Select Case True
                          Case Abs(rcWnd.y1 - .y1) < m_lSnapWidth: lOffy = .y1 - rcWnd.y1
                          Case Abs(rcWnd.y1 - .y2) < m_lSnapWidth: lOffy = .y2 - rcWnd.y1
                          Case Abs(rcWnd.y2 - .y1) < m_lSnapWidth: lOffy = .y1 - rcWnd.y2
                          Case Abs(rcWnd.y2 - .y2) < m_lSnapWidth: lOffy = .y2 - rcWnd.y2
                        End Select
                    End If
                End With
            End If
        End If
    Next lc1
    For lc1 = 1 To m_lWndCount
        If (m_uWndInfo(lc1).Glue And m_uWndInfo(lc1).hWndParent = hWnd) Then
            Call CopyMemory(rcTmp, m_rcWnd(lc1), LB_RECT)
            Call OffsetRect(rcTmp, m_ptCurr.x1 - m_ptAnchor.x1, 0)
            Call OffsetRect(rcTmp, 0, m_ptCurr.y1 - m_ptAnchor.y1)
            For lc2 = 0 To m_lWndCount
                If (lc1 <> lc2) Then
                    If (m_uWndInfo(lc2).Glue = False And m_uWndInfo(lc2).hWnd <> hWnd) Then
                        With m_rcWnd(lc2)
                            If (rcTmp.y1 < .y2 + m_lSnapWidth And rcTmp.y2 > .y1 - m_lSnapWidth) Then
                                Select Case True
                                  Case Abs(rcTmp.x1 - .x1) < m_lSnapWidth: lOffx = .x1 - rcTmp.x1
                                  Case Abs(rcTmp.x1 - .x2) < m_lSnapWidth: lOffx = .x2 - rcTmp.x1
                                  Case Abs(rcTmp.x2 - .x1) < m_lSnapWidth: lOffx = .x1 - rcTmp.x2
                                  Case Abs(rcTmp.x2 - .x2) < m_lSnapWidth: lOffx = .x2 - rcTmp.x2
                                End Select
                            End If
                            If (rcTmp.x1 < .x2 + m_lSnapWidth And rcTmp.x2 > .x1 - m_lSnapWidth) Then
                                Select Case True
                                  Case Abs(rcTmp.y1 - .y1) < m_lSnapWidth: lOffy = .y1 - rcTmp.y1
                                  Case Abs(rcTmp.y1 - .y2) < m_lSnapWidth: lOffy = .y2 - rcTmp.y1
                                  Case Abs(rcTmp.y2 - .y1) < m_lSnapWidth: lOffy = .y1 - rcTmp.y2
                                  Case Abs(rcTmp.y2 - .y2) < m_lSnapWidth: lOffy = .y2 - rcTmp.y2
                                End Select
                            End If
                        End With
                    End If
                End If
            Next lc2
        End If
    Next lc1
    Call OffsetRect(rcWnd, lOffx, lOffy)
    hDWP = BeginDeferWindowPos(1)
    For lc1 = 1 To m_lWndCount
        With m_uWndInfo(lc1)
            If (.hWndParent = hWnd And .Glue) Then
                lWId = pvWndGetInfoIndex(hWnd)
                With m_rcWnd(lc1)
                    Call DeferWindowPos(hDWP, m_uWndInfo(lc1).hWnd, 0, .x1 - (m_rcWnd(lWId).x1 - rcWnd.x1), .y1 - (m_rcWnd(lWId).y1 - rcWnd.y1), 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER)
                End With
            End If
        End With
    Next lc1
    Call EndDeferWindowPos(hDWP)
    m_ptLast = m_ptCurr
End Sub

Private Sub pvCheckGlueing()
  Dim lcMain As Long
  Dim lc1    As Long
  Dim lc2    As Long
  Dim lWId   As Long
    For lc1 = 1 To m_lWndCount
        Call GetWindowRect(m_uWndInfo(lc1).hWnd, m_rcWnd(lc1))
        m_uWndInfo(lc1).Glue = False
    Next lc1
    For lc1 = 1 To m_lWndCount
        If (m_uWndInfo(lc1).hWndParent) Then
            lWId = pvWndParentGetInfoIndex(m_uWndInfo(lc1).hWndParent)
            m_uWndInfo(lc1).Glue = pvWndsConnected(m_rcWnd(lWId), m_rcWnd(lc1))
        End If
    Next lc1
    For lcMain = 1 To m_lWndCount
        For lc1 = 1 To m_lWndCount
            If (m_uWndInfo(lc1).Glue) Then
                For lc2 = 1 To m_lWndCount
                    If (lc1 <> lc2) Then
                        If (m_uWndInfo(lc1).hWndParent = m_uWndInfo(lc2).hWndParent) Then
                            If (m_uWndInfo(lc2).Glue = False) Then
                                m_uWndInfo(lc2).Glue = pvWndsConnected(m_rcWnd(lc1), m_rcWnd(lc2))
                            End If
                        End If
                    End If
                Next lc2
            End If
        Next lc1
    Next lcMain
End Sub

Private Function pvWndsConnected(rcWnd1 As RECT2, rcWnd2 As RECT2) As Boolean
  Dim rcUnion As RECT2
    Call UnionRect(rcUnion, rcWnd1, rcWnd2)
    If ((rcUnion.x2 - rcUnion.x1) <= (rcWnd1.x2 - rcWnd1.x1) + (rcWnd2.x2 - rcWnd2.x1) And _
        (rcUnion.y2 - rcUnion.y1) <= (rcWnd1.y2 - rcWnd1.y1) + (rcWnd2.y2 - rcWnd2.y1) _
         ) Then
        If (rcWnd1.x1 = rcWnd2.x1 Or rcWnd1.x1 = rcWnd2.x2 Or _
            rcWnd1.x2 = rcWnd2.x1 Or rcWnd1.x2 = rcWnd2.x2 Or _
            rcWnd1.y1 = rcWnd2.y1 Or rcWnd1.y1 = rcWnd2.y2 Or _
            rcWnd1.y2 = rcWnd2.y1 Or rcWnd1.y2 = rcWnd2.y2 _
            ) Then
            pvWndsConnected = True
        End If
    End If
End Function

Private Function pvWndGetInfoIndex(ByVal hWnd As Long) As Long
  Dim lc As Long
    For lc = 1 To m_lWndCount
        If (m_uWndInfo(lc).hWnd = hWnd) Then
            pvWndGetInfoIndex = lc
            Exit For
        End If
    Next lc
End Function

Private Function pvWndParentGetInfoIndex(ByVal hWndParent As Long) As Long
  Dim lc As Long
    For lc = 1 To m_lWndCount
        If (m_uWndInfo(lc).hWnd = hWndParent) Then
            pvWndParentGetInfoIndex = lc
            Exit For
        End If
    Next lc
End Function

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
  Dim i                        As Long
  Dim J                        As Long
  Dim nSubIdx                  As Long
  Dim sSubCode                 As String
  Const GMEM_FIXED             As Long = 0
  Const PAGE_EXECUTE_READWRITE As Long = &H40&
  Const PATCH_01               As Long = 18
  Const PATCH_02               As Long = 68
  Const PATCH_03               As Long = 78
  Const PATCH_06               As Long = 116
  Const PATCH_07               As Long = 121
  Const PATCH_0A               As Long = 186
  Const FUNC_CWP               As String = "CallWindowProcA"
  Const FUNC_EBM               As String = "EbMode"
  Const FUNC_SWL               As String = "SetWindowLongA"
  Const MOD_USER               As String = "user32"
  Const MOD_VBA5               As String = "vba5"
  Const MOD_VBA6               As String = "vba6"
    If (sc_aBuf(1) = 0) Then
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
        i = 1
        Do While J < CODE_LEN
            J = J + 1
            sc_aBuf(J) = CByte("&H" & Mid$(sSubCode, i, 2))
            i = i + 2
        Loop
        If (Subclass_InIDE) Then
            sc_aBuf(16) = &H90
            sc_aBuf(17) = &H90
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (sc_pEbMode = 0) Then
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        Subclass_Start = nSubIdx
    End If
    With sc_aSubData(nSubIdx)
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)
        .hWnd = lng_hWnd
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)
    End With
End Function

Private Sub Subclass_StopAll()
  Dim i As Long
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hWnd <> 0) Then
                Call Subclass_Stop(.hWnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGE
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop
        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If
    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
      Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If
    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = 0
        If When = eMsgWhen.MSG_BEFORE Then
            nEntry = PATCH_05
          Else
            nEntry = PATCH_09
        End If
        Call zPatchVal(nAddr, nEntry, 0)
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = uMsg) Then
                aMsgTbl(nEntry) = 0
                Exit Do
            End If
        Loop
    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hWnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
