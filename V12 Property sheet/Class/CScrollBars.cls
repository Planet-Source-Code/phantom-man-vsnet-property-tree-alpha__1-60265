VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_ScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'TreeProps
'//---------------------------------------------------------------------------------------
'-- Module    : IAPP_ScrollBars
'-- DateTime  : 26/07/2004 13:45
'-- Author    : Gary Noble
'-- Purpose   : Implements Api ScrollBars
'-- Assumes   :
'//---------------------------------------------------------------------------------------
'-- History   : Initial Implementation    Gary Noble  26/07/2004
'//---------------------------------------------------------------------------------------
Option Explicit

Private Const WM_MOUSEMOVE As Integer = &H200

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128                      '--  Maintenance string for PSS usage
End Type
Private Type SCROLLINFO
    cbSize     As Long
    fMask      As Long
    nMin       As Long
    nMax       As Long
    nPage      As Long
    nPos       As Long
    nTrackPos  As Long
End Type
Private Const SB_BOTH As Integer = 3
Private Const SB_BOTTOM As Integer = 7
Private Const SB_CTL As Integer = 2
Private Const SB_ENDSCROLL As Integer = 8
Private Const SB_HORZ As Integer = 0
Private Const SB_LEFT As Integer = 6
Private Const SB_LINEDOWN As Integer = 1
Private Const SB_LINELEFT As Integer = 0
Private Const SB_LINERIGHT As Integer = 1
Private Const SB_LINEUP As Integer = 0
Private Const SB_PAGEDOWN As Integer = 3
Private Const SB_PAGELEFT As Integer = 2
Private Const SB_PAGERIGHT As Integer = 3
Private Const SB_PAGEUP As Integer = 2
Private Const SB_RIGHT As Integer = 7
Private Const SB_THUMBPOSITION As Integer = 4
Private Const SB_THUMBTRACK As Integer = 5
Private Const SB_TOP As Integer = 6
Private Const SB_VERT As Integer = 1
Private Const SIF_RANGE As Long = &H1
Private Const SIF_PAGE As Long = &H2
Private Const SIF_POS As Long = &H4
Private Const SIF_DISABLENOSCROLL As Long = &H8
Private Const SIF_TRACKPOS As Long = &H10
Private Const SIF_ALL As Double = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const ESB_DISABLE_BOTH As Long = &H3
Private Const ESB_ENABLE_BOTH As Long = &H0
Private Const SBS_SIZEGRIP As Long = &H10&
'-- Non-client messages:
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const WM_NCRBUTTONDOWN As Long = &HA4
Private Const WM_NCMBUTTONDOWN As Long = &HA7
'-- Hit test codes for scroll bars:
Private Const HTHSCROLL As Integer = 6
Private Const HTVSCROLL As Integer = 7
'-- Scroll bar messages:
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const WM_MOUSEWHEEL As Long = &H20A
'-- Mouse wheel stuff:
Private Const WHEEL_DELTA As Integer = 120
Private Const WHEEL_PAGESCROLL As Integer = -1
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const MSH_MOUSEWHEEL As String = "MSWHEEL_ROLLMSG"
Private Const MSH_WHEELMODULE_CLASS As String = "MouseZ"
Private Const MSH_WHEELMODULE_TITLE As String = "Magellan MSWHEEL"
Private Const MSH_WHEELSUPPORT As String = "MSH_WHEELSUPPORT_MSG"
Private Const MSH_SCROLL_LINES As String = "MSH_SCROLL_LINES_MSG"
'-- Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL As Long = &H1&
Private Const WSB_PROP_CXHSCROLL As Long = &H2&
Private Const WSB_PROP_CYHSCROLL As Long = &H4&
Private Const WSB_PROP_CXVSCROLL As Long = &H8&
Private Const WSB_PROP_CXHTHUMB As Long = &H10&
Private Const WSB_PROP_CYVTHUMB As Long = &H20&
Private Const WSB_PROP_VBKGCOLOR As Long = &H40&
Private Const WSB_PROP_HBKGCOLOR As Long = &H80&
Private Const WSB_PROP_VSTYLE As Long = &H100&
Private Const WSB_PROP_HSTYLE As Long = &H200&
Private Const WSB_PROP_WINSTYLE As Long = &H400&
Private Const WSB_PROP_PALETTE As Long = &H800&
Private Const WSB_PROP_MASK As Long = &HFFF&
Private Const FSB_FLAT_MODE As Long = 2
Private Const FSB_ENCARTA_MODE As Long = 1
Private Const FSB_REGULAR_MODE As Long = 0
'-- Message response:
Implements ISubclass
Private m_emr  As EMsgResponse
'-- Initialisation state:
Private m_bInitialised As Boolean
'-- Orientation
Public Enum EFSOrientationConstants
    efsoHorizontal
    efsoVertical
    efsoBoth
End Enum
#If False Then
    Private efsoHorizontal, efsoVertical, efsoBoth
#End If
Private m_eOrientation As EFSOrientationConstants
'-- Style
Public Enum EFSStyleConstants
    efsRegular = FSB_REGULAR_MODE
    efsEncarta = FSB_ENCARTA_MODE
    efsFlat = FSB_FLAT_MODE
End Enum
#If False Then
    Private efsRegular, efsEncarta, efsFlat
#End If
Private m_eStyle As EFSStyleConstants
'-- Bars:
Public Enum EFSScrollBarConstants
    efsHorizontal = SB_HORZ
    efsVertical = SB_VERT
End Enum
#If False Then
    Private efsHorizontal, efsVertical
#End If
'-- Can we have flat scroll bars?
Private m_bNoFlatScrollBars As Boolean
'-- hWnd we're adding scroll bars too:
Private m_hWnd As Long
'-- Small change amount
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
'-- Enabled:
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
'-- Visible
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean
'-- Number of lines to scroll for each wheel click:
Private m_lWheelScrollLines As Long
Public Event ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
Public Event Scroll(eBar As EFSScrollBarConstants)
Public Event Change(eBar As EFSScrollBarConstants)
Public Event MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
Public Event MouseMove(ByVal X As Long, ByVal Y As Long)
Public Event AllowScroll(bAllow As Boolean)

Private mvarSetFlat As Boolean
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Long, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal wBar As Long, _
                                                     ByVal bShow As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal n As Long, _
                                                     lpcScrollInfo As SCROLLINFO, _
                                                     ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal n As Long, _
                                                     lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nBar As Long, _
                                                      lpMinPos As Long, _
                                                      lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal nBar As Long, _
                                                    ByVal nPos As Long, _
                                                    ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nBar As Long, _
                                                      ByVal nMinPos As Long, _
                                                      ByVal nMaxPos As Long, _
                                                      ByVal bRedraw As Long) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, _
                                                       ByVal wSBflags As Long, _
                                                       ByVal wArrows As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                    ByVal int2 As Long, _
                                                                    ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                  ByVal code As Long, _
                                                                  ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollRange Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                   ByVal code As Long, _
                                                                   ByVal LPINT1 As Long, _
                                                                   ByVal LPINT2 As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                  ByVal code As Long, _
                                                                  lpScrollInfo As SCROLLINFO) As Long
Private Declare Function FlatSB_GetScrollPos Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                 ByVal code As Long) As Long
Private Declare Function FlatSB_GetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                  ByVal propIndex As Long, _
                                                                  ByVal LPINT As Long) As Long
Private Declare Function FlatSB_SetScrollPos Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                 ByVal code As Long, _
                                                                 ByVal Pos As Long, _
                                                                 ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                  ByVal code As Long, _
                                                                  lpScrollInfo As SCROLLINFO, _
                                                                  ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                   ByVal code As Long, _
                                                                   ByVal Min As Long, _
                                                                   ByVal Max As Long, _
                                                                   ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, _
                                                                  ByVal Index As Long, _
                                                                  ByVal newValue As Long, _
                                                                  ByVal fRedraw As Boolean) As Long
Private Declare Function InitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

Public Property Get CanBeFlat() As Boolean

    CanBeFlat = Not (m_bNoFlatScrollBars)

End Property

Private Sub Class_Initialize()

    m_lSmallChangeHorz = 1
    m_lSmallChangeVert = 1
    m_eStyle = efsRegular
    m_eOrientation = efsoBoth

End Sub

Private Sub Class_Terminate()

    pClearUp

End Sub

Public Sub Create(ByVal hWndA As Long)

    pClearUp
    m_hWnd = hWndA
    pCreateScrollBar
    pAttachMessages

End Sub

Public Property Get Enabled(ByVal eBar As EFSScrollBarConstants) As Boolean

    If eBar = efsHorizontal Then
        Enabled = m_bEnabledHorz
    Else                                              'NOT EBAR...
        Enabled = m_bEnabledVert
    End If

End Property

Public Property Let Enabled(ByVal eBar As EFSScrollBarConstants, _
                            ByVal bEnabled As Boolean)

    Dim lO     As Long
    Dim lF     As Long

    lO = eBar
    If bEnabled Then
        lF = ESB_ENABLE_BOTH
    Else                                              'BENABLED = FALSE/0
        lF = ESB_DISABLE_BOTH
    End If
    If m_bNoFlatScrollBars Then
        EnableScrollBar m_hWnd, lO, lF
    Else                                              'M_BNOFLATSCROLLBARS = FALSE/0
        FlatSB_EnableScrollBar m_hWnd, lO, lF
    End If

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

    ISubclass_MsgResponse = emrPostProcess

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

'

End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

    Dim lScrollCode As Long
    Dim tSI    As SCROLLINFO
    Dim lV     As Long
    Dim lSC    As Long
    Dim eBar   As EFSScrollBarConstants
    Dim zDelta As Long
    Dim lDelta As Long
    Dim pt     As POINTAPI
    Dim bAllowScroll As Boolean

    ''Dim wMKeyFlags  As Long
    Dim eBtn   As MouseButtonConstants

    Select Case iMsg
        Case WM_MOUSEMOVE
            GetCursorPos pt
            ScreenToClient m_hWnd, pt
            RaiseEvent MouseMove(pt.X, pt.Y)
        Case WM_MOUSEWHEEL

            RaiseEvent AllowScroll(bAllowScroll)
            If Not bAllowScroll Then Exit Function

            '-- Low-word of wParam indicates whether virtual keys
            '-- are down
            'wMKeyFlags = wParam And &HFFFF&
            '-- High order word is the distance the wheel has been rotated,
            '-- in multiples of WHEEL_DELTA:
            If (wParam And &H8000000) Then
                '-- Towards the user:
                zDelta = &H8000& - (wParam And &H7FFF0000) \ &H10000
            Else                                      'NOT (WPARAM...
                '-- Away from the user:
                zDelta = -((wParam And &H7FFF0000) \ &H10000)
            End If
            lDelta = (zDelta \ WHEEL_DELTA) * SmallChange(efsVertical) * m_lWheelScrollLines
            eBar = efsVertical
            RaiseEvent MouseWheel(eBar, lDelta)
            If Not (lDelta = 0) Then
                Value(eBar) = Value(eBar) + lDelta
                ISubclass_WindowProc = 1
            End If
        Case WM_VSCROLL, WM_HSCROLL

            RaiseEvent AllowScroll(bAllowScroll)
            If Not bAllowScroll Then Exit Function

            If iMsg = WM_HSCROLL Then
                eBar = efsHorizontal
            Else                                      'NOT IMSG...
                eBar = efsVertical
            End If
            lScrollCode = (wParam And &HFFFF&)
            Select Case lScrollCode
                Case SB_THUMBTRACK
                    '-- Is vertical/horizontal?
                    pGetSI eBar, tSI, SIF_TRACKPOS
                    Value(eBar) = tSI.nTrackPos
                    pRaiseEvent eBar, True
                Case SB_LEFT, SB_TOP
                    Value(eBar) = Min(eBar)
                    pRaiseEvent eBar, False
                Case SB_RIGHT, SB_BOTTOM
                    Value(eBar) = Max(eBar)
                    pRaiseEvent eBar, False
                Case SB_LINELEFT, SB_LINEUP
                    'Debug.Print "Line"
                    lV = Value(eBar)
                    If eBar = efsHorizontal Then
                        lSC = m_lSmallChangeHorz
                    Else                              'NOT EBAR...
                        lSC = m_lSmallChangeVert
                    End If
                    If (lV - lSC < Min(eBar)) Then
                        Value(eBar) = Min(eBar)
                    Else                              'NOT (LV...
                        Value(eBar) = lV - lSC
                    End If
                    pRaiseEvent eBar, False
                Case SB_LINERIGHT, SB_LINEDOWN
                    'Debug.Print "Line"
                    lV = Value(eBar)
                    If eBar = efsHorizontal Then
                        lSC = m_lSmallChangeHorz
                    Else                              'NOT EBAR...
                        lSC = m_lSmallChangeVert
                    End If
                    If (lV + lSC > Max(eBar)) Then
                        Value(eBar) = Max(eBar)
                    Else                              'NOT (LV...
                        Value(eBar) = lV + lSC
                    End If
                    pRaiseEvent eBar, False
                Case SB_PAGELEFT, SB_PAGEUP
                    Value(eBar) = Value(eBar) - LargeChange(eBar)
                    pRaiseEvent eBar, False
                Case SB_PAGERIGHT, SB_PAGEDOWN
                    Value(eBar) = Value(eBar) + LargeChange(eBar)
                    pRaiseEvent eBar, False
                Case SB_ENDSCROLL
                    pRaiseEvent eBar, False
            End Select
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN
            eBtn = IIf(iMsg = WM_NCLBUTTONDOWN, vbLeftButton, vbRightButton)
            If wParam = HTVSCROLL Then
                RaiseEvent ScrollClick(efsHorizontal, eBtn)
            ElseIf wParam = HTHSCROLL Then            'NOT WPARAM...
                RaiseEvent ScrollClick(efsVertical, eBtn)
            End If
    End Select

End Function

Public Property Get LargeChange(ByVal eBar As EFSScrollBarConstants) As Long

    Dim tSI    As SCROLLINFO

    pGetSI eBar, tSI, SIF_PAGE
    LargeChange = tSI.nPage

End Property

Public Property Let LargeChange(ByVal eBar As EFSScrollBarConstants, _
                                ByVal iLargeChange As Long)

    Dim tSI    As SCROLLINFO

    pGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE

End Property

Public Property Get Max(ByVal eBar As EFSScrollBarConstants) As Long

    Dim tSI    As SCROLLINFO

    pGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage

End Property

Public Property Let Max(ByVal eBar As EFSScrollBarConstants, _
                        ByVal iMax As Long)

    Dim tSI    As SCROLLINFO

    tSI.nMax = iMax + LargeChange(eBar)
    tSI.nMin = Min(eBar)
    pLetSI eBar, tSI, SIF_RANGE
    pRaiseEvent eBar, False

End Property

Public Property Get Min(ByVal eBar As EFSScrollBarConstants) As Long

    Dim tSI    As SCROLLINFO

    pGetSI eBar, tSI, SIF_RANGE
    Min = tSI.nMin

End Property

Public Property Let Min(ByVal eBar As EFSScrollBarConstants, _
                        ByVal iMin As Long)

    Dim tSI    As SCROLLINFO

    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar)
    pLetSI eBar, tSI, SIF_RANGE

End Property

Public Property Get Orientation() As EFSOrientationConstants

    Orientation = m_eOrientation

End Property

Public Property Let Orientation(ByVal eOrientation As EFSOrientationConstants)

    m_eOrientation = eOrientation
    pSetOrientation

End Property

Private Sub pAttachMessages()

    If m_hWnd <> 0 Then
        AttachMessage Me, m_hWnd, WM_HSCROLL
        AttachMessage Me, m_hWnd, WM_VSCROLL
        AttachMessage Me, m_hWnd, WM_MOUSEWHEEL
        AttachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
        AttachMessage Me, m_hWnd, WM_NCMBUTTONDOWN
        AttachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
        AttachMessage Me, m_hWnd, WM_MOUSEMOVE

        SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, m_lWheelScrollLines, 0
        If m_lWheelScrollLines <= 0 Then
            m_lWheelScrollLines = 3
        End If
        m_bInitialised = True
    End If

End Sub

Private Sub pClearUp()

    If m_hWnd <> 0 Then
        On Error Resume Next
        '-- Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If
        On Error GoTo 0
        '-- Remove subclass:
        DetachMessage Me, m_hWnd, WM_HSCROLL
        DetachMessage Me, m_hWnd, WM_VSCROLL
        DetachMessage Me, m_hWnd, WM_MOUSEWHEEL
        DetachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_NCMBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_MOUSEMOVE
    End If
    m_hWnd = 0
    m_bInitialised = False

End Sub

Public Property Get ScrollHwnd() As Long

    ScrollHwnd = m_hWnd

End Property

Private Sub pCreateScrollBar()

'Dim lR As Long
'-- Just checks for flag scroll bars...

    On Error Resume Next

    If Me.Style <> efsRegular Then
        InitialiseFlatSB m_hWnd
    Else                                              'NOT ME.STYLE...
        ShowScrollBar m_hWnd, 1, True
    End If
    If Err.Number <> 0 Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        '-- Means we have version prior to 4.71
        '-- We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else                                              'NOT ERR.NUMBER...
        Style = m_eStyle
    End If
    On Error GoTo 0

End Sub

Private Sub pGetSI(ByVal eBar As EFSScrollBarConstants, _
                   ByRef tSI As SCROLLINFO, _
                   ByVal fMask As Long)

    Dim lO     As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    If m_bNoFlatScrollBars Then
        GetScrollInfo m_hWnd, lO, tSI
    Else                                              'M_BNOFLATSCROLLBARS = FALSE/0
        FlatSB_GetScrollInfo m_hWnd, lO, tSI
    End If

End Sub

Private Sub pLetSI(ByVal eBar As EFSScrollBarConstants, _
                   ByRef tSI As SCROLLINFO, _
                   ByVal fMask As Long)

    Dim lO     As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    If m_bNoFlatScrollBars Then
        SetScrollInfo m_hWnd, lO, tSI, True
    Else                                              'M_BNOFLATSCROLLBARS = FALSE/0
        FlatSB_SetScrollInfo m_hWnd, lO, tSI, True
    End If

End Sub

Private Sub pRaiseEvent(ByVal eBar As EFSScrollBarConstants, _
                        ByVal bScroll As Boolean)

    Static s_lLastValue(0 To 1) As Long

    If Value(eBar) <> s_lLastValue(eBar) Then
        If bScroll Then
            RaiseEvent Scroll(eBar)
        Else                                          'BSCROLL = FALSE/0
            RaiseEvent Change(eBar)
        End If
        s_lLastValue(eBar) = Value(eBar)
    End If

End Sub

Private Sub pSetOrientation()

    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoHorizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoVertical))

End Sub

Public Property Get SetFlat() As Boolean

'Syntax: Debug.Print X.SetFlat

    SetFlat = mvarSetFlat

End Property

Public Property Let SetFlat(ByVal vData As Boolean)

'Syntax: X.SetFlat = 5

    mvarSetFlat = vData

End Property

Public Property Get SmallChange(ByVal eBar As EFSScrollBarConstants) As Long

    If eBar = efsHorizontal Then
        SmallChange = m_lSmallChangeHorz
    Else                                              'NOT EBAR...
        SmallChange = m_lSmallChangeVert
    End If

End Property

Public Property Let SmallChange(ByVal eBar As EFSScrollBarConstants, _
                                ByVal lSmallChange As Long)

    If eBar = efsHorizontal Then
        m_lSmallChangeHorz = lSmallChange
    Else                                              'NOT EBAR...
        m_lSmallChangeVert = lSmallChange
    End If

End Property

Public Property Get Style() As EFSStyleConstants

    Style = m_eStyle

End Property

Public Property Let Style(ByVal eStyle As EFSStyleConstants)

'Dim lR As Long

    If eStyle <> efsRegular Then
        If m_bNoFlatScrollBars Then
            '-- can't do it..
            Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
            Exit Property                             '---> Bottom
        End If
    End If
    ShowScrollBar m_hWnd, 1, True
    If (m_eOrientation = efsoHorizontal) Or (m_eOrientation = efsoBoth) Then
        FlatSB_SetScrollProp m_hWnd, WSB_PROP_HSTYLE, eStyle, True
    End If
    If (m_eOrientation = efsoVertical) Or (m_eOrientation = efsoBoth) Then
        FlatSB_SetScrollProp m_hWnd, WSB_PROP_VSTYLE, eStyle, True
    End If
    m_eStyle = eStyle

End Property

Public Property Get Value(ByVal eBar As EFSScrollBarConstants) As Long

    Dim tSI    As SCROLLINFO

    pGetSI eBar, tSI, SIF_POS
    Value = tSI.nPos

End Property

Public Property Let Value(ByVal eBar As EFSScrollBarConstants, _
                          ByVal iValue As Long)

    Dim tSI    As SCROLLINFO

    If iValue <> Value(eBar) Then
        tSI.nPos = iValue
        pLetSI eBar, tSI, SIF_POS
        pRaiseEvent eBar, False
    End If

End Property

Public Property Get Visible(ByVal eBar As EFSScrollBarConstants) As Boolean

    If eBar = efsHorizontal Then
        Visible = m_bVisibleHorz
    Else                                              'NOT EBAR...
        Visible = m_bVisibleVert
    End If

End Property

Public Property Let Visible(ByVal eBar As EFSScrollBarConstants, _
                            ByVal bState As Boolean)

    If eBar = efsHorizontal Then
        m_bVisibleHorz = bState
    Else                                              'NOT EBAR...
        m_bVisibleVert = bState
    End If
    If m_bNoFlatScrollBars Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else                                              'M_BNOFLATSCROLLBARS = FALSE/0
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If

End Property

