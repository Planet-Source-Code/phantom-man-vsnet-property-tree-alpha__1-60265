VERSION 5.00
Begin VB.UserControl pePropertySelector 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "pePropertySelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x          As Long
    y          As Long
End Type

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

' Change region of a window:
Private Declare Function SetWindowRgn Lib "user32" _
                                      (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" _
                                       (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
' GDI Clear up:
Private Declare Function DeleteObject Lib "gdi32" _
                                      (ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
' Region combination:
Private Declare Function CombineRgn Lib "gdi32" _
                                    (ByVal hDestRgn As Long, _
                                     ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
                                     ByVal nCombineMode As Long) As Long

' Region combination types:
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_MAX = RGN_COPY
Private Const RGN_MIN = RGN_AND
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
' Region combination return values:
Private Const COMPLEXREGION = 3
Private Const SIMPLEREGION = 2
Private Const NULLREGION = 1
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private m_bUserMode As Boolean

Private m_lButtonWidth As Long
Private m_hRgn As Long
Private m_rcButton As RECT
Private m_bMouseOver As Boolean
Private m_bMouseOverButton As Boolean
Private m_bMouseDownButton As Boolean
Private WithEvents m_tmrMouseTrack As IAPP_Timer
Attribute m_tmrMouseTrack.VB_VarHelpID = -1

Private m_oBorderColor As OLE_COLOR

Public Event Click()
Public Event DoHide()

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_oBorderColor
End Property
Public Property Let BorderColor(ByVal oColor As OLE_COLOR)
    If Not (m_oBorderColor = oColor) Then
        m_oBorderColor = oColor
        Draw
        PropertyChanged "BorderColor"
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    If Not (UserControl.BackColor = oColor) Then
        UserControl.BackColor = oColor
        PropertyChanged "BackColor"
    End If
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
    If Not (UserControl.ForeColor = oColor) Then
        UserControl.ForeColor = oColor
        PropertyChanged "ForeColor"
    End If
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bState As Boolean)
    If Not (UserControl.Enabled = bState) Then
        UserControl.Enabled = bState
        Draw
        PropertyChanged "Enabled"
    End If
End Property

Private Sub setRegion()
    Dim RC     As RECT
    Dim rcIn   As RECT
    Dim bRightToLeft As Boolean
    Dim hRgnOutBorder As Long
    Dim hRgnInBorder As Long
    Dim hRgnButton As Long
    Dim lR     As Long

    bRightToLeft = UserControl.RightToLeft

    GetClientRect UserControl.hWnd, RC

    m_hRgn = CreateRectRgn(0, 0, 0, 0)
    hRgnOutBorder = CreateRectRgn(RC.Left, RC.Top, RC.Right, RC.Bottom)
    LSet rcIn = RC
    InflateRect rcIn, -1, -1
    hRgnInBorder = CreateRectRgn(rcIn.Left, rcIn.Top, rcIn.Right, rcIn.Bottom)
    lR = CombineRgn(m_hRgn, hRgnOutBorder, hRgnInBorder, RGN_DIFF)

    m_rcButton.Top = rcIn.Top + 1
    m_rcButton.Bottom = rcIn.Bottom - 1
    If (bRightToLeft) Then
        m_rcButton.Left = rcIn.Left + 1
        m_rcButton.Right = rcIn.Left + m_lButtonWidth + 1
    Else
        m_rcButton.Left = rcIn.Right - m_lButtonWidth - 1
        m_rcButton.Right = rcIn.Right - 1
    End If
    hRgnButton = CreateRectRgn(m_rcButton.Left, m_rcButton.Top, m_rcButton.Right, m_rcButton.Bottom)

    lR = CombineRgn(m_hRgn, m_hRgn, hRgnButton, RGN_OR)

    DeleteObject hRgnButton
    DeleteObject hRgnInBorder
    DeleteObject hRgnOutBorder

    SetWindowRgn UserControl.hWnd, m_hRgn, 1

End Sub

Private Sub DrawDropDownArrow( _
        ByVal lhDC As Long, _
        ByVal lLeft As Long, _
        ByVal lTop As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByVal bEnabled As Boolean, _
        ByVal oColor As OLE_COLOR _
        )
    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim color  As Long

    color = TranslateColor(oColor)


    lCentreX = lLeft + lWidth \ 2
    lCentreY = lTop + lHeight \ 2

    SetPixel lhDC, lCentreX - 2, lCentreY - 1, color
    SetPixel lhDC, lCentreX - 1, lCentreY - 1, color
    SetPixel lhDC, lCentreX, lCentreY - 1, color
    SetPixel lhDC, lCentreX + 1, lCentreY - 1, color
    SetPixel lhDC, lCentreX + 2, lCentreY - 1, color
    SetPixel lhDC, lCentreX - 1, lCentreY, color
    SetPixel lhDC, lCentreX, lCentreY, color
    SetPixel lhDC, lCentreX + 1, lCentreY, color
    SetPixel lhDC, lCentreX, lCentreY + 1, color

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long
' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Function BlendColor( _
        ByVal oColorFrom As OLE_COLOR, _
        ByVal oColorTo As OLE_COLOR, _
        Optional ByVal Alpha As Long = 128 _
        ) As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    Dim lSrcR  As Long
    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000


    BlendColor = RGB( _
                 ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
                 ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
                 ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
                 )

End Function


Private Sub Draw()
    Dim lhDC   As Long
    Dim RC     As RECT
    Dim hBr    As Long
    Dim hBrBorder As Long
    Dim rcButton As RECT

    LSet rcButton = m_rcButton

    lhDC = UserControl.hDC
    GetClientRect UserControl.hWnd, RC

    ' Draw border
    hBrBorder = CreateSolidBrush(TranslateColor(m_oBorderColor))
    FrameRect lhDC, RC, hBrBorder

    ' Fill button
    If (m_bMouseDownButton) Then
        hBr = CreateSolidBrush(BlendColor(m_oBorderColor, UserControl.BackColor, 192))
    Else
        If (m_bMouseOverButton) Then
            hBr = CreateSolidBrush(BlendColor(m_oBorderColor, UserControl.BackColor))
        Else
            hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
        End If
    End If
    FillRect lhDC, rcButton, hBr
    DeleteObject hBr

    ' Draw drop down
    DrawDropDownArrow lhDC, rcButton.Left, rcButton.Top, _
                      rcButton.Right - rcButton.Left, rcButton.Bottom - rcButton.Top, _
                      True, UserControl.ForeColor

    ' Frame Button
    If (m_bMouseOverButton Or m_bMouseDownButton) Then
        FrameRect lhDC, rcButton, hBrBorder
    Else
        hBr = CreateSolidBrush(TranslateColor(m_oBorderColor))
        FrameRect lhDC, rcButton, hBr
        DeleteObject hBr
    End If

    DeleteObject hBrBorder

    UserControl.Refresh

End Sub

Private Sub pInitialise()
    If (UserControl.Ambient.UserMode) Then
        m_bUserMode = True
        Set m_tmrMouseTrack = New IAPP_Timer
    End If
End Sub

Private Sub m_tmrMouseTrack_ThatTime()
'
    Dim tP     As POINTAPI
    Dim RC     As RECT

    If (m_bMouseOver) Then
        GetCursorPos tP
        GetWindowRect UserControl.hWnd, RC
        If (PtInRect(RC, tP.x, tP.y) = 0) Then
            m_bMouseOver = False
            m_bMouseOverButton = False
            Draw
        Else
            If (m_bMouseOverButton) Then
                ScreenToClient UserControl.hWnd, tP
                If (PtInRect(m_rcButton, tP.x, tP.y) = 0) Then
                    m_bMouseOverButton = False
                    Draw
                End If
            End If
        End If
    End If

    If Not (m_bMouseOver) Then
        m_tmrMouseTrack.Interval = 0
    End If
    '
End Sub



Private Sub UserControl_Initialize()
    m_lButtonWidth = 18
    m_oBorderColor = vbHighlight
End Sub

Private Sub UserControl_InitProperties()
'
    pInitialise
    '
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
    If (m_bUserMode And Enabled) Then
        If (Button = vbLeftButton) Or (Button = vbRightButton) Then
            If Not (PtInRect(m_rcButton, x, y) = 0) Then
                m_bMouseOver = True
                m_bMouseOverButton = True
                m_bMouseDownButton = True
                Draw

                RaiseEvent Click

                m_tmrMouseTrack_ThatTime

            End If
        End If
    End If
    '
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim bChange As Boolean
    '
    If (m_bUserMode And Enabled) Then

        If Not (PtInRect(m_rcButton, x, y) = 0) Then
            If Not (m_bMouseOverButton) Then
                m_bMouseOverButton = True
                bChange = True
            End If
        End If
        If Not (m_bMouseOver) Then
            m_bMouseOver = True
            bChange = True
        End If

        If (bChange) Then
            m_tmrMouseTrack.Interval = 50
            Draw
        End If

    End If
    '
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
    If (m_bMouseDownButton And Enabled) Then
        m_bMouseDownButton = False
        m_bMouseOver = False
        m_bMouseOverButton = False
        RaiseEvent DoHide
        Draw
    End If
    '
End Sub

Public Sub Repaint()
    Draw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
    BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    BorderColor = PropBag.ReadProperty("BorderColor", vbHighlight)
    Enabled = PropBag.ReadProperty("Enabled", True)
    '
    pInitialise
    '
End Sub

Private Sub UserControl_Resize()

    setRegion

    Draw

End Sub

Private Sub UserControl_Show()
    setRegion
End Sub

Private Sub UserControl_Terminate()
    If Not (m_tmrMouseTrack Is Nothing) Then
        m_tmrMouseTrack.Interval = 0
        Set m_tmrMouseTrack = Nothing
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
    PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
    PropBag.WriteProperty "BorderColor", BorderColor, vbHighlight
    PropBag.WriteProperty "Enabled", Enabled, True
    '
End Sub
