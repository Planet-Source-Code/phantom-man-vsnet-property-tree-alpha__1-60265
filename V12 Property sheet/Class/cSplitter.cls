VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_SplitterBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-------------------------------------------------------------------------------------------
'-- Class:    IAPP_SplitterBar
'-- Filename: IAPP_SplitterBar.cls
'-- Author:   Gary Noble (gwnoble@msn.com)
'-- Date:     15 Jan 2002
'-------------------------------------------------------------------------------------------
'-- History
'-------------------------------------------------------------------------------------------
'--         15-1-2002  Initial Version  Gary Noble
'-------------------------------------------------------------------------------------------
'-- Notes:
'--
'--
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
'-- Global declarations
'-------------------------------------------------------------------------------------------

Private bDraw  As Boolean
Private rcCurrent As RECT
Private rcNew  As RECT
Private rcWindow As RECT
Private rcWindows As RECT

Private LastX  As Long
Private LastY  As Long
Private m_lSplitterBottomOffSet As Long

Dim hBrush     As Long

Private Type POINTAPI
    X          As Long
    Y          As Long
End Type

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Declare Function Rectangle Lib "gdi32" (ByVal HDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal HDC As Long, ByVal nDrawMode As Long) As Long
Private Const R2_BLACK = 1
Private Const R2_COPYPEN = 13
Private Const R2_LAST = 16
Private Const R2_MASKNOTPEN = 3
Private Const R2_MASKPEN = 9
Private Const R2_MASKPENNOT = 5
Private Const R2_MERGENOTPEN = 12
Private Const R2_MERGEPEN = 15
Private Const R2_MERGEPENNOT = 14
Private Const R2_NOP = 11
Private Const R2_NOT = 6
Private Const R2_NOTCOPYPEN = 4
Private Const R2_NOTMASKPEN = 8
Private Const R2_NOTMERGEPEN = 2
Private Const R2_NOTXORPEN = 7                        '10
Private Const R2_WHITE = 16
Private Const R2_XORPEN = 7
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15


'-------------------------------------------------------------------------------------------
'-- COMBINEREGION
'-------------------------------------------------------------------------------------------
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Const WM_NCACTIVATE As Integer = &H86
Private Const WM_MOVE As Integer = &H3
Private Const WM_CLOSE As Integer = &H10
Private Const HTCAPTION = 2
Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONCLOSE As Integer = &H0
Private Const DFCS_FLAT As Integer = &H4000
Private Const DFCS_INACTIVE As Integer = &H100
Private Const DFCS_PUSHED As Integer = &H200
Private Const SPI_GETNONCLIENTMETRICS = 41

'-------------------------------------------------------------------------------------------
'-- Sizer constants
'-------------------------------------------------------------------------------------------
Private Const HTRIGHT = 11
Private Const HTLEFT = 10
Private Const HTTOP = 12
Private Const HTBOTTOM = 15
Private Const HTBOTTOMRIGHT = 17


'-------------------------------------------------------------------------------------------
'-- global declarations
'-------------------------------------------------------------------------------------------
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function apiCreateBitmap Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function apiCreatePatternBrush Lib "gdi32" Alias "CreatePatternBrush" (ByVal hBitmap As Long) As Long
Private Declare Function apiCreateRectRgn Lib "gdi32" Alias "CreateRectRgn" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function apiCreateRectRgnIndirect Lib "gdi32" Alias "CreateRectRgnIndirect" (lpRect As RECT) As Long
Private Declare Function APIDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function apiGetClipBox Lib "gdi32" Alias "GetClipBox" (ByVal HDC As Long, lpRect As RECT) As Long
Private Declare Function apiPatBlt Lib "gdi32" Alias "PatBlt" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function apiSelectClipRgn Lib "gdi32" Alias "SelectClipRgn" (ByVal HDC As Long, ByVal hRgn As Long) As Long
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function apiSetRectRgn Lib "gdi32" Alias "SetRectRgn" (ByVal hRgn As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'-------------------------------------------------------------------------------------------
'-- USER32
'-------------------------------------------------------------------------------------------
Private Declare Function apiCopyRect Lib "user32" Alias "CopyRect" (lpDstRect As RECT, lpSrcRect As RECT) As Long
Private Declare Function ApiFillRect Lib "user32" Alias "FillRect" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function apiInflateRect Lib "user32" Alias "InflateRect" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function apiIntersectRect Lib "user32" Alias "IntersectRect" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function apiIsRectEmpty Lib "user32" Alias "IsRectEmpty" (lpRect As RECT) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal HDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetWindowPosAPI Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private m_hdc  As Long
Private m_hBmp As Long
Private m_hBmpOld As Long

'-------------------------------------------------------------------------------------------
'-- private members
'-------------------------------------------------------------------------------------------
Private m_lngPower2(0 To 31) As Long

'-------------------------------------------------------------------------------------------
'-- client metrics (for caption height)
'-------------------------------------------------------------------------------------------
Private m_bButtonDown As Boolean
Private m_bButtonOver As Boolean
Private m_bMouseDown As Boolean
Private m_bRunTime As Boolean
Private m_sCaption As String
Private m_bDragging As Boolean
Private m_lOldSize As Long

Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants
Private m_lBorder(1 To 4) As Long
Private m_lSpecial(1 To 4) As Long
Private m_oSplit As Object

Private m_bUseInternalAjustments As Boolean
Private m_lSPBottom As Long
Private m_lSPTop As Long

Private m_bIsMDI As Boolean
Private m_bSplitting As Boolean

Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum

Public Enum ESplitBorderTypes
    espbLeft = 1
    espbtop = 2
    espbright = 3
    espbBottom = 4
End Enum

Public Enum ESPValuesTypes
    espvleft = 1
    espvtop = 2
    espvright = 3
    espvBottom = 4
End Enum

Public Property Let UseInternalAjustments(ByVal vData As Boolean)
    m_bUseInternalAjustments = vData
End Property


Public Property Get UseInternalAjustments() As Boolean
    UseInternalAjustments = m_bUseInternalAjustments
End Property

Public Property Get SplitObject() As Object
    Set SplitObject = m_oSplit
End Property
Public Property Let SplitObject(ByRef oThis As Object)
    Set m_oSplit = oThis
    On Error Resume Next
    '    oThis.BorderStyle = 0
    If (m_eOrientation = espHorizontal) Then
        oThis.MousePointer = vbSizeNS
    Else
        '        oThis.MousePointer = vbSizeWE
    End If
End Property
Public Property Let Border(ByVal eBorderType As ESplitBorderTypes, ByVal lSize As Long)
    m_lBorder(eBorderType) = lSize
End Property
Public Property Get Border(ByVal eBorderType As ESplitBorderTypes) As Long
    Border = m_lBorder(eBorderType)
End Property

Public Property Let SPValues(ByVal eSPV As ESPValuesTypes, ByVal lSize As Long)
    m_lSpecial(eSPV) = lSize
End Property
Public Property Get SPValues(ByVal eSPV As ESPValuesTypes) As Long
    SPValues = m_lSpecial(eSPV)
End Property


Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)

    m_eOrientation = eOrientation

    If Not (m_oSplit Is Nothing) Then

        If (m_eOrientation = espHorizontal) Then

            m_oSplit.MousePointer = vbSizeNS

            m_lBorder(espbtop) = 64

            m_lBorder(espbBottom) = 64

            m_lBorder(espbLeft) = 0

            m_lBorder(espbright) = 0

        Else

            'm_oSplit.MousePointer = vbSizeWE

            m_lBorder(espbtop) = 0

            m_lBorder(espbBottom) = 0

            m_lBorder(espbLeft) = 64

            m_lBorder(espbright) = 64

        End If

    End If

End Property

Public Property Let SplitterBottomOffSet(lY As Long)

    m_lSplitterBottomOffSet = lY

End Property

Public Sub SplitterMouseDown( _
       ByVal hwnd As Long, _
       ByVal X As Long, _
       ByVal Y As Long _
       )
    Dim tP     As POINTAPI

    m_hWnd = hwnd


    Dim cd     As Control


    '-------------------------------------------------------------------------------------------
    '-- Send subsequent mouse messages to the owner window
    '-------------------------------------------------------------------------------------------
    SetCapture m_hWnd

    '-------------------------------------------------------------------------------------------
    '-- Get the window rectangle on the desktop of the owner window
    '-------------------------------------------------------------------------------------------
    GetWindowRect m_hWnd, rcWindow

    '-------------------------------------------------------------------------------------------
    '-- Set The Splitter Limits
    '-------------------------------------------------------------------------------------------
    If (m_eOrientation = espVertical) Then

        rcWindow.Left = rcWindow.Left + m_lBorder(espbLeft) + 10

        rcWindow.Right = rcWindow.Right - m_lBorder(espbright)

    Else

        rcWindow.Left = rcWindow.Left + m_lBorder(espbLeft)

        rcWindow.Right = rcWindow.Right - m_lBorder(espbright)

        rcWindow.Top = rcWindow.Top + m_lBorder(espbtop)

        rcWindow.Bottom = rcWindow.Bottom - m_lBorder(espbBottom)

    End If



    '-------------------------------------------------------------------------------------------
    '-- Clip the cursor so it can't move outside the window
    '-------------------------------------------------------------------------------------------
    ClipCursorRect rcWindow

    '-------------------------------------------------------------------------------------------
    '-- Check if this is an MDI form
    '-------------------------------------------------------------------------------------------
    If (ClassName(m_hWnd) = "ThunderMDIForm") Then
        '-------------------------------------------------------------------------------------------
        ' Get the inside portion of the MDI form:
        ' Assuming we have a caption,menu and border in your MDI
        '-------------------------------------------------------------------------------------------
        rcWindow.Left = rcWindow.Left + GetSystemMetrics(SM_CXBORDER)

        rcWindow.Right = rcWindow.Right - GetSystemMetrics(SM_CXBORDER)

        rcWindow.Bottom = rcWindow.Bottom - GetSystemMetrics(SM_CYBORDER)

        rcWindow.Top = rcWindow.Top + GetSystemMetrics(SM_CYBORDER) * 3 + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU)

        m_bIsMDI = True

        tP.X = rcWindow.Right

        tP.Y = rcWindow.Bottom


        If m_eOrientation = espVertical Then

            If SPValues(espvtop) > 0 Then
                rcWindow.Top = rcWindow.Top + SPValues(espvtop) / Screen.TwipsPerPixelY
            End If

            If SPValues(espvBottom) > 0 Then
                rcWindow.Bottom = rcWindow.Bottom - SPValues(espvBottom) / Screen.TwipsPerPixelY
            End If



        End If


    Else


        '-------------------------------------------------------------------------------------------
        '-- Get the client rectangle of the window in screen coordinates
        '-------------------------------------------------------------------------------------------
        GetClientRect m_hWnd, rcWindow

        If m_eOrientation = espVertical Then

            tP.X = rcWindow.Left + m_lBorder(espbLeft)

            tP.Y = rcWindow.Top

        End If

        ClientToScreen m_hWnd, tP
        rcWindow.Left = tP.X
        rcWindow.Top = tP.Y

        If m_eOrientation = espHorizontal Then

            tP.X = rcWindow.Left
            tP.X = rcWindow.Right

            If SPValues(espvleft) > 0 Then
                rcWindow.Left = rcWindow.Left + SPValues(espvleft) / Screen.TwipsPerPixelX
            End If

            If SPValues(espvright) > 0 Then
                rcWindow.Right = SPValues(espvright) / Screen.TwipsPerPixelX    ' rcWindow.Right - ((SPValues(espvright) / Screen.TwipsPerPixelX))  'SPValues(espvright) * Screen.TwipsPerPixelX))
                tP.X = rcWindow.Right
            End If

        Else

            tP.X = rcWindow.Right

            tP.Y = rcWindow.Bottom


            '-------------------------------------------------------------------------------------------
            '-- Adjust Rectangle Accordingly
            '-------------------------------------------------------------------------------------------
            If SPValues(espvBottom) > 0 Then
                tP.Y = rcWindow.Bottom - SPValues(espvBottom) / Screen.TwipsPerPixelY
            End If

            If SPValues(espvtop) > 0 Then
                rcWindow.Top = rcWindow.Top + SPValues(espvtop) / Screen.TwipsPerPixelY
            End If

        End If

        ClientToScreen m_hWnd, tP

        '-------------------------------------------------------------------------------------------
        '-- Set Parameters
        '-------------------------------------------------------------------------------------------
        rcWindow.Right = tP.X
        rcWindow.Bottom = tP.Y

        If m_eOrientation = espHorizontal Then

            tP.X = rcWindow.Right

            tP.Y = rcWindow.Bottom

        Else

            tP.X = rcWindow.Right

            tP.Y = rcWindow.Bottom

        End If


        m_bIsMDI = False

    End If

    bDraw = True                                      '-- start actual drawing from next move message

    rcCurrent.Left = 0: rcCurrent.Top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0

    '    X = (m_oSplit.Left + X) \ Screen.TwipsPerPixelX
    '    Y = (m_oSplit.Top + Y) \ Screen.TwipsPerPixelY

    SplitterContainer_MouseMove X, Y

End Sub

Public Function SplitterContainer_MouseMove( _
       ByVal X As Long, _
       ByVal Y As Long)


    Dim HDC    As Long
    Dim tP     As POINTAPI
    Dim hBrush As Long

    If (bDraw) Then
        '-----------------------------------------------------------------------
        '-- Draw two rectangles in the screen DC to cause splitting:
        '-- First get the Desktop DC:
        '-----------------------------------------------------------------------
        HDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

        '-----------------------------------------------------------------------
        '-- Set the draw mode to XOR:
        '-----------------------------------------------------------------------
        ' SetROP2 hDC, R2_NOTXORPEN

        '-----------------------------------------------------------------------
        '-- Get The Cursor Position
        '-----------------------------------------------------------------------
        GetCursorPos tP

        '-----------------------------------------------------------------------
        '-- Determine where to draw the splitter:
        '-----------------------------------------------------------------------
        If m_bIsMDI Then
            If (m_eOrientation = espHorizontal) Then

                rcNew.Left = rcWindow.Left

                rcNew.Right = rcWindow.Right

                If (tP.Y >= rcWindow.Top) And (tP.Y < rcWindow.Bottom) Then

                    rcNew.Top = tP.Y - 4

                    rcNew.Bottom = tP.Y + 2

                Else
                    If (tP.Y < rcWindow.Top) Then

                        rcNew.Top = rcWindow.Top - 10

                        rcNew.Bottom = rcNew.Top + 20

                    Else

                        rcNew.Top = rcWindow.Bottom - 10

                        rcNew.Bottom = rcNew.Top + 20

                    End If

                End If

            Else


                '-----------------------------------------------------------------------
                '-- Vertical
                '-----------------------------------------------------------------------
                rcNew.Top = rcWindow.Top
                rcNew.Bottom = rcWindow.Bottom



                If (tP.X >= rcWindow.Left) And (tP.X <= rcWindow.Right) Then

                    rcNew.Left = tP.X                 '- 1

                    rcNew.Right = tP.X + 5

                Else

                    If (tP.X < rcWindow.Left) Then

                        rcNew.Left = rcWindow.Left    '+ 5 ' 40

                        rcNew.Right = rcNew.Left + 5  ' + 5

                    Else

                        rcNew.Left = rcWindow.Right   '+ 5 ' 50

                        rcNew.Right = rcNew.Left      '+ 5

                    End If

                End If

            End If

        Else


            If (m_eOrientation = espHorizontal) Then

                rcNew.Left = rcWindow.Left

                rcNew.Right = rcWindow.Right

                If (tP.Y >= rcWindow.Top) And (tP.Y < rcWindow.Bottom) Then

                    rcNew.Top = tP.Y - 4

                    rcNew.Bottom = tP.Y + 2

                Else                                  '(m_eOrientation = espVertical)

                    If (tP.Y < rcWindow.Top) Then

                        rcNew.Top = rcWindow.Top - 10

                        rcNew.Bottom = rcNew.Top + 20

                    Else

                        rcNew.Top = rcWindow.Bottom - 10

                        rcNew.Bottom = rcNew.Top + 20

                    End If

                End If

            Else


                '-----------------------------------------------------------------------
                '-- Vertical
                '-----------------------------------------------------------------------
                rcNew.Top = rcWindow.Top

                '-- Adjust Splitter Rect To Handle Custom Size
                If m_lSplitterBottomOffSet > 0 Then
                    rcNew.Bottom = rcNew.Top + (m_lSplitterBottomOffSet)
                Else
                    rcNew.Bottom = rcWindow.Bottom
                End If

                If (tP.X >= rcWindow.Left) And (tP.X <= rcWindow.Right) Then

                    rcNew.Left = tP.X - 1

                    rcNew.Right = tP.X + 2

                Else

                    If (tP.X < rcWindow.Left) Then

                        rcNew.Left = rcWindow.Left + 5    ' 40

                        rcNew.Right = rcNew.Left + 2

                    Else

                        rcNew.Left = rcWindow.Right + 5    ' 50

                        rcNew.Right = rcNew.Left + 2

                    End If

                End If

            End If

        End If

DrawSplitter:

        '-----------------------------------------------------------------------
        '-- Draw the new rectangle
        '-----------------------------------------------------------------------
        hBrush = gdiCreateHalftoneBrush()
        gdiDrawDragRectangle HDC, rcNew.Left, rcNew.Top, rcNew.Right - rcNew.Left, rcNew.Bottom - rcNew.Top, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right - rcCurrent.Left, rcCurrent.Bottom - rcCurrent.Top, 15, 15, hBrush, 0

        '-----------------------------------------------------------------------
        '-- Store this position so we can erase it next time:
        '-----------------------------------------------------------------------
        LSet rcCurrent = rcNew

        '-----------------------------------------------------------------------
        '-- Free the reference to the Desktop DC we got (make sure you do this!)
        '-----------------------------------------------------------------------
        DeleteDC HDC


    End If


End Function

Public Function SplitterContainer_MouseUp(ByVal X As Long, ByVal Y As Long) As Boolean


    Dim HDC    As Long
    Dim tP     As POINTAPI
    Dim hWndClient As Long
    Dim hBrush As Long

    '-------------------------------------------------------------------------------------------
    '-- Don't leave orphaned rectangle on desktop; erase last rectangle.
    '-------------------------------------------------------------------------------------------
    If (bDraw) Then
        bDraw = False

        '-------------------------------------------------------------------------------------------
        '-- Release mouse capture
        '-------------------------------------------------------------------------------------------
        ReleaseCapture

        '-------------------------------------------------------------------------------------------
        '-- Release the cursor clipping region (must do this!)
        '-------------------------------------------------------------------------------------------
        ClipCursorClear 0&

        '-------------------------------------------------------------------------------------------
        '-- Get the Desktop DC
        '-------------------------------------------------------------------------------------------
        HDC = CreateDCAsNull("DISPLAY", 0, 0, 0)

        '-------------------------------------------------------------------------------------------
        '-- Set to XOR drawing mode
        '-- SetROP2 hDC, R2_NOTXORPEN
        '-- Erase the last rectangle
        '-------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------
        '-- Modification: New Splitter Implementation (01-05-2002 G Noble)
        '-- Delete The Last Splitter Bar
        '-----------------------------------------------------------------------
        hBrush = gdiCreateHalftoneBrush()
        gdiDrawDragRectangle HDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right - rcCurrent.Left, rcCurrent.Bottom - rcCurrent.Top, 0, 0, 0, 0, 5, 5, hBrush, 0

        '-------------------------------------------------------------------------------------------
        '-- Clear up the desktop DC
        '-------------------------------------------------------------------------------------------
        DeleteDC HDC

        '-------------------------------------------------------------------------------------------
        '-- Here we ensure the splitter is within bounds before releasing
        '-------------------------------------------------------------------------------------------
        GetCursorPos tP

        ScreenToClient m_hWnd, tP

        '-------------------------------------------------------------------------------------------
        '-- Move the splitter to the validated final position
        '-------------------------------------------------------------------------------------------
        If (m_eOrientation = espHorizontal) Then

            m_oSplit.Top = (tP.Y - 4) * Screen.TwipsPerPixelY

        Else

            '           m_oSplit.Left = (tp.X) * Screen.TwipsPerPixelX

        End If

        '-------------------------------------------------------------------------------------------
        '-- Return true to tell the owner we have completed splitting
        '-------------------------------------------------------------------------------------------
        SplitterContainer_MouseUp = True

    End If

    RefreshALLControls


End Function

Private Function RefreshALLControls()
    On Error Resume Next

    Dim ctrl   As Control

    For Each ctrl In Me.SplitObject.Parent.Controls
        ctrl.Refresh
    Next



End Function

Private Sub Class_Initialize()

    m_eOrientation = espVertical

    m_lBorder(espbLeft) = 64

    m_lBorder(espbright) = 64

End Sub
Private Function ClassName(ByVal lhWnd As Long) As String

    Dim lLen   As Long
    Dim sBuf   As String

    lLen = 260

    sBuf = String$(lLen, 0)

    lLen = GetClassName(lhWnd, sBuf, lLen)

    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If

End Function





Private Sub gdiDrawDragRectangle(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal LastX As Long, ByVal LastY As Long, ByVal LastDX As Long, ByVal LastDY As Long, ByVal Size As Long, ByVal LastSize As Long, ByVal hBrush As Long, ByVal hBrushLast As Long, Optional ByVal hClipRgn As Long)

'-----------------------------------------------------------------------
'-- draw drag rectangle, erasing old as needed
'-----------------------------------------------------------------------

    Dim tRect As RECT, tRectLast As RECT
    Dim hOutsideRgn As Long, hInsideRgn As Long
    Dim hLastRgn As Long, hUpdateRgn As Long
    Dim hNewRgn As Long
    Dim tNewRect As RECT
    Dim hOldBrush As Long

    '    Debug.Assert hBrush <> 0

    '-----------------------------------------------------------------------
    '-- init vars
    '-----------------------------------------------------------------------
    With tRect
        .Left = X
        .Top = Y
        .Right = X + dx
        .Bottom = Y + dy

    End With
    With tRectLast
        .Left = LastX
        .Top = LastY
        .Right = LastX + LastDX
        .Bottom = LastY + LastDY

    End With

    '-----------------------------------------------------------------------
    '-- determine update region and select it
    '-----------------------------------------------------------------------
    hOutsideRgn = apiCreateRectRgnIndirect(tRect)
    apiCopyRect tNewRect, tRect
    apiInflateRect tNewRect, -Size, -Size
    apiIntersectRect tNewRect, tNewRect, tRect
    hInsideRgn = apiCreateRectRgnIndirect(tNewRect)
    hNewRgn = apiCreateRectRgn(0, 0, 0, 0)
    CombineRgn hNewRgn, hOutsideRgn, hInsideRgn, RGN_XOR
    If hBrushLast = 0 Then
        hBrushLast = hBrush

    End If
    If apiIsRectEmpty(tRectLast) = 0 Then

        '-----------------------------------------------------------------------
        '-- find difference between new region and old region
        '-----------------------------------------------------------------------
        hLastRgn = apiCreateRectRgn(0, 0, 0, 0)
        apiSetRectRgn hOutsideRgn, tRectLast.Left, tRectLast.Top, tRectLast.Right, tRectLast.Bottom
        apiCopyRect tNewRect, tRectLast
        apiInflateRect tNewRect, -LastSize, -LastSize
        apiIntersectRect tNewRect, tNewRect, tRectLast
        apiSetRectRgn hInsideRgn, tNewRect.Left, tNewRect.Top, tNewRect.Right, tNewRect.Bottom
        CombineRgn hLastRgn, hOutsideRgn, hInsideRgn, RGN_XOR

        '-----------------------------------------------------------------------
        '-- only diff them if brushes are the same
        '-----------------------------------------------------------------------
        If (hBrush = hBrushLast) Then
            hUpdateRgn = apiCreateRectRgn(0, 0, 0, 0)
            CombineRgn hUpdateRgn, hLastRgn, hNewRgn, RGN_XOR

        End If

    End If
    If (hBrush <> hBrushLast) And (apiIsRectEmpty(tRectLast) = 0) Then
        '-----------------------------------------------------------------------
        '-- brushes are different -- erase old region first
        '-----------------------------------------------------------------------
        apiSelectClipRgn HDC, hLastRgn
        apiGetClipBox HDC, tNewRect
        hOldBrush = apiSelectObject(HDC, hBrushLast)
        apiPatBlt HDC, tNewRect.Left, tNewRect.Top, tNewRect.Right - tNewRect.Left, tNewRect.Bottom - tNewRect.Top, vbPatInvert
        apiSelectObject HDC, hOldBrush
        APIDeleteObject hOldBrush

    End If

    '-----------------------------------------------------------------------
    '-- draw into the update/new region
    '-----------------------------------------------------------------------
    If hUpdateRgn <> 0 Then
        If hClipRgn <> 0 Then
            CombineRgn hUpdateRgn, hClipRgn, hUpdateRgn, RGN_AND

        End If
        apiSelectClipRgn HDC, hUpdateRgn

    Else
        If hClipRgn <> 0 Then
            CombineRgn hNewRgn, hClipRgn, hNewRgn, RGN_AND

        End If
        apiSelectClipRgn HDC, hNewRgn

    End If
    apiGetClipBox HDC, tNewRect
    hOldBrush = apiSelectObject(HDC, hBrush)
    apiPatBlt HDC, tNewRect.Left, tNewRect.Top, tNewRect.Right - tNewRect.Left, tNewRect.Bottom - tNewRect.Top, vbPatInvert

    '-----------------------------------------------------------------------
    '-- cleanup DC
    '-----------------------------------------------------------------------
    If (hOldBrush <> 0) Then
        apiSelectObject HDC, hOldBrush
    End If
    apiSelectClipRgn HDC, 0&

    '-----------------------------------------------------------------------
    '-- delete objects (A MUST)
    '-----------------------------------------------------------------------
    If hUpdateRgn <> 0 Then
        APIDeleteObject hUpdateRgn
    End If

    If hLastRgn <> 0 Then
        APIDeleteObject hLastRgn

    End If
    If hNewRgn <> 0 Then
        APIDeleteObject hNewRgn

    End If
    If hInsideRgn <> 0 Then
        APIDeleteObject hInsideRgn

    End If
    If hOutsideRgn <> 0 Then
        APIDeleteObject hOutsideRgn
    End If

    '-- >>
    If hBrush <> 0 Then
        APIDeleteObject hBrush
    End If

    If hOldBrush <> 0 Then
        APIDeleteObject hOldBrush
    End If


End Sub
Private Sub gdiDrawSolidRectangleBrush(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal hBrush As Long)

'-----------------------------------------------------------------------
'-- draw filled rectangle using supplied brush
'-----------------------------------------------------------------------
    Dim tRect  As RECT

    '-----------------------------------------------------------------------
    '-- build RECT
    '-----------------------------------------------------------------------
    gdiSetRect tRect, X, Y, dx - 2, dy

    '-----------------------------------------------------------------------
    '-- draw filled rect
    '-----------------------------------------------------------------------
    ApiFillRect HDC, tRect, hBrush

End Sub
Private Sub gdiSetRect(tRect As RECT, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long)

'-----------------------------------------------------------------------
'-- build RECT
'-----------------------------------------------------------------------
    With tRect
        .Left = X
        .Top = Y
        .Right = X + dx
        .Bottom = Y + dy

    End With
End Sub
Public Sub gdiDeleteBrush(ByVal hBrush As Long)

'-----------------------------------------------------------------------
'-- delete brush object
'-----------------------------------------------------------------------
    APIDeleteObject hBrush

End Sub
Private Function gdiCreateHalftoneBrush() As Long

'-----------------------------------------------------------------------
'-- Create a halftoned brush object
'-----------------------------------------------------------------------

    Dim nGreyPattern(8) As Integer
    Dim nBitCount As Integer
    Dim hBitmap As Long
    Dim lngResult As Long

    '-----------------------------------------------------------------------
    '-- create greyed bitmap
    '-----------------------------------------------------------------------
    For nBitCount = 0 To 7
        nGreyPattern(nBitCount) = LShiftWord(&H5555, (nBitCount And 1))

    Next nBitCount
    hBitmap = apiCreateBitmap(8, 8, 1, 1, nGreyPattern(0))

    '-----------------------------------------------------------------------
    '-- create halftone pattern brush
    '-----------------------------------------------------------------------
    gdiCreateHalftoneBrush = apiCreatePatternBrush(hBitmap)

    '-----------------------------------------------------------------------
    '-- delete bitmap
    '-----------------------------------------------------------------------
    APIDeleteObject hBitmap
    '->>gobj_APIDeleteObject gdiCreateHalftoneBrush


End Function
Private Function LShiftWord(ByVal nWord As Integer, ByVal nBitCount As Integer) As Integer

'-----------------------------------------------------------------------
'-- left shift dword value lngDWord by nBitCount
'-----------------------------------------------------------------------
    Dim lngDWord As Long

    Debug.Assert (nBitCount >= 0 And nBitCount <= 15)    'design-time only
    If Not (nBitCount >= 0 And nBitCount <= 15) Then Exit Function

    lngDWord = nWord * Power2(nBitCount)
    If lngDWord And &H8000& Then
        LShiftWord = CInt(lngDWord And &H7FFF&) Or &H8000

    Else
        LShiftWord = lngDWord And &HFFFF&

    End If
End Function
Private Function Power2(ByVal nIndex As Integer) As Long

'-----------------------------------------------------------------------
'-- build lookup table for bitshifting operations
'-----------------------------------------------------------------------
    If Not (nIndex >= 0 And nIndex <= 31) Then Exit Function

    If m_lngPower2(0) = 0 Then                        'only set array if not yet been used
        m_lngPower2(0) = &H1&
        m_lngPower2(1) = &H2&
        m_lngPower2(2) = &H4&
        m_lngPower2(3) = &H8&
        m_lngPower2(4) = &H10&
        m_lngPower2(5) = &H20&
        m_lngPower2(6) = &H40&
        m_lngPower2(7) = &H80&
        m_lngPower2(8) = &H100&
        m_lngPower2(9) = &H200&
        m_lngPower2(10) = &H400&
        m_lngPower2(11) = &H800&
        m_lngPower2(12) = &H1000&
        m_lngPower2(13) = &H2000&
        m_lngPower2(14) = &H4000&
        m_lngPower2(15) = &H8000&
        m_lngPower2(16) = &H10000
        m_lngPower2(17) = &H20000
        m_lngPower2(18) = &H40000
        m_lngPower2(19) = &H80000
        m_lngPower2(20) = &H100000
        m_lngPower2(21) = &H200000
        m_lngPower2(22) = &H400000
        m_lngPower2(23) = &H800000
        m_lngPower2(24) = &H1000000
        m_lngPower2(25) = &H2000000
        m_lngPower2(26) = &H4000000
        m_lngPower2(27) = &H8000000
        m_lngPower2(28) = &H10000000
        m_lngPower2(29) = &H20000000
        m_lngPower2(30) = &H40000000
        m_lngPower2(31) = &H80000000

    End If
    Power2 = m_lngPower2(nIndex)
End Function







