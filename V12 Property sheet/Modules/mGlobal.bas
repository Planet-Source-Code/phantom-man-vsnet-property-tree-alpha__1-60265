Attribute VB_Name = "mGlobal"
Option Explicit

Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPosAPI Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long

Public Const WM_ACTIVATE As Integer = &H6
Public Const WM_NCHITTEST As Integer = &H84&

Public Type POINTAPI
    X          As Long
    Y          As Long
End Type
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

'-- Copy Memory For Ptr
#If Win32 Then
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                  lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#Else
    Public Declare Sub CopyMemory Lib "KERNEL" Alias "hmemcpy" ( _
                                  lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Declare Function DrawEdge Lib "user32" (ByVal HDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public CustClrs() As OLE_COLOR
Public LastSavedCustClr As Integer

Public DefClr  As OLE_COLOR
Public CurClr  As OLE_COLOR

Public DefCap  As String
Public MorCap  As String

Public ShwDef  As Boolean
Public ShwCus  As Boolean
Public ShwMor  As Boolean
Public ShwSys  As Boolean

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

'//---------------------------------------------------------------------------------------
'//--Procedure : AppThemed
'//--Type      : Function
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Determines If The Current Window is Themed
'//--Returns   : Boolean
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  01/02/2005
'//---------------------------------------------------------------------------------------
Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'//---------------------------------------------------------------------------------------
' Procedure : DrawDropDownArrow
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draw A Dropdown Arrow
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Sub DrawDropDownArrow( _
       ByVal lhdc As Long, _
       ByVal lLeft As Long, _
       ByVal lTop As Long, _
       ByVal lWidth As Long, _
       ByVal lHeight As Long, _
       ByVal bEnabled As Boolean, _
       ByVal oColor As OLE_COLOR, _
       Optional ByVal bUp As Boolean = False)

    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim Color  As Long

    Color = TranslateColor(oColor)

    lCentreX = lLeft + lWidth \ 2
    lCentreY = lTop + lHeight \ 2

    If Not bUp Then
        SetPixel lhdc, lCentreX - 2, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY - 1, Color
        SetPixel lhdc, lCentreX, lCentreY - 1, Color
        SetPixel lhdc, lCentreX + 1, lCentreY - 1, Color
        SetPixel lhdc, lCentreX + 2, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY, Color
        SetPixel lhdc, lCentreX + 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color
    Else                                              'NOT NOT...

        SetPixel lhdc, lCentreX, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY, Color
        SetPixel lhdc, lCentreX + 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY, Color
        SetPixel lhdc, lCentreX - 1, lCentreY + 1, Color
        SetPixel lhdc, lCentreX + 1, lCentreY + 1, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color
        SetPixel lhdc, lCentreX - 2, lCentreY + 1, Color
        SetPixel lhdc, lCentreX + 2, lCentreY + 1, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color

    End If

End Sub

Public Sub DrawHiddenPropertyArrow( _
       ByVal lhdc As Long, _
       ByVal lLeft As Long, _
       ByVal lTop As Long, _
       ByVal lWidth As Long, _
       ByVal lHeight As Long, _
       ByVal bEnabled As Boolean, _
       ByVal oColor As OLE_COLOR, _
       Optional ByVal bUp As Boolean = False)

    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim Color  As Long

    Color = TranslateColor(oColor)

    lCentreX = lLeft + lWidth \ 2
    lCentreY = lTop + lHeight \ 2

    SetPixel lhdc, lCentreX + 4, lCentreY - 2, Color
    SetPixel lhdc, lCentreX + 3, lCentreY - 2, Color
    SetPixel lhdc, lCentreX + 5, lCentreY - 2, Color
    SetPixel lhdc, lCentreX + 4, lCentreY - 3, Color

    SetPixel lhdc, lCentreX + 4, lCentreY, Color
    SetPixel lhdc, lCentreX + 4, lCentreY - 1, Color
    SetPixel lhdc, lCentreX + 4, lCentreY, Color
    SetPixel lhdc, lCentreX + 4, lCentreY + 1, Color

    SetPixel lhdc, lCentreX + 4, lCentreY + 2, Color
    SetPixel lhdc, lCentreX + 3, lCentreY + 2, Color
    SetPixel lhdc, lCentreX + 5, lCentreY + 2, Color
    SetPixel lhdc, lCentreX + 4, lCentreY + 3, Color

End Sub
Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR  As Long

    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color

    Call OleTranslateColor(oClr, hPal, TranslateColor)

End Function

