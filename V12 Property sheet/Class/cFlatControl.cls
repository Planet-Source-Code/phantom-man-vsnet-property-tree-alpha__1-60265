VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const WM_TIMER = &H113
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const SM_CXHTHUMB = 10

Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21

Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_RIGHT = &H1000&
Private Const WS_EX_LEFTSCROLLBAR = &H4000&
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Const CBN_DROPDOWN = 7
Private Const CBN_CLOSEUP = 8
Private Const CB_GETDROPPEDSTATE = &H157
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function RedrawWindow Lib "user32" ( _
   ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_ERASENOW = &H200
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100
Private Const RDW_VALIDATE = &H8
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Implements ISubclass

Public Enum EFlatControlStyle
   eFlatStyleOffice9
   eFlatStyleOffice10
   eFlatStyleOffice11
End Enum

Private Enum EDrawStyle
   FC_DRAWNORMAL = &H1
   FC_DRAWRAISED = &H2
   FC_DRAWPRESSED = &H4
End Enum

Private m_hWnd             As Long
Private m_hWndEdit         As Long
Private m_hWndParent       As Long
Private m_bSubclass        As Boolean
Private m_bMouseOver       As Boolean

Private m_bLBtnDown As Boolean
Private m_bCombo As Boolean
Private m_bText As Boolean

Private m_eStyle As EFlatControlStyle

Public Property Get FlatStyle() As EFlatControlStyle
   FlatStyle = m_eStyle
End Property
Public Property Let FlatStyle(ByVal eStyle As EFlatControlStyle)
   m_eStyle = eStyle
End Property


Public Sub Attach(ByRef objthis As Object)
Dim lStyle As Long
Dim lhWnd As Long
   
   pRelease
   
   On Error Resume Next
   lhWnd = objthis.hWnd
   If (Err.Number <> 0) Then
      Err.Raise vbObjectError + 1048 + 1, _
                App.EXEName & ".cFlatControl", _
                "Incorrect control type passed to 'Attach' parameter - must be a control with a hWnd property."
      Exit Sub
   End If
   
   m_bCombo = False
   m_bText = False
   
   ' Check for standard VB types:
   If TypeName(objthis) = "ImageCombo" Then
      ' Support VB6 ImageCombos.  Thanks to Paul Wilde for the tip.
      ' WM_COMMAND notifications to the control:
      m_hWndParent = lhWnd
      ' Get the actual ComboBox in the control:
      lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&)
      m_bCombo = True
   ElseIf TypeName(objthis) = "ComboBox" Then
      ' WM_COMMAND notifications go to the parent:
      m_hWndParent = GetParent(objthis.hWnd)
      m_bCombo = True
   ElseIf TypeName(objthis) = "OwnerDrawComboList" Then
      ' WM_COMMAND notifications to the control:
      m_hWndParent = lhWnd
      m_bCombo = True
   ElseIf TypeName(objthis) = "TextBox" Then
      m_bText = True
      ' Shrink...
      With objthis
         .Move .Left + 2 * Screen.TwipsPerPixelX, .Top + 2 * Screen.TwipsPerPixelY, .Width - 4 * Screen.TwipsPerPixelX, .Height - 4 * Screen.TwipsPerPixelY
      End With
   Else
      lStyle = GetWindowLong(lhWnd, GWL_STYLE)
      If ((lStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Or ((lStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then
         ' An API combo box of some sort.
         ' Assume WM_COMMAND notifications go to the parent form:
         m_hWndParent = objthis.Parent.hWnd
         m_bCombo = True
      Else
         ' Assume a standard VB control with no border.
         ' Shrink...
         With objthis
            .Move .Left + 2 * Screen.TwipsPerPixelX, .Top + 2 * Screen.TwipsPerPixelY, .Width - 4 * Screen.TwipsPerPixelX, .Height - 4 * Screen.TwipsPerPixelY
         End With
      End If
   End If
   
   pAttach lhWnd
   
End Sub

Private Sub RemoveTheme(ByVal hWnd As Long)
   On Error Resume Next
   Dim lR As Long
   lR = SetWindowTheme(hWnd, StrPtr(" "), StrPtr(" "))
   On Error GoTo 0
End Sub

Private Sub pAttach(ByRef hWndA As Long)
Dim lStyle As Long
   
   m_hWnd = hWndA
   If (m_hWnd <> 0) Then
      lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      If (lStyle And CBS_DROPDOWN) = CBS_DROPDOWN Then
         m_hWndEdit = GetWindow(m_hWnd, GW_CHILD)
      End If
      AttachMessage Me, m_hWnd, WM_PAINT
      AttachMessage Me, m_hWnd, WM_MOUSEACTIVATE
      AttachMessage Me, m_hWnd, WM_SETFOCUS
      AttachMessage Me, m_hWnd, WM_KILLFOCUS
      AttachMessage Me, m_hWnd, WM_MOUSEMOVE
      AttachMessage Me, m_hWnd, WM_TIMER
      If (m_hWndEdit <> 0) Then
         AttachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
         AttachMessage Me, m_hWndEdit, WM_SETFOCUS
         AttachMessage Me, m_hWndEdit, WM_KILLFOCUS
         AttachMessage Me, m_hWndEdit, WM_MOUSEMOVE
      End If
      If (m_bCombo) Then
         AttachMessage Me, m_hWndParent, WM_COMMAND
         RemoveTheme m_hWnd
         If (m_hWndEdit <> 0) Then
            RedrawWindow m_hWndEdit, ByVal 0&, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_INTERNALPAINT
         End If
      End If
      m_bSubclass = True
   End If
End Sub
Private Sub pRelease()
   If (m_bSubclass) Then
      DetachMessage Me, m_hWnd, WM_PAINT
      DetachMessage Me, m_hWnd, WM_SETFOCUS
      DetachMessage Me, m_hWnd, WM_KILLFOCUS
      DetachMessage Me, m_hWnd, WM_MOUSEACTIVATE
      DetachMessage Me, m_hWnd, WM_MOUSEMOVE
      DetachMessage Me, m_hWnd, WM_TIMER
      If (m_hWndEdit <> 0) Then
         DetachMessage Me, m_hWndEdit, WM_MOUSEACTIVATE
         DetachMessage Me, m_hWndEdit, WM_SETFOCUS
         DetachMessage Me, m_hWndEdit, WM_KILLFOCUS
         DetachMessage Me, m_hWndEdit, WM_MOUSEMOVE
      End If
      If (m_bCombo) Then
         DetachMessage Me, m_hWndParent, WM_COMMAND
      End If
   End If
   m_hWnd = 0: m_hWndEdit = 0: m_hWndParent = 0
End Sub

Private Sub Draw( _
      ByVal dwStyle As EDrawStyle, _
      clrTopLeft As OLE_COLOR, _
      clrBottomRight As OLE_COLOR _
   )
   If m_hWnd = 0 Then Exit Sub
   If (m_bCombo) Then
      DrawCombo dwStyle, clrTopLeft, clrBottomRight
   Else
      DrawEdit dwStyle, clrTopLeft, clrBottomRight
   End If
End Sub

Private Sub DrawEdit( _
      ByVal dwStyle As EDrawStyle, _
      clrTopLeft As OLE_COLOR, _
      clrBottomRight As OLE_COLOR _
   )
Dim rcItem As RECT
Dim rcItem2 As RECT
Dim hdc As Long
Dim hWndFocus As Long
Dim tP As POINTAPI
Dim hWndP As Long
         
   ' Thanks to Peter Siebler for providing code to do this.
         
   hWndP = GetParent(m_hWnd)
   GetWindowRect m_hWnd, rcItem
   tP.x = rcItem.Left: tP.y = rcItem.Top
   ScreenToClient hWndP, tP
   rcItem.Left = tP.x: rcItem.Top = tP.y
   tP.x = rcItem.Right: tP.y = rcItem.Bottom
   ScreenToClient hWndP, tP
   rcItem.Right = tP.x: rcItem.Bottom = tP.y
   InflateRect rcItem, 2, 2
   hdc = GetDC(hWndP)
      
   If (m_eStyle = eFlatStyleOffice9) Or (clrTopLeft = vbHighlight) Then
      Draw3DRect hdc, rcItem, clrTopLeft, clrBottomRight
   Else
      If (m_bText) Then
         Draw3DRect hdc, rcItem, vbWindowBackground, vbWindowBackground
      Else
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      End If
   End If
   LSet rcItem2 = rcItem
   InflateRect rcItem, -1, -1
        
   If (IsWindowEnabled(m_hWnd) = 0) Then
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem, vb3DHighlight, vb3DHighlight
      Else
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      End If
   Else
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      Else
         If (m_bText) Then
            Draw3DRect hdc, rcItem, vbWindowBackground, vbWindowBackground
         Else
            Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
         End If
      End If
   End If
   
   If (IsWindowEnabled(m_hWnd) = 0) Then
      If (m_eStyle > eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem2, vb3DShadow, vb3DShadow
      End If
      ReleaseDC m_hWnd, hdc
      Exit Sub
   End If

   Select Case dwStyle
   Case FC_DRAWNORMAL
'      rcItem.Top = rcItem.Top - 1
'      rcItem.Bottom = rcItem.Bottom + 1
'      Draw3DRect hDC, rcItem, vb3DHighlight, vb3DHighlight
'      rcItem.Left = rcItem.Left - 1
'      rcItem.Right = rcItem.Right
'      Draw3DRect hDC, rcItem, vbWindowBackground, vbButtonShadow
'   Case FC_DRAWRAISED, FC_DRAWPRESSED
'      InflateRect rcItem, -1, -1
'      Draw3DRect hdc, rcItem, vb3DHighlight, vb3DHighlight
'      InflateRect rcItem, -1, -1
'      Draw3DRect hdc, rcItem, vb3DHighlight, vb3DHighlight
'      InflateRect rcItem, -1, -1
'      Draw3DRect hdc, rcItem, vb3DHighlight, vb3DHighlight
   'Case FC_DRAWPRESSED
   '   rcItem.Top = rcItem.Top - 1
   '   rcItem.Bottom = rcItem.Bottom
   '   Draw3DRect hDC, rcItem, vbButtonShadow, vb3DHighlight
   End Select
   ReleaseDC m_hWnd, hdc  'ReleaseDC(hDC);
   
End Sub
Private Function Draw3DRect( _
      ByVal hdc As Long, _
      ByRef rcItem As RECT, _
      ByVal oTopLeftColor As OLE_COLOR, _
      ByVal oBottomRightColor As OLE_COLOR _
   )
Dim hPen As Long
Dim hPenOld As Long
Dim tP As POINTAPI
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, rcItem.Left, rcItem.Bottom - 1, tP
   LineTo hdc, rcItem.Left, rcItem.Top
   LineTo hdc, rcItem.Right - 1, rcItem.Top
   SelectObject hdc, hPenOld
   DeleteObject hPen
   If (rcItem.Left <> rcItem.Right) Then
      hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
      hPenOld = SelectObject(hdc, hPen)
      LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
      LineTo hdc, rcItem.Left, rcItem.Bottom - 1
      SelectObject hdc, hPenOld
      DeleteObject hPen
   End If
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Function IsRightToLeft(ByVal hWnd As Long) As Boolean
Dim lStyle As Long
   lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   IsRightToLeft = ((lStyle And WS_EX_RIGHT) = WS_EX_RIGHT) Or _
      ((lStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR)
End Function

Private Sub DrawCombo( _
      ByVal dwStyle As EDrawStyle, _
      clrTopLeft As OLE_COLOR, _
      clrBottomRight As OLE_COLOR _
   )
Dim rcItem As RECT
Dim rcWork As RECT
Dim rcButton As RECT
Dim hdc As Long
Dim hWndFocus As Long
Dim tP As POINTAPI
Dim hBr As Long
Dim bRightToLeft As Long
   
   GetClientRect m_hWnd, rcItem
   hdc = GetDC(m_hWnd)
   
   If (IsWindowEnabled(m_hWnd) = 0) Then
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      Else
         Draw3DRect hdc, rcItem, vb3DShadow, vb3DShadow
      End If
      InflateRect rcItem, -1, -1
      
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem, vb3DHighlight, vb3DHighlight
      Else
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      End If
   
   Else
      Draw3DRect hdc, rcItem, clrTopLeft, clrBottomRight
      InflateRect rcItem, -1, -1
      
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw3DRect hdc, rcItem, vbButtonFace, vbButtonFace
      Else
         Draw3DRect hdc, rcItem, vbWindowBackground, vbWindowBackground
      End If
   End If

   If (m_eStyle = eFlatStyleOffice9) Then
      '// Cover up dark 3D shadow on drop arrow.
      LSet rcButton = rcItem
      InflateRect rcButton, -1, -1
      If (IsRightToLeft(m_hWnd)) Then
         rcButton.Right = GetSystemMetrics(SM_CXHTHUMB)
      Else
         rcButton.Left = rcButton.Right - GetSystemMetrics(SM_CXHTHUMB)
      End If
      Draw3DRect hdc, rcButton, vbButtonFace, vbButtonFace
      
      '// Cover up normal 3D shadow on drop arrow.
      InflateRect rcButton, -1, -1
      Draw3DRect hdc, rcButton, vbButtonFace, vbButtonFace
      
      If Not (IsWindowEnabled(m_hWnd) = 0) Then
         Select Case dwStyle
         Case FC_DRAWNORMAL
            rcButton.Top = rcButton.Top - 1
            rcButton.Bottom = rcButton.Bottom + 1
            Draw3DRect hdc, rcButton, vb3DHighlight, vb3DHighlight
            rcButton.Left = rcButton.Left - 1
            rcButton.Right = rcButton.Left
            Draw3DRect hdc, rcButton, vbWindowBackground, &H0
      
         Case FC_DRAWRAISED
            rcButton.Top = rcButton.Top - 1
            rcButton.Bottom = rcButton.Bottom + 1
            rcButton.Right = rcButton.Right + 1
            Draw3DRect hdc, rcButton, vb3DHighlight, vbButtonShadow
      
         Case FC_DRAWPRESSED
            rcButton.Left = rcButton.Left - 1
            rcButton.Top = rcButton.Top - 2
            OffsetRect rcButton, 1, 1
            Draw3DRect hdc, rcButton, vbButtonShadow, vb3DHighlight
         End Select
      End If
   
   Else
      If (IsWindowEnabled(m_hWnd) = 0) Then
         
         LSet rcButton = rcItem
         InflateRect rcButton, -1, -1
            bRightToLeft = IsRightToLeft(m_hWnd)
         If (bRightToLeft) Then
            rcButton.Right = GetSystemMetrics(SM_CXHTHUMB)
         Else
            rcButton.Left = rcButton.Right - GetSystemMetrics(SM_CXHTHUMB)
         End If

         hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
         FillRect hdc, rcButton, hBr
         DeleteObject hBr
               
         DrawComboDropDownGlyph hdc, rcButton, vb3DShadow
         
      Else
         LSet rcButton = rcItem
         bRightToLeft = IsRightToLeft(m_hWnd)
         If (bRightToLeft) Then
            rcButton.Right = GetSystemMetrics(SM_CXHTHUMB)
         Else
            rcButton.Left = rcButton.Right - GetSystemMetrics(SM_CXHTHUMB)
         End If
         
         If (dwStyle = FC_DRAWNORMAL) And (clrTopLeft <> vbHighlight) Then
            hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
         ElseIf (dwStyle = FC_DRAWPRESSED) Then
            hBr = CreateSolidBrush(VSNetPressedColor)
         Else
            hBr = CreateSolidBrush(VSNetSelectionColor)
         End If
         FillRect hdc, rcButton, hBr
         DeleteObject hBr
         
         LSet rcWork = rcButton
         If (bRightToLeft) Then
            rcWork.Right = rcButton.Right
            rcWork.Left = rcWork.Right
         Else
            rcWork.Left = rcButton.Left
            rcWork.Right = rcWork.Left
         End If
         If (dwStyle = FC_DRAWNORMAL) And (clrTopLeft <> vbHighlight) Then
            Draw3DRect hdc, rcWork, vbWindowBackground, vbWindowBackground
         Else
            Draw3DRect hdc, rcWork, vbHighlight, vbHighlight
         End If
         If (bRightToLeft) Then
            rcWork.Right = rcWork.Right + 1
            rcWork.Left = rcWork.Right
         Else
            rcWork.Left = rcWork.Left - 1
            rcWork.Right = rcWork.Left
         End If
         Draw3DRect hdc, rcWork, vbWindowBackground, vbWindowBackground
         DrawComboDropDownGlyph hdc, rcButton, vbWindowText
      End If
   End If
   
   ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawComboDropDownGlyph( _
      ByVal hdc As Long, _
      rcButton As RECT, _
      ByVal oColor As OLE_COLOR _
   )
Dim hPen As Long
Dim hPenOld As Long
Dim xC As Long
Dim yC As Long
Dim tJ As POINTAPI
   
   xC = rcButton.Left + (rcButton.Right - rcButton.Left) \ 2
   yC = rcButton.Top + (rcButton.Bottom - rcButton.Top) \ 2
   
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oColor))
   hPenOld = SelectObject(hdc, hPen)
   
   MoveToEx hdc, xC - 2, yC - 1, tJ
   LineTo hdc, xC + 3, yC - 1
   MoveToEx hdc, xC - 1, yC, tJ
   LineTo hdc, xC + 2, yC
   MoveToEx hdc, xC, yC - 1, tJ
   LineTo hdc, xC, yC + 2
   
   SelectObject hdc, hPenOld
   DeleteObject hPen
   
End Sub

Public Property Get DroppedDown() As Boolean
   If (m_bCombo) And (m_hWnd <> 0) Then
      DroppedDown = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
   End If
End Property

Private Sub OnPaint(ByVal bFocus As Boolean, ByVal bDropped As Boolean)
 'used for paint
   If bFocus Then
      Dim clrTopLeft As Long
      Dim clrBottomRight As Long
      If (m_eStyle = eFlatStyleOffice9) Then
         clrTopLeft = vbButtonShadow
         clrBottomRight = vb3DHighlight
      Else
         clrTopLeft = vbHighlight
         clrBottomRight = vbHighlight
      End If
      If (bDropped) Then
         Draw FC_DRAWPRESSED, clrTopLeft, clrBottomRight
      Else
         Draw FC_DRAWRAISED, clrTopLeft, clrBottomRight
      End If
   Else
      If (m_eStyle = eFlatStyleOffice9) Then
         Draw FC_DRAWNORMAL, vbButtonFace, vbButtonFace
      Else
         Draw FC_DRAWNORMAL, vbWindowBackground, vbWindowBackground
      End If
   End If
   
End Sub


Private Sub Class_Initialize()
   m_eStyle = eFlatStyleOffice10
End Sub

Private Sub Class_Terminate()
   pRelease
End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)
Dim bOver As Boolean
Dim rcItem As RECT
Dim tP As POINTAPI
   
   If (bCheckMouse) Then
      bOver = True
      GetCursorPos tP
      GetWindowRect m_hWnd, rcItem
      If (PtInRect(rcItem, tP.x, tP.y) = 0) Then
         bOver = False
      End If
   End If
   
   If Not (bOver) Then
      KillTimer m_hWnd, 1
      m_bMouseOver = False
   End If

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   ' intentionally blank
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If (CurrentMessage = WM_PAINT) Then
      ISubclass_MsgResponse = emrPreprocess
   Else
      ISubclass_MsgResponse = emrPostProcess
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bDown As Boolean
Dim bFocus As Boolean

   Select Case iMsg
   Case WM_COMMAND
      If (m_hWnd = lParam) Then
         ' Type of notification is in the hiword of wParam:
         Select Case wParam \ &H10000
         Case CBN_CLOSEUP
            OnPaint (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown), bDown
         End Select
         OnTimer False
      End If
      
   Case WM_PAINT
      bDown = DroppedDown()
      bFocus = (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown)
      OnPaint (bFocus), bDown
      If (bFocus) Then
         OnTimer False
      End If
      
   Case WM_SETFOCUS
      OnPaint True, False
      OnTimer False
      
   Case WM_KILLFOCUS
      OnPaint False, False

   Case WM_MOUSEMOVE
        Screen.MousePointer = vbDefault
        
      If Not (m_bMouseOver) Then
         bDown = DroppedDown()
         If Not (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown) Then
            OnPaint True, False
            m_bMouseOver = True
            ' Start checking to see if mouse is no longer over.
            SetTimer m_hWnd, 1, 10, 0
         End If
      End If
      
   Case WM_TIMER
      OnTimer True
      If Not (m_bMouseOver) Then
         OnPaint False, False
      End If
      
   End Select
   
End Function

Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
      ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
      ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255) _
      )
      
End Property

Private Property Get VSNetControlColor() As Long
   VSNetControlColor = BlendColor(vbButtonFace, VSNetBackgroundColor, 195)
End Property

Private Property Get VSNetBackgroundColor() As Long
   VSNetBackgroundColor = BlendColor(vbWindowBackground, vbButtonFace, 220)
End Property
Private Property Get VSNetCheckedColor() As Long
   VSNetCheckedColor = BlendColor(vbHighlight, vbWindowBackground, 30)
End Property
Private Property Get VSNetBorderColor() As Long
   VSNetBorderColor = TranslateColor(vbHighlight)
End Property
Private Property Get VSNetSelectionColor() As Long
   VSNetSelectionColor = BlendColor(vbHighlight, vbWindowBackground, 70)
End Property
Private Property Get VSNetPressedColor() As Long
   VSNetPressedColor = BlendColor(vbHighlight, VSNetSelectionColor, 70)
End Property


