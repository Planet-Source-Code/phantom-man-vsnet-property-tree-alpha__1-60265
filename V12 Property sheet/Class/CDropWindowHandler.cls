VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindowMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' WinAPI declares:
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_TOOLWINDOW = &H80&
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48&
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' Subclassing object to catch Alt-Tab
Implements ISubclass
Private Const WM_ACTIVATE = &H6
Private Const WM_KEYDOWN = &H100

Private m_lPtr As Long
Private m_hWndParent As Long
Private m_hWndForm As Long
Private m_hWndObject As Long
Private m_bShown As Boolean

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oT As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oT, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oT
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oT, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash
   ' but it will not be the uncounted reference...

End Property

Public Property Get DropDownObject() As Object
   Set DropDownObject = ObjectFromPtr(m_lPtr)
End Property

Public Sub Create(ByRef picThis As Object)
   If (m_lPtr <> 0) Then
      Destroy
   End If
   On Error Resume Next
   m_lPtr = ObjPtr(picThis)
   If (m_lPtr <> 0) Then
      With DropDownObject
         m_hWndParent = .Container.hwnd
       '  .BorderStyle = 0
         .Visible = False
      End With
   End If
   If (Err.Number <> 0) Then
      Err.Raise Err.Number, App.EXEName & ".Create", "Invalid object passed to Create"
      m_lPtr = 0
      m_hWndParent = 0
   End If
   
End Sub
Public Sub Destroy()
   If (m_hWndForm <> 0) Then
      DetachMessage Me, m_hWndForm, WM_ACTIVATE
      m_hWndForm = 0
   End If
   If (m_lPtr <> 0) Then
      SetParent DropDownObject.hwnd, m_hWndParent
      m_lPtr = 0
   End If
End Sub
Public Sub Show(ByVal x As Long, ByVal y As Long)
Dim tP As POINTAPI
Dim hWndDesktop As Long
Dim lStyle As Long
Dim lhWnd As Long
Dim lParenthWNd As Long
   
   ' Make sure the picture box won't appear in the
   ' task bar by making it into a Tool Window:
   lhWnd = DropDownObject.hwnd
   lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_TOOLWINDOW
   lStyle = lStyle And Not (WS_EX_APPWINDOW)
   SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
   
   ' Determine where to show it in Screen coordinates:
   tP.x = x: tP.y = y
   lParenthWNd = DropDownObject.Parent.hwnd
   'ClientToScreen lhWnd, tP
   
  
   
   ' Show the form:
   SetWindowPos lhWnd, lParenthWNd, tP.x, tP.y, DropDownObject.Width \ Screen.TwipsPerPixelX, DropDownObject.Height \ Screen.TwipsPerPixelY, SWP_SHOWWINDOW
   
   ' Make the picture box a child of the desktop (so
   ' it can be fully shown even if it extends beyond
   ' the form boundaries):
   DropDownObject.Move tP.x, tP.y
   SetParent lhWnd, hWndDesktop
   
   ' Tell VB it is shown:
   DropDownObject.Visible = True
   DropDownObject.ZOrder 0
   
   ' Try to set focus:
   SetFocusAPI lhWnd
   
   ' Capture all mouse messages.
   SetCapture lhWnd
   
   ' Start subclassing for Alt-tab
   m_hWndForm = lParenthWNd
   m_hWndObject = lhWnd
   AttachMessage Me, m_hWndForm, WM_ACTIVATE
   
   ' Store a flag saying we're shown:
   m_bShown = True
End Sub

Public Sub Hide()
   ' Stop subclassing for Alt-tab
   If (m_hWndForm <> 0) Then
      DetachMessage Me, m_hWndForm, WM_ACTIVATE
   End If
   If (m_hWndObject <> 0) Then
      ' Hide the picturebox:
      DropDownObject.Visible = False
   End If
   m_hWndForm = 0
   m_hWndObject = 0
   
   ' Stop capturing mouse messages:
   ReleaseCapture
         
   ' Store a flag saying we're not shown:
   m_bShown = False
   
End Sub
Public Property Get IsShown() As Boolean
   ' Return whether we are shown or not.
   IsShown = m_bShown
End Property
Public Property Get InRect(ByVal x As Single, ByVal y As Single) As Boolean
Dim tR As RECT
   If (IsShown()) Then
      GetClientRect DropDownObject.hwnd, tR
      x = x \ Screen.TwipsPerPixelX
      y = y \ Screen.TwipsPerPixelY
      If (PtInRect(tR, x, y) = 1) Then
         InRect = True
      End If
   End If
End Property
Public Sub Resize(ByVal lNewWidth As Long, ByVal lNewHeight As Long)
Dim tWR As RECT, tSR As RECT
Dim lR As Long

   ' Get the size of the window on screen:
   GetWindowRect DropDownObject.hwnd, tWR
   ' Check if it will fit:
   lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tSR, 0)
   If (lR = 0) Then
      ' Call failed - just use standard screen:
      tSR.Left = 0
      tSR.Top = 0
      tSR.Right = Screen.Width \ Screen.TwipsPerPixelX
      tSR.Bottom = Screen.Height \ Screen.TwipsPerPixelY
   End If
   If (tWR.Left + lNewWidth > tSR.Right) Then
      ' too big in x
      lNewWidth = tSR.Right - tWR.Left
   End If
   If (tWR.Top + lNewHeight > tSR.Bottom) Then
      ' too big in y
      lNewHeight = tSR.Bottom - tWR.Top
   End If
   MoveWindow DropDownObject.hwnd, tWR.Left, tWR.Top, lNewWidth, lNewHeight, 1

End Sub

Private Sub Class_Terminate()
   ' Clear up
   Destroy
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   ' NR
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ' Respond to the message after windows has done its stuff:
   ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case iMsg
   Case WM_ACTIVATE
      ' This effectively catches Alt-Tabbing:
      If (IsShown()) Then
         Hide
      End If
   Case WM_KEYDOWN
      Debug.Print "KeyDown"
   End Select
End Function


