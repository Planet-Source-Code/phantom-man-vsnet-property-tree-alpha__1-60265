VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMouseTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_DESTROY = &H2&

' 1. USER32 method:
Private Const WM_MOUSEHOVER = &H2A1&
Private Const WM_MOUSELEAVE = &H2A3&

Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const TME_QUERY = &H40000000
Private Const TME_CANCEL = &H80000000

Private Const HOVER_DEFAULT = &HFFFFFFFF

Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" _
   (lpEventTrack As tagTRACKMOUSEEVENT) As Long

' 2. The COMCTL32.DLL Method:
'// Declare _TrackMouseEvent.  This API tries to use the window manager's
'// implementation of TrackMouseEvent if it is present, otherwise it emulates.
Private Declare Function CCTrackMouseEvent Lib "COMCTL32.DLL" Alias "_TrackMouseEvent" _
   (lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Const MK_LBUTTON = &H1&
Private Const MK_RBUTTON = &H2&
Private Const MK_SHIFT = &H4&
Private Const MK_CONTROL = &H8&
Private Const MK_MBUTTON = &H10&

' 3 If ALL else fails, then use the work-around:
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSEUP = &H200
Private Const WM_ACTIVATE = &H6
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONUP = &H205

' Version detection:
' For OS:
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long
   szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

' For COMCTL32.DLL
Private Const S_OK = &H0
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long


' Implementation:
Implements ISubclass

Public Event MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
Public Event MouseLeave()
Public Event MouseClickOutsideControl()

Private m_bTracking As Boolean
Private m_hWnd As Long
Private m_hWndParent As Long
Private m_bUseCC As Boolean
Private m_bUseCapture As Boolean
Public Enum EMouseTrackMethods
   eMouseTrackDetect = -1
   eMouseTrackUser32 = 0
   eMouseTrackComCtl32 = 1
   eMouseTrackWorkAround = 2
End Enum
Private m_eMethod As EMouseTrackMethods

Private Sub pDetectMethod()
Dim tVI As OSVERSIONINFO
      
   ' Default to use COMCTL32.DLL. (Requires IE4.0 or higher installed).
   m_bUseCC = True
   
   ' Now we check for a window manager (user32.dll) implementation of
   ' TrackMouseEvent.  We can rely on COMCTL32.DLL's version to use
   ' the window manager's version directly, except IE4 may not be installed.
   tVI.dwOSVersionInfoSize = Len(tVI)
   If (GetVersionEx(tVI) <> 0) Then
      ' NT4 or higher supports TrackMouseEvent in User32:
      If (tVI.dwPlatformID = VER_PLATFORM_WIN32_NT) And (tVI.dwMajorVersion > 3) Then
         ' Using NT
         m_bUseCC = False
      ' Win98 or higher supports TrackMouseEvent in User32:
      ElseIf (tVI.dwMajorVersion >= 5) Then
         ' Using 98
         m_bUseCC = False
      End If
   End If
   If (m_bUseCC) Then
      Dim hMod As Long
      Dim lR As Long
      Dim lptrDLLVersion As Long
      Dim tDVI As DLLVERSIONINFO
      Dim bCC As Boolean
      
      hMod = LoadLibrary("comctl32.dll")
      If (hMod <> 0) Then
         lR = S_OK
         '/*
         ' You must get this function explicitly because earlier versions of the DLL
         ' don't implement this function. That makes the lack of implementation of the
         ' function a version marker in itself. */
         lptrDLLVersion = GetProcAddress(hMod, "DllGetVersion")
         If (lptrDLLVersion <> 0) Then
            tDVI.cbSize = Len(tDVI)
            lR = DllGetVersion(tDVI)
            If (lR = S_OK) Then
               If (tDVI.dwMajor > 4) Then
                  bCC = True
               ElseIf (tDVI.dwMajor = 4) And (tDVI.dwMinor > 71) Then
                  bCC = True
               End If
            End If
         End If
         FreeLibrary hMod
      End If
      
      If Not (bCC) Then
         m_bUseCC = False
         m_bUseCapture = True
      End If
   End If

   If (m_bUseCC) Then
      m_eMethod = eMouseTrackComCtl32
   ElseIf (m_bUseCapture) Then
      m_eMethod = eMouseTrackWorkAround
   Else
      m_eMethod = eMouseTrackUser32
   End If

End Sub
Public Property Get Method() As EMouseTrackMethods
   Method = m_eMethod
End Property

Public Sub AttachMouseTracking( _
      objTo As Object, _
      Optional ByVal eForceMethod As EMouseTrackMethods = eMouseTrackDetect _
   )
   
   m_bUseCapture = False
   m_bUseCC = False
   
   ' Check for tracking type if not forced:
   If (eForceMethod = eMouseTrackDetect) Then
      pDetectMethod
   Else
      Select Case eForceMethod
      Case eMouseTrackWorkAround
         m_bUseCapture = True
      Case eMouseTrackComCtl32
         m_bUseCC = True
      End Select
      m_eMethod = eForceMethod
   End If
   
   ' Start subclassing for WM_MOUSEHOVER and WM_MOUSELEAVE
   ' messages:
   
   DetachMouseTracking
   m_hWnd = objTo.hwnd
   If (m_hWnd <> 0) Then
      If (m_bUseCapture) Then
         AttachMessage Me, m_hWnd, WM_MOUSEMOVE
         AttachMessage Me, m_hWnd, WM_LBUTTONUP
         AttachMessage Me, m_hWnd, WM_MBUTTONUP
         AttachMessage Me, m_hWnd, WM_RBUTTONUP
         m_hWndParent = objTo.Parent.hwnd
         AttachMessage Me, m_hWndParent, WM_ACTIVATE
      Else
         AttachMessage Me, m_hWnd, WM_MOUSEHOVER
         AttachMessage Me, m_hWnd, WM_MOUSELEAVE
      End If
      AttachMessage Me, m_hWnd, WM_DESTROY

   End If
   
End Sub


Public Sub StartMouseTracking()
Dim tET As tagTRACKMOUSEEVENT
Dim lR As Long

On Error GoTo ErrorHandler

   ' Tells Windows to start tracking the mouse over the specified
   ' hWnd:

   If Not (m_bTracking) Then
      ' Tracking will stop whenever a WM_MOUSEHOVER or WM_MOUSELEAVE
      ' event occurs.
      tET.cbSize = Len(tET)
      tET.dwFlags = TME_HOVER Or TME_LEAVE
      tET.dwHoverTime = HOVER_DEFAULT
      tET.hwndTrack = m_hWnd
      If (m_bUseCC) Then
         lR = CCTrackMouseEvent(tET)
      ElseIf (m_bUseCapture) Then
         SetCapture m_hWnd
      Else
         lR = TrackMouseEvent(tET)
      End If
      m_bTracking = True
   End If
   
   Exit Sub

ErrorHandler:
   ' This occurs because the user has forced a method
   ' which is not supported.  Raise error!
   Err.Raise Err.Number, App.EXEName & ".cMouseTrack", Err.Description
   ' But don't allow this to get set...
   m_bTracking = False
   
End Sub
Public Sub DetachMouseTracking()

   ' Stops subclassing for mouse tracking commands.
   ' Called automatically when the class terminates.
   If (m_hWnd <> 0) Then
      If (m_bUseCapture) Then
         ReleaseCapture
         DetachMessage Me, m_hWnd, WM_MOUSEMOVE
         DetachMessage Me, m_hWnd, WM_LBUTTONUP
         DetachMessage Me, m_hWnd, WM_MBUTTONUP
         DetachMessage Me, m_hWnd, WM_RBUTTONUP
         If (m_hWndParent <> 0) Then
            DetachMessage Me, m_hWndParent, WM_ACTIVATE
         End If
      Else
         DetachMessage Me, m_hWnd, WM_MOUSEHOVER
         DetachMessage Me, m_hWnd, WM_MOUSELEAVE
      End If
      DetachMessage Me, m_hWnd, WM_DESTROY
      m_hWnd = 0
   End If
   
End Sub
Public Property Get Tracking() As Boolean
   ' Returns whether windows is tracking or not (it stops
   ' everyime a WM_MOUSEHOVER or WM_MOUSELEAVE event is fired):
   Tracking = m_bTracking
End Property

Private Sub Class_Initialize()
   '
End Sub

Private Sub Class_Terminate()
   ' Clear up subclass:
   DetachMouseTracking
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   ' Let Windows pre-process message:
   ISubClass_MsgResponse = emrPreprocess
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Single, y As Single

   ' Respond to WM_MOUSEHOVER and WM_MOUSELEAVE messages:
   Select Case iMsg
   ' ===============================================
   ' To effect the TrackMouseEvent User32 or
   ' Comctl32 methods:
   Case WM_MOUSEHOVER
      Dim Button As MouseButtonConstants
      Dim Shift As ShiftConstants
      
      m_bTracking = False
      If (wParam And MK_LBUTTON) = MK_LBUTTON Then
         Button = Button Or vbLeftButton
      End If
      If (wParam And MK_RBUTTON) = MK_RBUTTON Then
         Button = Button Or vbRightButton
      End If
      If (wParam And MK_MBUTTON) = MK_MBUTTON Then
         Button = Button Or vbMiddleButton
      End If
      If (wParam And MK_CONTROL) = MK_CONTROL Then
         Shift = Shift Or vbCtrlMask
      End If
      If (wParam And MK_SHIFT) = MK_SHIFT Then
         Shift = Shift Or vbShiftMask
      End If
      x = lParam And &HFFFF&
      y = lParam \ &H10000
      RaiseEvent MouseHover(Button, Shift, x, y)
   Case WM_MOUSELEAVE
      m_bTracking = False
      RaiseEvent MouseLeave
   ' ===============================================
      
   ' ===============================================
   ' To effect the SetCapture/ReleaseCapture method:
   Case WM_MOUSEMOVE, WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
      Dim tR As RECT, tP As POINTAPI
      GetClientRect m_hWnd, tR
      x = lParam And &HFFFF&
      y = lParam \ &H10000
      tP.x = x
      tP.y = y
      ClientToScreen m_hWnd, tP
      If (PtInRect(tR, x, y) = 0) Or (WindowFromPoint(tP.x, tP.y) <> m_hWnd) Then
         If (GetAsyncKeyState(vbKeyLButton) = 0) And (GetAsyncKeyState(vbKeyMButton) = 0) And (GetAsyncKeyState(vbKeyRButton) = 0) Then
            m_bTracking = False
            ReleaseCapture
            RaiseEvent MouseLeave
            RaiseEvent MouseClickOutsideControl
         End If
      ElseIf (iMsg <> WM_MOUSEMOVE) Then
         m_bTracking = False
         StartMouseTracking
      End If
   
   Case WM_ACTIVATE
      If (m_bTracking) Then
         m_bTracking = False
         ReleaseCapture
         RaiseEvent MouseLeave
      End If
      
   Case WM_DESTROY
      'Debug.Print "WM_DESTROY"
      DetachMouseTracking
      
   End Select
   ' ===============================================
   
End Function


