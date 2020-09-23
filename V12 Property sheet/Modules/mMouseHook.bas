Attribute VB_Name = "mMouseHook"
Option Explicit

Private Type MOUSEHOOKSTRUCT ' 20 bytes
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long


' ----------------------------- Window Messages -----------------------------
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_COMMAND As Long = &H111
Private Const WM_DESTROY As Long = &H2
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const WM_NCMBUTTONDOWN As Long = &HA7
Private Const WM_NCRBUTTONDOWN As Long = &HA4
Private Const WM_NCXBUTTONDOWN As Long = &HAB
Private Const WM_PARENTNOTIFY As Long = &H210
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_VSCROLL = &H115
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WH_MOUSE As Long = 7      ' used with mouse hook
Dim m_oldWndProc As Long                ' previous window procedure
Dim m_parentWndProc As Long             ' parent's previous window procedure
Private hMsgHook As Long                ' message hook ID
Private bAbortFlag As Boolean           ' multipurpose flag
Public mHWND As Long

Public Sub HookMsgQueue(bSet As Boolean)
' toggle switch for hooking into the message queue
If bSet Then
    hMsgHook = SetWindowsHookEx(WH_MOUSE, AddressOf msgHookProc, App.hInstance, App.ThreadID)
Else
    If hMsgHook Then UnhookWindowsHookEx hMsgHook
    hMsgHook = 0
'    Debug.Print "unhooked mouse"
End If
End Sub

Private Function msgHookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If ncode > -1 Then
    
    Select Case wParam
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_NCXBUTTONDOWN, _
        WM_NCLBUTTONDOWN, WM_NCMBUTTONDOWN, WM_NCRBUTTONDOWN
        'If GetCapture() = 0 Then
        
            Dim mouseMsg As MOUSEHOOKSTRUCT
            CopyMemory mouseMsg, ByVal lParam, &H14
        Debug.Print mHWND = mouseMsg.hwnd
            Select Case mouseMsg.hwnd
            Case Is <> mHWND ' m_hWnd, sb_hWnd, lb_hWnd, m_ZoomHwnd
            Case Else
                HookMsgQueue False
                msgHookProc = 1     ' prevent message from continuing on
                Exit Function
            End Select
        'End If
    Case Else
    End Select

End If
msgHookProc = CallNextHookEx(hMsgHook, ncode, wParam, lParam)
End Function

