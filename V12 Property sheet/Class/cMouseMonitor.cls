VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_MouseMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'API functions

Implements IWindowsHook

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Event MouseEvent(blnUp As Boolean)
Public Event KeyboardEvent()

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_SYSTEMKEYDOWN = &H104
Private Const WM_SYSTEMKEYUP = &H105
Dim mH As New IAPP_GlobHook

Public Sub StartMonitor()
        Set mH = New IAPP_GlobHook
        
      InstallHook Me, WH_JOURNALRECORD
      InstallHook Me, WH_GETMESSAGE

End Sub

Public Sub StopMonitor()

      RemoveHook Me, WH_GETMESSAGE
      RemoveHook Me, WH_JOURNALRECORD
        Set mH = Nothing
End Sub


Private Function IWindowsHook_HookProc(ByVal eType As EHTHookTypeConstants, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, bConsume As Boolean) As Long


If (eType = WH_JOURNALRECORD) Then
   
      Dim keyCode As Integer
      Dim keyAscii As Integer
      Dim sMsg As String
      Dim sMsgBit As String
      Dim lhWNd As Long
      Dim tP As POINTAPI
      
      Dim cEvent As IAPP_JournallParam
      Set cEvent = mH.JournalRecordlParam(lParam)
      
      lhWNd = cEvent.hwnd
      
      
      Select Case cEvent.Msg
      
      Case WM_KEYDOWN
         RaiseEvent KeyboardEvent
         keyCode = (cEvent.lParamLow And &HFF&)
         
      Case WM_SYSTEMKEYDOWN
         
         keyCode = (cEvent.lParamLow And &HFF&)
      RaiseEvent KeyboardEvent
      Case WM_KEYUP
         keyCode = (cEvent.lParamLow And &HFF&)
      RaiseEvent KeyboardEvent
      Case WM_SYSTEMKEYUP
         
         keyCode = (cEvent.lParamLow And &HFF&)
        RaiseEvent KeyboardEvent
        
      Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
         
         sMsg = "Mouse Down: "
         Select Case cEvent.Msg
         Case WM_LBUTTONDOWN
            sMsg = sMsg & "Left Button"
            RaiseEvent MouseEvent(False)
         Case WM_RBUTTONDOWN
            sMsg = sMsg & "Right Button"
            RaiseEvent MouseEvent(False)
         Case WM_MBUTTONDOWN
            sMsg = sMsg & "Middle Button"
            RaiseEvent MouseEvent(False)
         End Select
         
      Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
         
         Select Case cEvent.Msg
         Case WM_LBUTTONUP
            sMsg = sMsg & "Left Button"
            RaiseEvent MouseEvent(True)
         Case WM_RBUTTONUP
            sMsg = sMsg & "Right Button"
            RaiseEvent MouseEvent(True)
         Case WM_MBUTTONUP
            sMsg = sMsg & "Middle Button"
            RaiseEvent MouseEvent(True)
         End Select
         
         
      Case WM_MOUSEWHEEL
        RaiseEvent MouseEvent(False)
      Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_MBUTTONDBLCLK
         RaiseEvent MouseEvent(True)
         sMsg = "Double Click: "
         Select Case cEvent.Msg
         Case WM_LBUTTONDBLCLK
            RaiseEvent MouseEvent(True)
            
            sMsg = sMsg & "Left Button"
         Case WM_RBUTTONDBLCLK
            sMsg = sMsg & "Right Button"
            RaiseEvent MouseEvent(True)
         Case WM_MBUTTONDBLCLK
            sMsg = sMsg & "Middle Button"
            RaiseEvent MouseEvent(True)
         End Select
         
               
      End Select
               
   End If
     
    'Debug.Print sMsg

End Function
