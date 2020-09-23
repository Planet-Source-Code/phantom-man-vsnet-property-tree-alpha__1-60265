VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFocusWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     vbAccelerator cActiveTitleBar class
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     25 August 1999
'
' Requires: SSUBTMR.DLL
'
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Provides a class which if attached to a form's window handle will
' prevent the form's titlebar from going inactive.  Ideal for
' emulating MS Office style ToolWindows over your main form.
'
' ======================================================================================

Private Const WM_NCACTIVATE = &H86
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_ACTIVATE = &H6
Private Const WM_DESTROY = &H2
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Implements ISubclass

Private m_hWnd As Long

Public Sub Attach(ByVal hWndA As Long)
   Detach
   m_hWnd = hWndA
   AttachMessage Me, m_hWnd, WM_NCACTIVATE
   AttachMessage Me, m_hWnd, WM_ACTIVATEAPP
   AttachMessage Me, m_hWnd, WM_ACTIVATE
   AttachMessage Me, m_hWnd, WM_DESTROY

End Sub
Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_NCACTIVATE
      DetachMessage Me, m_hWnd, WM_ACTIVATEAPP
      DetachMessage Me, m_hWnd, WM_ACTIVATE
      DetachMessage Me, m_hWnd, WM_DESTROY
   End If
End Sub

Private Sub Class_Terminate()
   Detach
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   Select Case CurrentMessage
   Case WM_NCACTIVATE
      ISubclass_MsgResponse = emrConsume
   Case Else
      ISubclass_MsgResponse = emrPostProcess
   End Select
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Static iRefCount As Long
   Select Case iMsg
   Case WM_NCACTIVATE
      Debug.Print "WM_NCACTIVATE"
      If wParam = 0 Then
         iRefCount = iRefCount + 1
         If iRefCount < 2 Then
            LockWindowUpdate hWnd
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
            LockWindowUpdate 0
         Else
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
         End If
      Else
         ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
      End If
   Case WM_ACTIVATEAPP
      Debug.Print "ACTIVATEAPP"
      If (wParam = 0) Then
         iRefCount = 0
         ' app being deactivated
         CallOldWindowProc m_hWnd, WM_NCACTIVATE, 0, 0
      Else
         ' app being activated
         ' if not the active form then we should repaint
         ' the title bar
         CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
      End If
   Case WM_ACTIVATE
      If wParam = 0 Then
         iRefCount = 0
         ' deactivating the window, lParam is the window that is being activated:
         Debug.Print lParam
      End If

   Case WM_DESTROY
      ' In case the user does not set the class
      ' to nothing before the owning form is
      ' closed:
      Debug.Print "WM_DESTROY"
      Detach
   End Select
   
End Function

