Attribute VB_Name = "MSubclass"
Option Explicit
'-- declares:
Private Const GWL_WNDPROC As Long = (-4)
Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080                          '-- WindowProc
    eeCantSubclass                                    '-- Can't subclass window
    eeAlreadyAttached                                 '-- Message already handled by another class
    eeInvalidWindow                                   '-- Invalid window
    eeNoExternalWindow                                '-- Can't modify external window
End Enum
#If False Then
    Private eeBaseWindowProc, eeCantSubclass, eeAlreadyAttached, eeInvalidWindow, eeNoExternalWindow
#End If
Private m_iCurrentMessage As Long
Private m_iProcOld As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, _
                                                                ByVal lpString As String, _
                                                                ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, _
                                                                      ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal MSG As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
                                                                lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Public Sub AttachMessage(iwp As ISubclass, _
                         ByVal lngHwnd As Long, _
                         ByVal iMsg As Long)

    Dim procOld As Long
    Dim f      As Long
    Dim c      As Long
    Dim iC     As Long
    Dim bFail  As Boolean

    '-- Validate window

    If IsWindow(lngHwnd) = False Then
        ErrRaise eeInvalidWindow
    End If
    If IsWindowLocal(lngHwnd) = False Then
        ErrRaise eeNoExternalWindow
    End If
    '-- Get the message count
    c = GetProp(lngHwnd, "C" & lngHwnd)
    If c = 0 Then
        '-- Subclass window by installing window procecure
        procOld = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WindowProc)
        If procOld = 0 Then
            ErrRaise eeCantSubclass
        End If
        '-- Associate old procedure with handle
        f = SetProp(lngHwnd, lngHwnd, procOld)
        Debug.Assert f <> 0
        '-- Count this message
        c = 1
        f = SetProp(lngHwnd, "C" & lngHwnd, c)
    Else                                              'NOT C...
        '-- Count this message
        c = c + 1
        f = SetProp(lngHwnd, "C" & lngHwnd, c)
    End If
    Debug.Assert f <> 0
    '-- SPM - in this version I am allowing more than one class to
    '-- make a subclass to the same hWnd and Msg.  Why am I doing
    '-- this?  Well say the class in question is a control, and it
    '-- wants to subclass its container.  In this case, we want
    '-- all instances of the control on the form to receive the
    '-- form notification message.
    c = GetProp(lngHwnd, lngHwnd & "#" & iMsg & "C")
    If c > 0 Then
        For iC = 1 To c
            If (GetProp(lngHwnd, lngHwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                ErrRaise eeAlreadyAttached
                bFail = True
                Exit For                              'loop varying ic
            End If
        Next iC
    End If
    If Not (bFail) Then
        c = c + 1
        '-- Increase count for hWnd/Msg:
        f = SetProp(lngHwnd, lngHwnd & "#" & iMsg & "C", c)
        Debug.Assert f <> 0
        '-- Associate object with message at the count:
        f = SetProp(lngHwnd, lngHwnd & "#" & iMsg & "#" & c, ObjPtr(iwp))
        Debug.Assert f <> 0
    End If

End Sub

Public Property Get CurrentMessage() As Long

    CurrentMessage = m_iCurrentMessage

End Property

Public Sub DetachMessage(iwp As ISubclass, _
                         ByVal lngHwnd As Long, _
                         ByVal iMsg As Long)

    Dim procOld As Long

    'Dim f       As Long
    Dim c      As Long
    Dim iC     As Long
    Dim iP     As Long
    Dim LPTR   As Long
    '-- Get the message count

    c = GetProp(lngHwnd, "C" & lngHwnd)
    If c = 1 Then
        '-- This is the last message, so unsubclass
        procOld = GetProp(lngHwnd, lngHwnd)
        Debug.Assert procOld <> 0
        '-- Unsubclass by reassigning old window procedure
        SetWindowLong lngHwnd, GWL_WNDPROC, procOld
        '-- Remove unneeded handle (oldProc)
        RemoveProp lngHwnd, lngHwnd
        '-- Remove unneeded count
        RemoveProp lngHwnd, "C" & lngHwnd
    Else                                              'NOT C...
        '-- Uncount this message
        c = GetProp(lngHwnd, "C" & lngHwnd)
        c = c - 1
        SetProp lngHwnd, "C" & lngHwnd, c
    End If
    '-- SPM - in this version I am allowing more than one class to
    '-- make a subclass to the same hWnd and Msg.  Why am I doing
    '-- this?  Well say the class in question is a control, and it
    '-- wants to subclass its container.  In this case, we want
    '-- all instances of the control on the form to receive the
    '-- form notification message.
    '-- How many instances attached to this hwnd/msg?
    c = GetProp(lngHwnd, lngHwnd & "#" & iMsg & "C")
    If c > 0 Then
        '-- Find this iwp object amongst the items:
        For iC = 1 To c
            If (GetProp(lngHwnd, lngHwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                iP = iC
                Exit For                              'loop varying ic
            End If
        Next iC
        If iP <> 0 Then
            '-- Remove this item:
            For iC = iP + 1 To c
                LPTR = GetProp(lngHwnd, lngHwnd & "#" & iMsg & "#" & iC)
                SetProp lngHwnd, lngHwnd & "#" & iMsg & "#" & (iC - 1), LPTR
            Next iC
        End If
        '-- Decrement the count
        RemoveProp lngHwnd, lngHwnd & "#" & iMsg & "#" & c
        c = c - 1
        SetProp lngHwnd, lngHwnd & "#" & iMsg & "C", c
    End If

End Sub

Private Sub ErrRaise(ByVal e As Long)

    Dim sText  As String
    Dim sSource As String

    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
            Case eeCantSubclass
                sText = "Can't subclass window"
            Case eeAlreadyAttached
                sText = "Message already handled by another class"
            Case eeInvalidWindow
                sText = "Invalid window"
            Case eeNoExternalWindow
                sText = "Can't modify external window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else                                              'NOT E...
        '-- Raise standard Visual Basic error
        Err.Raise e, sSource
    End If

End Sub

Private Function IsWindowLocal(ByVal lngHwnd As Long) As Boolean

    Dim idWnd  As Long

    '-- Cheat! Cut and paste from MWinTool rather than reusing
    '-- file because reusing file would cause many unneeded dependencies

    GetWindowThreadProcessId lngHwnd, idWnd
    IsWindowLocal = (idWnd = GetCurrentProcessId())

End Function

Private Function WindowProc(ByVal lngHwnd As Long, _
                            ByVal iMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    Dim procOld As Long
    Dim pSubclass As Long
    Dim iwp    As ISubclass
    Dim iwpT   As ISubclass
    Dim iPC    As Long
    Dim iP     As Long
    Dim bNoProcess As Long
    Dim bCalled As Boolean

    '
    '-- Get the old procedure from the window

    procOld = GetProp(lngHwnd, lngHwnd)
    Debug.Assert procOld <> 0
    '-- SPM - in this version I am allowing more than one class to
    '-- make a subclass to the same hWnd and Msg.  Why am I doing
    '-- this?  Well say the class in question is a control, and it
    '-- wants to subclass its container.  In this case, we want
    '-- all instances of the control on the form to receive the
    '-- form notification message.
    '-- Get the number of instances for this msg/hwnd:
    bCalled = False
    iPC = GetProp(lngHwnd, lngHwnd & "#" & iMsg & "C")
    If iPC > 0 Then
        '-- For each instance attached to this msg/hwnd, call the subclass:
        For iP = 1 To iPC
            bNoProcess = False
            '-- Get the object pointer from the message
            pSubclass = GetProp(lngHwnd, lngHwnd & "#" & iMsg & "#" & iP)
            If pSubclass = 0 Then
                '-- This message not handled, so pass on to old procedure
                WindowProc = CallWindowProc(procOld, lngHwnd, iMsg, wParam, ByVal lParam)
                bNoProcess = True
            End If
            If Not (bNoProcess) Then
                '-- Turn the pointer into an illegal, uncounted interface
                CopyMemory iwpT, pSubclass, 4
                '-- Do NOT hit the End button here! You will crash!
                '-- Assign to legal reference
                Set iwp = iwpT
                '-- Still do NOT hit the End button here! You will still crash!
                '-- Destroy the illegal reference
                CopyMemory iwpT, 0&, 4
                '-- OK, hit the End button if you must--you'll probably still crash,
                '-- but it will be because of the subclass, not the uncounted reference
                '-- Store the current message, so the client can check it:
                m_iCurrentMessage = iMsg
                m_iProcOld = procOld
                '-- Use the interface to call back to the class
                With iwp
                    '-- Preprocess (only check this the first time around):
                    If iP = 1 Then
                        If .MsgResponse = emrPreprocess Then
                            If Not (bCalled) Then
                                WindowProc = CallWindowProc(procOld, lngHwnd, iMsg, wParam, ByVal lParam)
                                bCalled = True
                            End If
                        End If
                    End If
                    '-- Consume (this message is always passed to all control
                    '-- instances regardless of whether any single one of them
                    '-- requests to consume it):
                    WindowProc = .WindowProc(lngHwnd, iMsg, wParam, ByVal lParam)
                    '-- PostProcess (only check this the last time around):
                    If iP = iPC Then
                        If .MsgResponse = emrPostProcess Then
                            If Not (bCalled) Then
                                WindowProc = CallWindowProc(procOld, lngHwnd, iMsg, wParam, ByVal lParam)
                                bCalled = True
                            End If
                        End If
                    End If
                End With                              'IWP
            End If
        Next iP
    Else                                              'NOT IPC...
        '-- This message not handled, so pass on to old procedure
        WindowProc = CallWindowProc(procOld, lngHwnd, iMsg, wParam, ByVal lParam)
    End If

End Function

''

Public Function CallOldWindowProc(ByVal lngHwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    CallOldWindowProc = CallWindowProc(m_iProcOld, lngHwnd, iMsg, wParam, lParam)

End Function
''

