VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'==================================================================================================
'cHook - module-less, IDE safe, machine code hooking thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20030901 Changes to allow some global, system-wide hooks....................................
'v1.02 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.03 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               Made InIDE public..................................................................
'
'==================================================================================================

Option Explicit

Private nAddrHook As Long                             'Address of the hook thunk
Private hHook  As Long                                'Hook handle

'============================================
'Class creation/destruction
'============================================

'Build the hook thunk into allocated memory

Private Sub Class_Initialize()

    Const PATCH_01 As Long = 17                       'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_03 As Long = 71                       'Relative address of UnhookWindowsHookEx
    Const PATCH_05 As Long = 101                      'Relative address of CallNextHookEx
    Const FUNC_EBM As String = "EbMode"               'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_UWH As String = "UnhookWindowsHookEx"  'We use CallWindowProc to call the original WndProc
    Const FUNC_CNH As String = "CallNextHookEx"       'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_VBA5 As String = "vba5"                 'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                 'Location of the EbMode function if running VB6
    Const MOD_USER As String = "user32"               'Location of the KillTimer function
    Dim I      As Long                                'Loop index
    Dim nLen   As Long                                'String lengths
    Dim sHex   As String                              'Hex code string
    Dim sCode  As String                              'Binary code string

'Store the hex pair machine code representation in sHex

    sHex = "5589E583C4F831D28955FC8955F8EB0EE8xxxxx01x83F802742085C07423E82A000000837DF800750AE828000000E83A0000008B45FCC9C20C00E817000000EBF268xxxxx02xE8xxxxx03xEBE631D24AE820000000C3FF7510FF750CFF750868xxxxx04xE8xxxxx05x8945FCC331D2E801000000C38D4510508D450C508D4508508D45FC508D45F85052B8xxxxx06x508B00FF501CC3"
    nLen = Len(sHex)                                  'Length of hex pair string

    'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
    For I = 1 To nLen Step 2                          'For each pair of hex characters
        sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, I, 2)))    'Convert a pair of hex characters to a byte and append to the ASCII string
    Next I                                            'Next pair

    nLen = LenB(sCode)                                'Get the machine code length
    nAddrHook = GlobalAlloc(0, nLen)                  'Allocate fixed memory for machine code buffer

    'Copy the code to allocated memory
    Call WinSubHook2.CopyMemory(ByVal nAddrHook, ByVal StrPtr(sCode), nLen)

    If InIDE Then
        'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
        Call WinSubHook2.CopyMemory(ByVal nAddrHook + 14, &H9090, 2)

        I = AddrFunc(MOD_VBA6, FUNC_EBM)              'Get the address of EbMode in vba6.dll
        If I = 0 Then                                 'Found?
            I = AddrFunc(MOD_VBA5, FUNC_EBM)          'VB5 perhaps, try vba5.dll
        End If

        Debug.Assert I                                'Ensure the EbMode function was found
        Call PatchRel(PATCH_01, I)                    'Patch the relative address to the EbMode api function
    End If

    'Patch the the runtime values that are known
    Call PatchRel(PATCH_03, AddrFunc(MOD_USER, FUNC_UWH))    'Relative address of UnhookWindowsHookEx
    Call PatchRel(PATCH_05, AddrFunc(MOD_USER, FUNC_CNH))    'Relative address of CallNextHookEx

End Sub

'Unhook if required and release the allocated memory

Private Sub Class_Terminate()

    Call UnHook                                       'UnHook if the hook thunk is active
    Call GlobalFree(nAddrHook)                        'Release the allocated memory

End Sub

'============================================
'Public interface
'============================================

'Set the hook

Public Function Hook(ByVal Owner As WinSubHook2.iHook, ByVal HookType As WinSubHook2.eHookType, Optional ByVal bThread As Boolean = True) As Boolean

    Const PATCH_02 As Long = 66                       'Hook handle for UnhookWindowsHookEx
    Const PATCH_04 As Long = 96                       'Hook handle for CallNextHookEx
    Const PATCH_06 As Long = 139                      'Address of the owner object
    Dim nThreadID As Long                             'App.ThreadID

    If hHook = 0 Then
        If bThread Then                               'Validate the parameters with regard to hook type vs thread or system hooking

            Select Case HookType
                Case WH_JOURNALPLAYBACK, WH_JOURNALRECORD, WH_SYSMSGFILTER
                    Debug.Assert False                'Inapropriate thread hooks
            End Select

            nThreadID = App.ThreadID
        End If

        Call PatchVal(PATCH_06, ObjPtr(Owner))        'Owner object address

        'Set the hook
        hHook = WinSubHook2.SetWindowsHookEx(HookType, nAddrHook, App.hInstance, nThreadID)

        If hHook <> 0 Then
            Hook = True
            Call PatchVal(PATCH_02, hHook)            'Hook handle for UnhookWindowsHookEx
            Call PatchVal(PATCH_04, hHook)            'Hook handle for CallNextHookEx
        End If
    End If

    Debug.Assert Hook

End Function

'Return whether we're running in the IDE. Public for general utility purposes

Public Function InIDE() As Boolean

    Debug.Assert SetTrue(InIDE)

End Function

'Call this method to unhook

Public Function UnHook() As Boolean

    If hHook <> 0 Then
        If WinSubHook2.UnhookWindowsHookEx(hHook) <> 0 Then    'Unhook the hook
            UnHook = True                             'Success
            hHook = 0                                 'Hook inactive
        End If
    End If

End Function

'============================================
'lParam cast helper routines
'============================================

Public Property Get xCBT_CREATEWND(ByVal lParam As Long) As WinSubHook2.tCBT_CREATEWND

    Call WinSubHook2.CopyMemory(xCBT_CREATEWND, ByVal lParam, LenB(xCBT_CREATEWND))

End Property

Public Property Get xCREATESTRUCT(ByVal lParam As Long) As WinSubHook2.tCREATESTRUCT

    Call WinSubHook2.CopyMemory(xCREATESTRUCT, ByVal lParam, LenB(xCREATESTRUCT))

End Property

Public Property Get xCWPSTRUCT(ByVal lParam As Long) As WinSubHook2.tCWPSTRUCT

    Call WinSubHook2.CopyMemory(xCWPSTRUCT, ByVal lParam, LenB(xCWPSTRUCT))

End Property

Public Property Get xCWPRETSTRUCT(ByVal lParam As Long) As WinSubHook2.tCWPRETSTRUCT

    Call WinSubHook2.CopyMemory(xCWPRETSTRUCT, ByVal lParam, LenB(xCWPRETSTRUCT))

End Property

Public Property Get xDEBUGSTRUCT(ByVal lParam As Long) As WinSubHook2.tDEBUGHOOKINFO

    Call WinSubHook2.CopyMemory(xDEBUGSTRUCT, ByVal lParam, LenB(xDEBUGSTRUCT))

End Property

Public Property Get xEVENTMSG(ByVal lParam As Long) As WinSubHook2.tEVENTMSG

    Call WinSubHook2.CopyMemory(xEVENTMSG, ByVal lParam, LenB(xEVENTMSG))

End Property

Public Property Get xKBDLLHOOKSTRUCT(ByVal lParam As Long) As WinSubHook2.tKBDLLHOOKSTRUCT

    Call WinSubHook2.CopyMemory(xKBDLLHOOKSTRUCT, ByVal lParam, LenB(xKBDLLHOOKSTRUCT))

End Property

Public Property Get xMOUSEHOOKSTRUCT(ByVal lParam As Long) As WinSubHook2.tMOUSEHOOKSTRUCT

    Call WinSubHook2.CopyMemory(xMOUSEHOOKSTRUCT, ByVal lParam, LenB(xMOUSEHOOKSTRUCT))

End Property

Public Property Get xMSG(ByVal lParam As Long) As WinSubHook2.tMSG

    Call WinSubHook2.CopyMemory(xMSG, ByVal lParam, LenB(xMSG))

End Property

Public Property Get xMSLLHOOKSTRUCT(ByVal lParam As Long) As WinSubHook2.tMSLLHOOKSTRUCT

    Call WinSubHook2.CopyMemory(xMSLLHOOKSTRUCT, ByVal lParam, LenB(xMSLLHOOKSTRUCT))

End Property

Public Property Get xRECT(ByVal lParam As Long) As WinSubHook2.tRect

    Call WinSubHook2.CopyMemory(xRECT, ByVal lParam, LenB(xRECT))

End Property

'============================================
'Private interface
'============================================

'Return the address of the passed function in the passed dll

Private Function AddrFunc(ByVal sDLL As String, _
                          ByVal sProc As String) As Long

    AddrFunc = WinSubHook2.GetProcAddress(WinSubHook2.GetModuleHandle(sDLL), sProc)

    'You may want to comment out the following line if you're using vb5 else the EbMode
    'GetProcAddress will stop here everytime because we look in vba6.dll first
    Debug.Assert AddrFunc

End Function

'Patch the machine code buffer offset with the relative address to the target address

Private Sub PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)

    Call WinSubHook2.CopyMemory(ByVal (nAddrHook + nOffset), nTargetAddr - nAddrHook - nOffset - 4, 4)

End Sub

'Patch the machine code buffer offset with the passed value

Private Sub PatchVal(ByVal nOffset As Long, ByVal nValue As Long)

    Call WinSubHook2.CopyMemory(ByVal (nAddrHook + nOffset), nValue, 4)

End Sub

'Worker function for InIDE - will only be called whilst running in the IDE

Private Function SetTrue(bValue As Boolean) As Boolean

    SetTrue = True
    bValue = True

End Function

