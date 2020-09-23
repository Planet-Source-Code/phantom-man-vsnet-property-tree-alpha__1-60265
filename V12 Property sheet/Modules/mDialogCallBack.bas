Attribute VB_Name = "mDialogCallBack"
'------------------------------------------------------------------------------
'-- Module Name.....: mDialogCallBack
'-- Description.....: Module that provides the callback function for Dialoge Boxes
'--
'-- Notes...........: The callback procedure receives messages or notifications intended for
'--                   the default dialog box procedure of dialog boxes, so you can subclass
'--                   the standard controls of the common dialog box. In this implementation
'--                   we can center the standard dialog on the screen or we can set another
'--                   title. If you want to perform other changes modify the following function.
'--
'--
'-- Author, date....: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- FontDialogCallBack   Callback (global) routine for ICDLG_FontDialogHandler. It is used to center the
'--                      dialog and to set the caption text.
'-- FileOpenSaveDialogCallbackEx - Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--                                to center the dialog box. Must be used with eFileOpenSaveFlag_Explorer.
'-- FileOpenSaveDialogCallback     Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--                                to center the dialog box. Does not cover eFileOpenSaveFlag_Explorer.
'-- ColourDialogCallBack Callback (global) routine for the ICDLG_ColorDialogHandler class. It is used to
'--                      center the dialog box and to set the caption text.
'--
'-- BrowseForFolderCallBack - Callback (global) routine for the ICDLG_BrowseForFolderHandler class. It is used to
'--                           center the dialog and to set the caption text.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'------------------------------------------------------------------------------

Option Explicit

'-- Private constants - BrowseForFolder
Private Const MAX_PATH = 512
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELECTIONCHANGED = 2
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXT = BFFM_SETSTATUSTEXTA
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTION = BFFM_SETSELECTIONA

'-- Private class type definitions
Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

'-- Private class constants
Private Const WM_INITDIALOG = &H110

'-- Private class API function declarations
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long


'-- Private API function declarations - BrowseForFolder
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'-- Public variables used for communication between a ICDLG_FontDialogHandler object and
'-- the callback routine implemented in this module
Public g_bCenterFontDialog As Boolean
Public g_sFontDialogTitle As String

'-- Public variable used for communication between a ICDLG_ColorDialogHandler object and the callback routine
'-- implemented in this module
Public g_bCenterColourDialog As Boolean
Public g_sColourDialogTitle As String


'------------------------------------------------------------------------------
'-- Function    : FontDialogCallback
'-- Notes       : Callback (global) routine for ICDLG_FontDialogHandler. It is used to center the
'--               dialog and to set the caption text.
'------------------------------------------------------------------------------
Public Function FontDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight As Long
    Dim rcWidth As Long
    Dim rc     As RECT
    Dim rcDesk As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            '-- Set the new title
            If Len(Trim$(g_sFontDialogTitle)) > 0 Then SetWindowText hwnd, g_sFontDialogTitle

            '-- Center the window
            If g_bCenterFontDialog Then
                Call GetWindowRect(GetDesktopWindow, rcDesk)
                Call GetWindowRect(hwnd, rc)

                rcHeight = rc.Bottom - rc.Top
                rcWidth = rc.Right - rc.Left
                rc.Left = Abs(((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
                rc.Top = Abs(((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

                MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1
            End If
        Case Else
            '

    End Select

    FontDialogCallback = 0&
End Function

'------------------------------------------------------------------------------
'-- Function    : FileOpenSaveDialogCallbackEx
'-- Notes       : Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--               to center the dialog box. Must be used with eFileOpenSaveFlag_Explorer.
'------------------------------------------------------------------------------
Public Function FileOpenSaveDialogCallbackEx(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight As Long
    Dim rcWidth As Long
    Dim HWndParent As Long
    Dim rc     As RECT
    Dim rcDesk As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            '-- Center the window
            HWndParent = GetParent(hwnd)

            Call GetWindowRect(GetDesktopWindow, rcDesk)
            Call GetWindowRect(HWndParent, rc)

            rcHeight = rc.Bottom - rc.Top
            rcWidth = rc.Right - rc.Left
            rc.Left = (((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
            rc.Top = (((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

            MoveWindow HWndParent, rc.Left, rc.Top, rcWidth, rcHeight, 1

        Case Else
            '

    End Select

    FileOpenSaveDialogCallbackEx = 0&
End Function


'------------------------------------------------------------------------------
'-- Function    : FileOpenSaveDialogCallback
'-- Notes       : Old-style callback (global) routine for ICDLG_FileOpenSaveHandler. It is used
'--               to center the dialog box. Does not cover eFileOpenSaveFlag_Explorer.
'------------------------------------------------------------------------------
Public Function FileOpenSaveDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight As Long
    Dim rcWidth As Long
    Dim rc     As RECT
    Dim rcDesk As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            '-- Center the window
            Call GetWindowRect(GetDesktopWindow, rcDesk)
            Call GetWindowRect(hwnd, rc)

            rcHeight = rc.Bottom - rc.Top
            rcWidth = rc.Right - rc.Left
            rc.Left = (((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
            rc.Top = (((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

            MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1

        Case Else
            '

    End Select

    FileOpenSaveDialogCallback = 0&
End Function

'------------------------------------------------------------------------------
'-- Function    : ColourDialogCallback
'-- Notes       : Callback (global) routine for the ICDLG_ColorDialogHandler class. It is used to
'--               center the dialog box and to set the caption text.
'------------------------------------------------------------------------------
Public Function ColourDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim rcHeight As Long
    Dim rcWidth As Long
    Dim rc     As RECT
    Dim rcDesk As RECT

    Select Case uMsg

        Case WM_INITDIALOG
            '-- Set the new title
            If Len(Trim$(g_sColourDialogTitle)) > 0 Then SetWindowText hwnd, g_sColourDialogTitle

            '-- Center the window
            If g_bCenterColourDialog Then
                Call GetWindowRect(GetDesktopWindow, rcDesk)
                Call GetWindowRect(hwnd, rc)

                rcHeight = rc.Bottom - rc.Top
                rcWidth = rc.Right - rc.Left
                rc.Left = Abs(((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
                rc.Top = Abs(((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)

                MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1
            End If
        Case Else
            '

    End Select

    ColourDialogCallback = 0&
End Function


'------------------------------------------------------------------------------
'-- Function    : BrowseForFolderCallback
'-- Notes       : Callback (global) routine for the ICDLG_BrowseForFolderHandler class. It is used to
'--               center the dialog and to set the caption text.
'------------------------------------------------------------------------------
Public Function BrowseForFolderCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next

    Dim s      As String

    Select Case uMsg

        Case BFFM_INITIALIZED
            '-- Browse dialog box has finished initializing (lParam is NULL) so set the selection
            If lpData Then SendMessage ByVal hwnd, BFFM_SETSELECTION, ByVal True, ByVal lpData

        Case BFFM_SELECTIONCHANGED
            '-- Selection has changed (lParam contains pidl of selected folder) so display full path if status area if enabled
            s = String$(MAX_PATH, 0)
            If SHGetPathFromIDList(lParam, s) Then
                s = Left$(s, InStr(s, vbNullChar) - 1)
            Else
                s = ""
            End If
            SendMessage ByVal hwnd, BFFM_SETSTATUSTEXT, ByVal 0&, ByVal s

        Case Else:

    End Select

    BrowseForFolderCallback = 0
End Function

