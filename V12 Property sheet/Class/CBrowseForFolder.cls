VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_BOFHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'-- Class Name.....: IAPP_BrowseForFolderHandler
'-- Description....: Class for using the "BrowseForFolder" dialog box.
'--
'-- Notes..........:  Use this class to provide your application the possibility to allow to
'--                   the user to select a system folder.
'--
'-- Author, date...: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'-- Flags                EBrowseForFolderFlags - Returns or sets the flags specifying the options for the
'--                                              Browse for Folder dialog box.
'-- DialogTitle          String        Returns or sets the dialog's prompt string.
'-- InitialPath          String        Returns or sets the folder used as a pre-selected folder in the
'--                                    Browse for Folder dialog.
'-- SelectedPath         String        Returns the folder selected in the Browse for Folder dialog.
'--                                    If the dialog was canceled returns the empty string ("").
'-- HwndOwner            Long          Returns or sets a handle to the owner window for the Browse for Folder dialog box.
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- Init                 Use this routine for basic object initialization.
'-- Browse               Displays a dialog box that enables the user to select a shell folder.
'--                      Returns True is a folder is selected and the OK button was pressed, False
'--                      otherwise.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'------------------------------------------------------------------------------

Option Explicit

'-- Error handling definitions
Private Const E_ERR_BASE = 17060 + vbObjectError
Public Enum EErrBrowseForFolder
    eErrBrowseForFolder_CantGetPath = E_ERR_BASE + 1
    eErrBrowseForFolder_InvalidWindowHandle
    eErrBrowseForFolder_ComponentFailure
End Enum
Private Const S_ERR_CantGetPath = "Could not get the path of the selected folder"
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "IAPP_BrowseForFolderHandler component failure"

'-- Private class constants
Private Const MAX_LENGTH = 512
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'-- Public class enums
Public Enum EBrowseForFolderFlags
    eBrowseForFolderFlag_ReturnOnlyFSDirs = &H1       'Only returns file system directories
    eBrowseForFolderFlag_DontGoBelowDomain = &H2      'Does not include network folders below domain level
    eBrowseForFolderFlag_StatusText = &H4             'Includes status area in the dialog for use with callback
    eBrowseForFolderFlag_ReturnFSAncestors = &H8      'Only returns file system ancestors.
    eBrowseForFolderFlag_EditBox = &H10               'allows user to rename selection
    eBrowseForFolderFlag_Validate = &H20              'insist on valid editbox result (or CANCEL)
    eBrowseForFolderFlag_UseNewUI = &H40              'Version 5.0. Use the new user-interface. To use
    'this flag, you must call OleInitialize or
    'CoInitialize before calling SHBrowseForFolder.
    eBrowseForFolderFlag_BrowseForComputer = &H1000   'Only returns computers.
    eBrowseForFolderFlag_BrowseForPrinter = &H2000    'Only returns printers.
    eBrowseForFolderFlag_IncludeFiles = &H4000        'Browse for everything
End Enum

'-- Private class type definitions
Private Type BROWSEINFO
    hOwner     As Long
    pIdlRoot   As Long
    pszDisplayName As String
    lpszTitle  As String
    ulFlags    As Long
    lpfn       As Long
    lParam     As Long
    iImage     As Long
End Type

'-- Private class API function declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Private variables to hold property values
Private m_lFlags As EBrowseForFolderFlags
Private m_lHWndOwner As Long
Private m_sInitialPath As String
Private m_sSelectedPath As String
Private m_sDialogTitle As String

'-- Private variables for internal class use
Private m_biInfo As BROWSEINFO
Private m_lImage As Long
Private m_lParam As Long
Private m_lPIdlRoot As Long
Private m_lPFN As Long
Private m_sDisplayName As String

'------------------------------------------------------------------------------
'-- Sub         : Init
'-- Notes       : Use this routine for basic object initialization.
'------------------------------------------------------------------------------
Public Sub Init(Optional InitialPath As String = "", Optional HWndOwner As Long = 0, Optional DialogTitle As String = "", Optional Flags As EBrowseForFolderFlags = eBrowseForFolderFlag_ReturnOnlyFSDirs)
    On Error GoTo hComponentFailure

    m_sInitialPath = InitialPath
    m_sDialogTitle = DialogTitle

    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrBrowseForFolder_InvalidWindowHandle, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = HWndOwner
    m_lFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Function    : Browse
'-- Notes       : Displays a dialog box that enables the user to select a shell folder.
'--               Returns True is a folder is selected and the OK button was pressed, False
'--               otherwise.
'------------------------------------------------------------------------------
Public Function Browse() As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- String
    '------------------------------------------------------------------------------
    Dim sPath  As String * MAX_LENGTH
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim lpInitialPath As Long
    Dim pidl   As Long
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    Browse = False

    With m_biInfo
        .hOwner = m_lHWndOwner
        .pIdlRoot = m_lPIdlRoot
        .pszDisplayName = m_sDisplayName
        .lpszTitle = m_sDialogTitle
        .ulFlags = m_lFlags

        .lpfn = FarProc(AddressOf BrowseForFolderCallback)

        lpInitialPath = 0
        If Len(m_sInitialPath) > 0 Then
            m_sInitialPath = m_sInitialPath & vbNullChar
            lpInitialPath = LocalAlloc(LPTR, Len(m_sInitialPath))

            If lpInitialPath <> 0 Then CopyMemory ByVal lpInitialPath, ByVal m_sInitialPath, Len(m_sInitialPath)
        End If

        .lParam = lpInitialPath

        .iImage = m_lImage
    End With

    pidl = SHBrowseForFolder(m_biInfo)

    If pidl Then
        '-- The user selected a folder and pressed OK; we must get the path of the selected folder
        sPath = String$(MAX_LENGTH, 0)

        If SHGetPathFromIDList(pidl, sPath) Then
            m_sSelectedPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
            Browse = True
        Else
            On Error GoTo 0
            Err.Raise eErrBrowseForFolder_CantGetPath, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_CantGetPath
        End If

        CoTaskMemFree pidl
        pidl = 0
    Else
        '-- The user canceled
    End If

    If lpInitialPath Then LocalFree lpInitialPath

    Exit Function

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Property    : Flags
'-- Notes       : Returns or sets the flags specifying the options for the
'--               Browse for Folder dialog box.
'------------------------------------------------------------------------------
Public Property Let Flags(ByVal Value As EBrowseForFolderFlags)
    On Error GoTo hComponentFailure

    m_lFlags = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property

Public Property Get Flags() As EBrowseForFolderFlags
    On Error GoTo hComponentFailure

    Flags = m_lFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DialogTitle
'-- Notes       : Returns or sets the dialog's prompt string.
'------------------------------------------------------------------------------
Public Property Let DialogTitle(ByVal Value As String)
    On Error GoTo hComponentFailure

    m_sDialogTitle = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property

Public Property Get DialogTitle() As String
    On Error GoTo hComponentFailure

    DialogTitle = m_sDialogTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : InitialPath
'-- Notes       : Returns or sets the folder used as a pre-selected folder in the
'--               Browse for Folder dialog.
'------------------------------------------------------------------------------
Public Property Let InitialPath(ByVal Value As String)
    On Error GoTo hComponentFailure

    m_sInitialPath = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property

Public Property Get InitialPath() As String
    On Error GoTo hComponentFailure

    InitialPath = m_sInitialPath

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : SelectedPath
'-- Notes       : Returns the folder selected in the Browse for Folder dialog.
'--               If the dialog was canceled returns the empty string ("").
'------------------------------------------------------------------------------
Public Property Get SelectedPath() As String
    On Error GoTo hComponentFailure

    SelectedPath = m_sSelectedPath

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : HWndOwner
'-- Notes       : Returns or sets a handle to the owner window for the Browse for Folder
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure

    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrBrowseForFolder_InvalidWindowHandle, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property

Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure

    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Sub         : Class_Initialize
'-- Notes       : Class data space initialization.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    m_lHWndOwner = 0
    m_lPIdlRoot = 0
    m_sDisplayName = String$(MAX_LENGTH, 0)
    m_sDialogTitle = ""
    m_lPFN = 0
    m_lFlags = eBrowseForFolderFlag_ReturnOnlyFSDirs + eBrowseForFolderFlag_StatusText + eBrowseForFolderFlag_UseNewUI
    m_lParam = 0
    m_lImage = 0
    m_sInitialPath = ""

    Exit Sub

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Function    : FarProc
'-- Notes       : Returns the address of the BrowseForFolder dialog callback proc.
'--               This function is necessary because AddressOff can be used only as a
'--               parameter part, i.e. it will be used in something like
'--               m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure

    FarProc = pfn

    Exit Function

hComponentFailure:
    Err.Raise eErrBrowseForFolder_ComponentFailure, App.EXEName & ".IAPP_BrowseForFolderHandler", S_ERR_ComponentFailure
End Function
