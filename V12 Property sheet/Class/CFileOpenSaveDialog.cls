VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_FOSHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'-- Class Name.....: IAPP_FileOpenSaveHandler
'-- Description....: Class for displaying the File Open/Save Common Dialog.
'--
'-- Notes..........: Use this class to display the File Open/Save common dialog without the
'--                  CommDlg32.ocx file loaded and retrieve the selected file name and path.
'--                  It also offers great support for multiple file selections and more.
'--
'--
'-- Author, date...: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'-- CenterDialog         Boolean       Returns or sets if the standard File Open/Save dialog box dialog box
'--                                    should appear centered on the screen or in its default position.
'-- DefaultEXT           String        Returns or sets the default filename extension for the standard File
'--                                    Open/Save dialog box.
'-- DefaultFilter        Integer       Returns or sets a default filter for the standard File Open/Save dialog box.
'-- DialogTitle          String        Returns or sets the standard File Open/Save dialog box title.
'-- FileCount            Integer       Returns an integer value containing the number of files selected in the
'--                                    standard File Open/Save dialog box.
'-- FileName             String        Returns or sets the path and the name of the selected file.
'-- FileTitle            String        Returns or sets the name (without the path) of the file to open or save at run time.
'-- Filter               String        Returns or sets the filters that are displayed in the Type list box of
'--                                    the standard File Open/Save dialog box.
'-- Flags                EFileOpenSaveFlags Returns or sets the flags specifying the options for the standard File
'--                                         Open/Save dialog box.
'-- HwndOwner            Long          Returns or sets a handle to the owner window for the standard File
'--                                    Open/Save dialog box.
'-- InitialDir           String        Returns or sets the initial file directory.
'-- MaxFileSize          Integer       Returns or sets the maximum size of the name of the file selected using
'--                                    the standard File Open/Save dialog box.
'-- SelectedPath         String        Returns a string value containing the path for the selected files,
'--                                    without the filename.
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- Init                 Use this routine for basic object initialization.
'-- GetNextFilename      Returns a string value containing the path and filename of the selected
'--                      file specified by Index.
'-- Show                 Displays the standard File Open/Save dialog box. Returns True if OK was
'--                      selected, False otherwise.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'-- PrepareString        Private function manipulating the string resulted after the file
'--                      selection is made.
'-- FarProc              Returns the address of the File Open/Save dialog callback proc.
'--                      This function is necessary because AddressOff can be used only as a
'--                      parameter part, i.e. it will be used in something like
'--                      m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------

Option Explicit

'-- Error handling definitions
Private Const E_ERR_BASE = 17460 + vbObjectError
Public Enum EErrFileOpenSaveDialog
    eErrFileOpenSaveDialog_InvalidWindowHandle = E_ERR_BASE + 1
    eErrFileOpenSaveDialog_ComponentFailure
End Enum
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "IAPP_FileOpenSaveHandler component failure"

'-- Public class enums
Public Enum EDialogType
    eDialogType_OpenFile
    eDialogType_SaveFile
End Enum
Public Enum EFileOpenSaveFlags
    eFileOpenSaveFlag_ReadOnly = &H1                  '-- Checks Read Only check box for Open and Save As dialog boxes.Checks Read Only check box for Open and Save As dialog boxes.
    eFileOpenSaveFlag_OverwritePrompt = &H2           '-- Generates a message box if the selected file already exists.
    eFileOpenSaveFlag_HideReadOnly = &H4              '-- Hides the Read Only check box.
    eFileOpenSaveFlag_NoChangeDir = &H8               '-- Sets the current directory to what it was when the dialog box was invoked.
    eFileOpenSaveFlag_HelpButton = &H10               '-- Causes the dialog box to display the Help button.
    eFileOpenSaveFlag_EnableHook = &H20
    eFileOpenSaveFlag_EnableTemplate = &H40
    eFileOpenSaveFlag_EnableTemplateHandle = &H80
    eFileOpenSaveFlag_NoValidate = &H100              '-- Allows invalid characters in the returned file name.
    eFileOpenSaveFlag_AllowMultiselect = &H200        '-- Allows the File Name list box to have multiple selections.
    eFileOpenSaveFlag_ExtensionDifferent = &H400      '-- Extension of returned file name is different from the one set by DefaultExt.
    eFileOpenSaveFlag_PathMustExist = &H800           '-- User can enter only valid path names.
    eFileOpenSaveFlag_FileMustExist = &H1000          '-- User can enter only names of existing files.
    eFileOpenSaveFlag_CreatePrompt = &H2000           '-- Asks if the user wants to create a file that does not currently exist.
    eFileOpenSaveFlag_ShareAware = &H4000             '-- Sharing violation errors will be ignored.
    eFileOpenSaveFlag_NoReadOnlyReturn = &H8000       '-- The returned file will not have the Read Only attribute set.
    eFileOpenSaveFlag_NoTestFileCreate = &H10000
    eFileOpenSaveFlag_NoNetworkButton = &H20000
    eFileOpenSaveFlag_NoLongNames = &H40000           '-- No long filenames.
    eFileOpenSaveFlag_Explorer = &H80000              '-- Windows 95 Open A File dialog box template.
    eFileOpenSaveFlag_NoDereferenceLinks = &H100000   '-- No shortcuts.
    eFileOpenSaveFlag_LongNames = &H200000            '-- Long filenames.
End Enum

'-- Private class type definitions
Private Type OPENFILENAME
    lStructSize As Long
    HWndOwner  As Long
    hInstance  As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile  As String
    nMaxFile   As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags      As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData  As Long
    lpfnHook   As Long
    lpTemplateName As String
End Type

'-- Private class API function declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Private variables to hold property values
Private m_bCentered As Boolean
Private m_eFlags As EFileOpenSaveFlags
Private m_iFileCount As Integer
Private m_iDefaultFilter As Integer
Private m_iMaxFileSize As Integer
Private m_lHWndOwner As Long
Private m_sDefaultExt As String
Private m_sDialogTitle As String
Private m_sFileName As String
Private m_sFileTitle As String
Private m_sFilter As String
Private m_sInitialDir As String
Private m_sSelectedPath As String

'-- Private variables for internal class use
Private m_sArray() As String

'-- Private class constants
Private Const MAX_LENGTH As Integer = 1024


'------------------------------------------------------------------------------
'-- Sub         : Init
'-- Notes       : Use this routine for basic object initialization.
'------------------------------------------------------------------------------
Public Sub Init(Filter As String, Optional DefaultFilter As Integer = 1, Optional InitialDir As String = "", Optional FileName As String = "", Optional DialogTitle As String = "", Optional HWndOwner As Long = 0, Optional Flags As EFileOpenSaveFlags = eFileOpenSaveFlag_Explorer + eFileOpenSaveFlag_LongNames)
    On Error GoTo hComponentFailure

    m_sFilter = Filter
    m_iDefaultFilter = DefaultFilter
    m_sFileName = FileName
    m_sDialogTitle = DialogTitle
    m_sInitialDir = InitialDir

    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrFileOpenSaveDialog_InvalidWindowHandle, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = HWndOwner

    m_eFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Sub

'------------------------------------------------------------------------------
'-- Property    : CenterDialog
'-- Notes       : Returns or sets if the standard File Open/Save dialog box dialog box
'--               should appear centered on the screen or in its default position.
'------------------------------------------------------------------------------
Public Property Let CenterDialog(Value As Boolean)
    On Error GoTo hComponentFailure

    m_bCentered = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Get CenterDialog() As Boolean
    On Error GoTo hComponentFailure

    CenterDialog = m_bCentered

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DefaultExt
'-- Notes       : Returns or sets the default filename extension for the standard File
'--               Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get DefaultExt() As String
    On Error GoTo hComponentFailure

    DefaultExt = m_sDefaultExt

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let DefaultExt(ByVal strExt As String)
    On Error GoTo hComponentFailure

    m_sDefaultExt = strExt

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DefaultFilter
'-- Notes       : Returns or sets a default filter for the standard File Open/Save dialog
'--               box.
'------------------------------------------------------------------------------
Public Property Get DefaultFilter() As Integer
    On Error GoTo hComponentFailure

    DefaultFilter = m_iDefaultFilter

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let DefaultFilter(ByVal Value As Integer)
    On Error GoTo hComponentFailure

    m_iDefaultFilter = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DialogTitle
'-- Notes       : Returns or sets the standard File Open/Save dialog box title.
'------------------------------------------------------------------------------
Public Property Get DialogTitle() As String
    On Error GoTo hComponentFailure

    DialogTitle = m_sDialogTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let DialogTitle(ByVal strTitle As String)
    On Error GoTo hComponentFailure

    m_sDialogTitle = strTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : FileCount
'-- Notes       : Returns an integer value containing the number of files selected in the
'--               standard File Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get FileCount() As Integer
    On Error GoTo hComponentFailure

    FileCount = m_iFileCount

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : FileName
'-- Notes       : Returns or sets the path and the name of the selected file.
'------------------------------------------------------------------------------
Public Property Get FileName() As String
    On Error GoTo hComponentFailure

    FileName = m_sFileName

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let FileName(ByVal strFileName As String)
    On Error GoTo hComponentFailure

    m_sFileName = strFileName

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : FileTitle
'-- Notes       : Returns or sets the name (without the path) of the file to open or save
'--               at run time.
'------------------------------------------------------------------------------
Public Property Get FileTitle() As String
    On Error GoTo hComponentFailure

    FileTitle = m_sFileTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let FileTitle(ByVal strFileTitle As String)
    On Error GoTo hComponentFailure

    m_sFileTitle = strFileTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Filter
'-- Notes       : Returns or sets the filters that are displayed in the Type list box of
'--               the standard File Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get Filter() As String
    On Error GoTo hComponentFailure

    Filter = m_sFilter

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let Filter(ByVal strFilter As String)
    On Error GoTo hComponentFailure

    m_sFilter = strFilter

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Flags
'-- Notes       : Returns or sets the flags specifying the options for the standard File
'--               Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get Flags() As EFileOpenSaveFlags
    On Error GoTo hComponentFailure

    Flags = m_eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let Flags(ByVal eFlags As EFileOpenSaveFlags)
    On Error GoTo hComponentFailure

    m_eFlags = eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : HWndOwner
'-- Notes       : Returns or sets a handle to the owner window for the standard File
'--               Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure

    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure

    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrFileOpenSaveDialog_InvalidWindowHandle, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : InitialDir
'-- Notes       : Returns or sets the initial file directory.
'------------------------------------------------------------------------------
Public Property Get InitialDir() As String
    On Error GoTo hComponentFailure

    InitialDir = m_sInitialDir

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let InitialDir(ByVal strInitialDir As String)
    On Error GoTo hComponentFailure

    m_sInitialDir = strInitialDir

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : MaxFileSize
'-- Notes       : Returns or sets the maximum size of the name of the file selected using
'--               the standard File Open/Save dialog box.
'------------------------------------------------------------------------------
Public Property Get MaxFileSize() As Integer
    On Error GoTo hComponentFailure

    MaxFileSize = m_iMaxFileSize

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property

Public Property Let MaxFileSize(ByVal intMaxFileSize As Integer)
    On Error GoTo hComponentFailure

    m_iMaxFileSize = intMaxFileSize

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : SelectedPath
'-- Notes       : Returns a string value containing the path for the selected files,
'--               without the filename.
'------------------------------------------------------------------------------
Public Property Get SelectedPath() As String
    On Error GoTo hComponentFailure

    SelectedPath = m_sSelectedPath

    Exit Property

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Function    : GetNextFileName
'-- Notes       : Returns a string value containing the path and filename of the selected
'--               file specified by Index.
'------------------------------------------------------------------------------
Public Function GetNextFileName(Optional Index As Integer = 1) As String
    On Error GoTo hComponentFailure

    If (Index < 1) Or (m_iFileCount = 1) Or (Index > m_iFileCount) Then
        GetNextFileName = m_sFileName
    Else
        GetNextFileName = m_sSelectedPath & "\" & m_sArray(Index)
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : Show
'-- Notes       : Displays the standard File Open/Save dialog box. Returns True if OK was
'--               selected, False otherwise.
'------------------------------------------------------------------------------
Public Function Show(DialogType As EDialogType) As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- String
    '------------------------------------------------------------------------------
    Dim sChar  As String * 1
    Dim sTemp  As String
    '------------------------------------------------------------------------------
    '-- DataType
    '------------------------------------------------------------------------------
    Dim OF     As OPENFILENAME
    '------------------------------------------------------------------------------
    '-- Integer
    '------------------------------------------------------------------------------
    Dim iCounter As Integer
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    OF.lpstrTitle = m_sDialogTitle & ""
    OF.lpstrDefExt = m_sDefaultExt & ""
    OF.lStructSize = LenB(OF)
    OF.lpstrFilter = m_sFilter & "||"
    OF.nFilterIndex = m_iDefaultFilter

    If m_bCentered Then m_eFlags = m_eFlags Or eFileOpenSaveFlag_EnableHook

    OF.Flags = m_eFlags

    If (m_eFlags And eFileOpenSaveFlag_EnableHook) Then

        If (m_eFlags And eFileOpenSaveFlag_Explorer) Then
            OF.lpfnHook = FarProc(AddressOf FileOpenSaveDialogCallbackEx)
        Else
            OF.lpfnHook = FarProc(AddressOf FileOpenSaveDialogCallback)
        End If

    End If

    For iCounter = 1 To Len(m_sFilter)

        sChar = Mid$(m_sFilter, iCounter, 1)

        If sChar = "|" Then
            sTemp = sTemp & vbNullChar
        Else
            sTemp = sTemp & sChar
        End If

    Next

    sTemp = sTemp & vbNullChar & vbNullChar
    OF.lpstrFilter = sTemp

    sTemp = m_sFileName & String$(MAX_LENGTH - Len(m_sFileName), 0)
    OF.lpstrFile = sTemp
    OF.nMaxFile = MAX_LENGTH

    sTemp = m_sFileTitle & String$(MAX_LENGTH - Len(m_sFileTitle), 0)
    OF.lpstrFileTitle = sTemp
    OF.lpstrInitialDir = m_sInitialDir
    OF.nMaxFileTitle = MAX_LENGTH
    OF.HWndOwner = m_lHWndOwner

    Select Case DialogType

        Case eDialogType_OpenFile

            If GetOpenFileName(OF) Then
                Show = True
                m_sFileName = PrepareString(OF.lpstrFile, 1)
                m_sFileTitle = PrepareString(OF.lpstrFileTitle, 2)
            Else
                Show = False
            End If

        Case eDialogType_SaveFile

            If GetSaveFileName(OF) Then
                Show = True
                m_sFileName = PrepareString(OF.lpstrFile, 1)
                m_sFileTitle = PrepareString(OF.lpstrFileTitle, 2)
            Else
                Show = False
            End If

    End Select

    Exit Function

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : PrepareString
'-- Notes       : Private function manipulating the string resulted after the file
'--               selection is made.
'------------------------------------------------------------------------------
Private Function PrepareString(ByVal sIn As String, Index As Integer) As String
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- String
    '------------------------------------------------------------------------------
    Dim sTemp  As String
    '------------------------------------------------------------------------------
    '-- Integer
    '------------------------------------------------------------------------------
    Dim iPos   As Integer
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim l      As Long
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    sTemp = sIn
    If Len(Trim$(sTemp)) > 0 Then

        iPos = InStr(sTemp, vbNullChar + vbNullChar)
        If iPos > 1 Then sTemp = Left$(sTemp, iPos - 1)

        Select Case Index

            Case 1

                l = m_eFlags And eFileOpenSaveFlag_AllowMultiselect

                If l = 0 Then
                    iPos = InStr(sTemp, vbNullChar)
                    If iPos > 0 Then sTemp = IIf(iPos = 1, "", Left$(sTemp, iPos - 1))
                    m_iFileCount = 1
                    iPos = InStrRev(sTemp, "\")
                    If iPos > 0 Then m_sSelectedPath = Left$(sTemp, iPos)
                    PrepareString = sTemp
                Else
                    ReDim m_sArray(0)
                    m_sArray = Split(sTemp, vbNullChar)

                    m_iFileCount = UBound(m_sArray) - LBound(m_sArray)

                    If m_iFileCount > 1 Then
                        sTemp = m_sArray(0)
                        If Right$(sTemp, 1) = "\" Then sTemp = Left$(sTemp, Len(sTemp) - 1)
                        m_sSelectedPath = sTemp
                        PrepareString = sTemp & "\" & m_sArray(1)
                    Else
                        iPos = InStrRev(sTemp, "\")
                        If iPos > 2 Then
                            m_sSelectedPath = Left$(sTemp, iPos)
                        Else
                            m_sSelectedPath = ""
                        End If
                        PrepareString = m_sArray(0)
                        m_iFileCount = 1
                    End If
                End If

            Case 2
                iPos = InStr(sTemp, vbNullChar)
                If iPos > 0 Then
                    PrepareString = IIf(iPos = 1, "", Left$(sTemp, iPos - 1))
                Else
                    PrepareString = sTemp
                End If

        End Select

    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : FarProc
'-- Notes       : Returns the address of the File Open/Save dialog callback proc.
'--               This function is necessary because AddressOff can be used only as a
'--               parameter part, i.e. it will be used in something like
'--               m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure

    FarProc = pfn

    Exit Function

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Sub         : Class_Initialize
'-- Notes       : Class data space initialization.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    m_bCentered = False
    m_iFileCount = 0
    m_iDefaultFilter = 1
    m_iMaxFileSize = MAX_LENGTH
    m_lHWndOwner = 0
    m_eFlags = eFileOpenSaveFlag_Explorer + eFileOpenSaveFlag_LongNames
    m_sDefaultExt = ""
    m_sDialogTitle = ""
    m_sFileName = ""
    m_sFileTitle = ""
    m_sFilter = ""
    m_sInitialDir = ""
    m_sSelectedPath = ""
    ReDim m_sArray(0)

    Exit Sub

hComponentFailure:
    Err.Raise eErrFileOpenSaveDialog_ComponentFailure, App.EXEName & ".IAPP_FileOpenSaveHandler", S_ERR_ComponentFailure
End Sub
