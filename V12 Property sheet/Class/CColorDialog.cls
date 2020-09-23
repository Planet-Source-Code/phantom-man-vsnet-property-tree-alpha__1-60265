VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_ClrDialogHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'-- Class Name.....: IAPP_ColorDialogHandler
'-- Description....: Class for displaying the standard Colour dialog box.
'--
'-- Notes..........: Use this class to provide in your application the possibility of
'--                  displaying the standard Colour dialog box without using the file
'--                  CommDlg32.ocx.
'--
'--
'-- Author, date...: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'-- HwndOwner            Long          Returns or sets a handle to the owner window for the standard Colour dialog box.
'-- Flags                EColourFlags  Returns or sets the flags specifying the options for the standard Colour dialog box.
'-- InitialColour        Long          Returns or sets the standard Colour dialog box initial Colour.
'-- SelectedColour       Long          Returns the selected Colour if the OK button was pressed, -1 if the
'--                                    standard Colour dialog was canceled.
'-- DialogTitle          String        Returns or sets the standard Colour dialog title.
'-- CenterDialog         Boolean       Returns or sets if the standard Colour dialog box should appear centered
'--                                    on the screen or in its default position.
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- Init                 Use this routine for basic object initialization.
'-- Show                 Displays the standard Colour selection dialog box.
'--                      Returns true if OK was selected, false otherwise.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'-- FarProc              Returns the address of the Colour dialog callback proc.
'--                      This function is necessary because AddressOff can be used only as a
'--                      parameter part, i.e. it will be used in something like
'--                      m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------

Option Explicit

'-- Error handling definitions
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const E_ERR_BASE = 17120 + vbObjectError
Public Enum EErrColourDialog
    eErrColourDialog_InvalidWindowHandle = E_ERR_BASE + 1
    eErrColourDialog_ComponentFailure
End Enum
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "IAPP_ColorDialogHandler component failure"

'-- Public class enums
Public Enum EColourFlags
    eColourFlag_FullOpen = 2
    eColourFlag_PreventFullOpen = 4
    eColourFlag_HelpButton = 8
End Enum

'-- Private class constants
Private Const eColourFlag_RGBInit = 1
Private Const eColourFlag_EnableHook = &H10

'-- Private class type definitions
Private Type ChooseColor
    lStructSize As Long
    HWndOwner  As Long
    hInstance  As Long
    rgbResult  As Long
    lpCustColours As Long
    Flags      As Long
    lCustData  As Long
    lpfnHook   As Long
    lpTemplateName As String
End Type

'-- Private class API function declarations
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChooseColor As ChooseColor) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Private variables to hold property values
Private m_eFlags As EColourFlags
Private m_lHWndOwner As Long
Private m_lInitialColour As Long
Private m_lSelectedColour As Long

'-- Private variables for internal class use
Private m_bFirstTime As Boolean
Private m_lArrayCustomColours(15) As Long             'array to store the 16 custom Colours


'------------------------------------------------------------------------------
'-- Sub         : Init
'-- Notes       : Use this routine for basic object initialization.
'------------------------------------------------------------------------------
Public Sub Init(Optional InitialColour As Long = -1, Optional DialogTitle As String = "", Optional HWndOwner As Long = 0, Optional Flags As EColourFlags = eColourFlag_PreventFullOpen)
    On Error GoTo hComponentFailure

    m_lInitialColour = InitialColour
    g_sColourDialogTitle = DialogTitle

    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrColourDialog_InvalidWindowHandle, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = HWndOwner

    m_eFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Property    : HWndOwner
'-- Notes       : Returns or sets a handle to the owner window for the standard Colour
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure

    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrColourDialog_InvalidWindowHandle, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure

    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Flags
'-- Notes       : Returns or sets the flags specifying the options for the standard Colour
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let Flags(ByVal Value As EColourFlags)
    On Error GoTo hComponentFailure

    m_eFlags = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Flags() As EColourFlags
    On Error GoTo hComponentFailure

    Flags = m_eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : InitialColour
'-- Notes       : Returns or sets the standard Colour dialog box initial Colour.
'------------------------------------------------------------------------------
Public Property Let InitialColour(ByVal Value As Long)
    On Error GoTo hComponentFailure

    m_lInitialColour = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get InitialColour() As Long
    On Error GoTo hComponentFailure

    InitialColour = m_lInitialColour

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : SelectedColour
'-- Notes       : Returns the selected Colour if the OK button was pressed, -1 if the
'--               standard Colour dialog was canceled.
'------------------------------------------------------------------------------
Public Property Get SelectedColour() As Long
    On Error GoTo hComponentFailure

    SelectedColour = m_lSelectedColour

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DialogTitle
'-- Notes       : Returns or sets the standard Colour dialog title.
'------------------------------------------------------------------------------
Public Property Get DialogTitle() As String
    On Error GoTo hComponentFailure

    DialogTitle = g_sColourDialogTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Let DialogTitle(ByVal Value As String)
    On Error GoTo hComponentFailure

    g_sColourDialogTitle = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Function    : CenterDialog
'-- Notes       : Returns or sets if the standard Colour dialog box should appear centered
'--               on the screen or in its default position.
'------------------------------------------------------------------------------
Public Property Let CenterDialog(Value As Boolean)
    On Error GoTo hComponentFailure

    g_bCenterColourDialog = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get CenterDialog() As Boolean
    On Error GoTo hComponentFailure

    CenterDialog = g_bCenterColourDialog

    Exit Property

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Function    : Show
'-- Notes       : Displays the standard Colour selection dialog box.
'--               Returns true if OK was selected, false otherwise.
'------------------------------------------------------------------------------
Public Function Show() As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim lReturn As Long
    '------------------------------------------------------------------------------
    '-- DataType
    '------------------------------------------------------------------------------
    Dim CC     As ChooseColor
    '------------------------------------------------------------------------------
    '-- Integer
    '------------------------------------------------------------------------------
    Dim iColourMember As Integer
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    CC.lStructSize = Len(CC)
    CC.rgbResult = m_lInitialColour

    If Not m_bFirstTime Then

        For iColourMember = 0 To 15
            m_lArrayCustomColours(iColourMember) = &HFFFFFF
        Next

        m_bFirstTime = True

    End If

    CC.lpCustColours = VarPtr(m_lArrayCustomColours(0))

    If m_lInitialColour <> -1 Then m_eFlags = m_eFlags + eColourFlag_RGBInit
    If g_bCenterColourDialog Or (Len(g_sColourDialogTitle) > 0) Then m_eFlags = m_eFlags + eColourFlag_EnableHook

    CC.Flags = m_eFlags
    CC.HWndOwner = Me.HWndOwner

    ReleaseCapture

    If m_eFlags And eColourFlag_EnableHook = eColourFlag_EnableHook Then
        CC.lpfnHook = FarProc(AddressOf ColourDialogCallback)
    End If

    lReturn = ChooseColor(CC)

    If lReturn = 0 Then
        '-- The user canceled
        m_lSelectedColour = -1
        Show = False
    Else
        '-- A valid choice was made
        m_lSelectedColour = CC.rgbResult
        Show = True
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : FarProc
'-- Notes       : Returns the address of the Colour dialog callback proc.
'--               This function is necessary because AddressOff can be used only as a
'--               parameter part, i.e. it will be used in something like
'--               m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure

    FarProc = pfn

    Exit Function

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Sub         : Class_Initialize
'-- Notes       : Class data initialization.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    m_lSelectedColour = -1
    g_bCenterColourDialog = True
    g_sColourDialogTitle = "Colour"
    Exit Sub

hComponentFailure:
    Err.Raise eErrColourDialog_ComponentFailure, App.EXEName & ".IAPP_ColorDialogHandler", S_ERR_ComponentFailure
End Sub
