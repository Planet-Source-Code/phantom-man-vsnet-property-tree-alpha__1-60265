VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_FontDialogHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'-- Class Name.....: IAPP_FontDialogHandlerHandler
'-- Description....: Class for displaying the Font selection Dialog.
'--
'-- Notes..........: Use this class to display the Font common dialog without the
'--                  CommDlg32.ocx file loaded.
'--
'--
'-- Author, date...: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'-- CenterDialog         Boolean       Returns or sets if the standard Font dialog box should appear centered
'--                                    on the screen or in its default position.
'-- Colour               Long          Returns or sets the selected font Colour
'-- DialogTitle          String        Returns or sets the standard Font dialog box title.
'-- Flags                EFontFlags    Returns or sets the flags specifying the options for the standard Font dialog box.
'-- Font                 STDFont       Returns or set a Font object representing the selections made in the
'--                                    standard Font dialog box.
'-- FontType             EFontType     Returns or sets the flags specifying the font type displayed by the
'--                                    standard Font dialog box.
'-- HwndOwner            Long          Returns or sets a handle to the owner window for the standard Font dialog box.
'-- Max                  Integer       Returns or sets the maximum font size.
'-- Min                  Integer       Returns or sets the minimum font size.
'-- Style                Integer       Returns or sets the style of the font.
'--
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- Init                 Use this routine for basic object initialization.
'-- Show                 Displays the standard Font dialog box.
'--                      Returns True if OK was selected, False otherwise.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'-- FarProc              Returns the address of the Font dialog callback proc.
'--                      This function is necessary because AddressOff can be used only as a
'--                      parameter part, i.e. it will be used in something like
'--                      m_Address = FarProc(AddressOff ProcName)
'-- IsArrayEmpty         Checks if the specified array is empty.
'-- StrToArray           Converts a string to a bytes array.
'------------------------------------------------------------------------------

Option Explicit

'-- Error handling definitions
Private Const E_ERR_BASE = 17520 + vbObjectError
Public Enum EErrFontDialog
    eErrFontDialog_InvalidWindowHandle = E_ERR_BASE + 1
    eErrFontDialog_ComponentFailure
End Enum
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "IAPP_FontDialogHandler component failure"

'-- Private class constants
Private Const LF_FACESIZE = 32

'-- Private class type definitions
Private Type LOGFONT
    lfHeight   As Long
    lfWidth    As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight   As Long
    lfItalic   As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet  As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality  As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Private Type ChooseFont
    lStructSize As Long                               '-- Filled with UDT size
    HWndOwner  As Long                                '-- Caller's window handle
    HDC        As Long                                '-- Printer DC/IC or NULL
    lpLogFont  As Long                                '-- Pointer to LOGFONT
    iPointSize As Long                                '-- 10 * size in points of font
    Flags      As Long                                '-- Type flags
    rgbColours As Long                                '-- Returned text Colour
    lCustData  As Long                                '-- Data passed to hook function
    lpfnHook   As Long                                '-- Pointer to hook function
    lpTemplateName As String                          '-- Custom template name
    hInstance  As Long                                '-- Instance handle for template
    lpszStyle  As String                              '-- Return style field
    nFontType  As Integer                             '-- Font type bits
    iAlign     As Integer                             '-- Filler
    nSizeMin   As Long                                '-- Minimum point size allowed
    nSizeMax   As Long                                '-- Maximum point size allowed
End Type

'-- Private class API function declarations
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (chfont As ChooseFont) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Public class enums
Public Enum EFontFlags
    eFontFlag_ScreenFonts = &H1
    eFontFlag_PrinterFonts = &H2
    eFontFlag_Bots = &H3
    eFontFlag_FontShowHelp = &H4
    eFontFlag_EnableHook = &H8
    eFontFlag_EnableTemplate = &H10
    eFontFlag_EnableTemplateHandle = &H20
    eFontFlag_InitToLogFontStruct = &H40
    eFontFlag_UseStyle = &H80
    eFontFlag_Effects = &H100
    eFontFlag_Apply = &H200
    eFontFlag_ANSIOnly = &H400
    eFontFlag_NoVectorFonts = &H800
    eFontFlag_NoOemFonts = eFontFlag_NoVectorFonts
    eFontFlag_NoSimulations = &H1000
    eFontFlag_LimitSize = &H2000
    eFontFlag_FixedPitchOnly = &H4000
    eFontFlag_WYSIWYG = &H8000                        '-- Must also have ScreenFonts And PrinterFonts
    eFontFlag_ForceFontExist = &H10000
    eFontFlag_ScalableOnly = &H20000
    eFontFlag_TTOnly = &H40000
    eFontFlag_NoFaceSel = &H80000
    eFontFlag_NoStyleSel = &H100000
    eFontFlag_NoSizeSel = &H200000
    eFontFlag_SelectScript = &H400000
    eFontFlag_NoScriptSel = &H800000
    eFontFlag_NoVertFonts = &H1000000
End Enum

Public Enum EFontType
    eFontType_Bold = &H100
    eFontType_Italic = &H200
    eFontType_Regular = &H400
    eFontType_Screen = &H2000
    eFontType_Printer = &H4000
    eFontType_Simulated = &H8000
End Enum

'-- Private variables to hold property values
Private m_eFlags As EFontFlags
Private m_eFontType As EFontType
Private m_iMax As Integer
Private m_iMin As Integer
Private m_iStyle As Integer
Private m_lColour As Long
Private m_lHWndOwner As Long
Private m_oFont As StdFont


'------------------------------------------------------------------------------
'-- Sub         : Init
'-- Notes       : Use this routine for basic object initialization.
'------------------------------------------------------------------------------
Public Sub Init(Optional Font As StdFont, Optional DialogTitle As String = "", Optional HWndOwner As Long = 0, Optional Flags As EFontFlags)
    On Error GoTo hComponentFailure

    Set m_oFont = Font
    g_sFontDialogTitle = DialogTitle

    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrFontDialog_InvalidWindowHandle, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = HWndOwner

    g_bCenterFontDialog = CenterDialog
    m_eFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Function    : CenterDialog
'-- Notes       : Returns or sets if the standard Font dialog box should appear centered
'--               on the screen or in its default position.
'------------------------------------------------------------------------------
Public Property Let CenterDialog(Value As Boolean)
    On Error GoTo hComponentFailure

    g_bCenterFontDialog = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get CenterDialog() As Boolean
    On Error GoTo hComponentFailure

    CenterDialog = g_bCenterFontDialog

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Colour
'-- Notes       : Returns or sets the selected font Colour.
'------------------------------------------------------------------------------
Public Property Let Colour(ByVal Value As Long)
    On Error GoTo hComponentFailure

    m_lColour = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Colour() As Long
    On Error GoTo hComponentFailure

    Colour = m_lColour

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : DialogTitle
'-- Notes       : Returns or sets the standard Font dialog box title.
'------------------------------------------------------------------------------
Public Property Get DialogTitle() As String
    On Error GoTo hComponentFailure

    DialogTitle = g_sFontDialogTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Let DialogTitle(ByVal strTitle As String)
    On Error GoTo hComponentFailure

    g_sFontDialogTitle = strTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Flags
'-- Notes       : Returns or sets the flags specifying the options for the standard Font
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let Flags(ByVal vData As EFontFlags)
    On Error GoTo hComponentFailure

    m_eFlags = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Flags() As EFontFlags
    On Error GoTo hComponentFailure

    Flags = m_eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Font
'-- Notes       : Returns or set a Font object representing the selections made in the
'--               standard Font dialog box.
'------------------------------------------------------------------------------
Public Property Let Font(ByVal Value As StdFont)
    On Error GoTo hComponentFailure

    Set m_oFont = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Font() As StdFont
    On Error GoTo hComponentFailure

    Set Font = m_oFont

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : FontType
'-- Notes       : Returns or sets the flags specifying the font type displayed by the
'--               standard Font dialog box.
'------------------------------------------------------------------------------
Public Property Let FontType(ByVal vData As EFontType)
    On Error GoTo hComponentFailure

    m_eFontType = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get FontType() As EFontType
    On Error GoTo hComponentFailure

    FontType = m_eFontType

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : HWndOwner
'-- Notes       : Returns or sets a handle to the owner window for the standard Font
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure

    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrFontDialog_InvalidWindowHandle, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure

    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Max
'-- Notes       : Returns or sets the maximum font size.
'------------------------------------------------------------------------------
Public Property Let Max(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iMax = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Max() As Integer
    On Error GoTo hComponentFailure

    Max = m_iMax

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Min
'-- Notes       : Returns or sets the minimum font size.
'------------------------------------------------------------------------------
Public Property Let Min(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iMin = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Min() As Integer
    On Error GoTo hComponentFailure

    Min = m_iMin

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Style
'-- Notes       : Returns or sets the style of the font.
'------------------------------------------------------------------------------
Public Property Let Style(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iStyle = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Style() As Integer
    On Error GoTo hComponentFailure

    Style = m_iStyle

    Exit Property

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Function    : Show
'-- Notes       : Displays the standard Font dialog box.
'--               Returns True if OK was selected, False otherwise.
'------------------------------------------------------------------------------
Public Function Show() As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- Const
    '------------------------------------------------------------------------------
    Const eFontFlag_FontNotSupported = eFontFlag_Apply Or eFontFlag_EnableHook Or eFontFlag_EnableTemplate
    Const PointsPerTwip = 1440 / 72
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim m_hdc  As Long
    '------------------------------------------------------------------------------
    '-- DataType
    '------------------------------------------------------------------------------
    Dim fnt    As LOGFONT
    Dim cf     As ChooseFont
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    If (m_oFont Is Nothing) Then Set m_oFont = New StdFont

    If m_eFlags And eFontFlag_PrinterFonts Then m_hdc = Printer.HDC
    If (m_eFlags And eFontFlag_PrinterFonts) = 0 Then m_eFlags = m_eFlags Or eFontFlag_ScreenFonts

    If m_lColour <> vbBlack Then m_eFlags = m_eFlags Or eFontFlag_Effects

    If m_iMin Then m_eFlags = m_eFlags Or eFontFlag_LimitSize
    If m_iMax Then m_eFlags = m_eFlags Or eFontFlag_LimitSize

    '-- Put in required internal flags and remove unsupported
    Flags = (Flags Or eFontFlag_InitToLogFontStruct) And Not eFontFlag_FontNotSupported


    If g_bCenterFontDialog Or (Len(g_sFontDialogTitle) > 0) Then m_eFlags = m_eFlags + eFontFlag_EnableHook

    If m_eFlags And eFontFlag_EnableHook = eFontFlag_EnableHook Then
        cf.lpfnHook = FarProc(AddressOf FontDialogCallback)
    End If

    fnt.lfHeight = -(Font.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = Font.Weight
    fnt.lfItalic = Font.Italic
    fnt.lfUnderline = Font.Underline
    fnt.lfStrikeOut = Font.Strikethrough

    StrToBytes fnt.lfFaceName, Font.Name

    cf.lStructSize = Len(cf)
    cf.HWndOwner = m_lHWndOwner
    cf.HDC = m_hdc
    cf.lpLogFont = VarPtr(fnt)
    cf.iPointSize = Font.Size * 10
    cf.Flags = m_eFlags
    cf.rgbColours = m_lColour
    cf.nSizeMin = m_iMin
    cf.nSizeMax = m_iMax

    If ChooseFont(cf) Then
        Show = True

        m_eFlags = cf.Flags
        m_lColour = cf.rgbColours

        Font.Bold = cf.nFontType And eFontType_Bold
        Font.Italic = fnt.lfItalic
        Font.Strikethrough = fnt.lfStrikeOut
        Font.Underline = fnt.lfUnderline
        Font.Weight = fnt.lfWeight
        Font.Size = cf.iPointSize / 10
        Font.Name = StrConv(fnt.lfFaceName, vbUnicode)
    Else
        Show = False
    End If

    Exit Function

hComponentFailure:
    Show = False
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : FarProc
'-- Notes       : Returns the address of the Font dialog callback proc.
'--               This function is necessary because AddressOff can be used only as a
'--               parameter part, i.e. it will be used in something like
'--               m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure

    FarProc = pfn

    Exit Function

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : IsArrayEmpty
'-- Notes       : Checks if the specified array is empty.
'------------------------------------------------------------------------------
Private Function IsArrayEmpty(va As Variant) As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- Variant
    '------------------------------------------------------------------------------
    Dim v      As Variant
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)

    Exit Function

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Sub         : StrToBytes
'-- Notes       : Converts a string to a bytes array.
'------------------------------------------------------------------------------
Private Sub StrToBytes(ByRef ab() As Byte, ByVal s As String)
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim cab    As Long
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------

    If IsArrayEmpty(ab) Then
        ab = StrConv(s, vbFromUnicode)
    Else
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), vbNullChar)

        CopyMemoryStr ab(LBound(ab)), s, cab
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Sub         : Class_Initialize
'-- Notes       : Class data space initialization.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    g_bCenterFontDialog = True

    Exit Sub

hComponentFailure:
    Err.Raise eErrFontDialog_ComponentFailure, App.EXEName & ".IAPP_FontDialogHandler", S_ERR_ComponentFailure
End Sub
