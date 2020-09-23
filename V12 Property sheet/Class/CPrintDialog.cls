VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_PrntDialogHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
'-- Class Name.....: IAPP_PrintDialogHandler
'-- Description....: Class for displaying the standard Print dialog box.
'--
'-- Notes..........: Use this class to provide your application the possibility to set the
'--                  printing options.
'--
'--
'-- Author, date...: Gary Noble (TDLcom) , 16 March 2002
'--
'--
'-- Property             Data Type     Description
'-- ------------------   ---------     --------------------------------------
'-- Copies               Integer       Returns or sets a value that determines the number of copies to be printed.
'-- Flags                EPrintFlags   Returns or sets the flags specifying the options for the standard Print dialog box.
'-- FromPage             Integer       Returns or sets the value for the first page to be printed.
'-- HDC                  Long          Returns or sets a handle to the printer's device context.
'-- HwndOwner            Long          Returns or sets a handle to the owner window for the standard Print dialog box.
'-- Max                  Integer       Returns or sets the maximum print range.
'-- Min                  Integer       Returns or sets the minimum print range.
'-- PrintRange           EPrintRange   Returns or sets the print range for the standard Print dialog box.
'-- ToPage               Integer       Returns or sets the value for the last page to be printed.
'--
'-- Method(Public)       Description
'-- ------------------   --------------------------------------
'-- INIT                 Use this routine for basic object initialization.
'-- FarProc              Returns the address of the standard Print dialog box callback proc.
'--                      This function is necessary because AddressOff can be used only as a
'--                      parameter part, i.e. it will be used in something like
'--                      m_Address = FarProc(AddressOff ProcName)
'-- Show                 Displays the standard Print dialog box.
'--                      Returns True if OK was selected, False otherwise.
'--
'-- Method(Private)      Description
'-- ------------------   --------------------------------------
'------------------------------------------------------------------------------

Option Explicit

'-- Error handling definitions
Private Const E_ERR_BASE = 17940 + vbObjectError
Public Enum EErrPrintDialog
    eErrPrintDialog_InvalidWindowHandle = E_ERR_BASE + 1
    eErrPrintDialog_ComponentFailure
End Enum
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "IAPP_PrintDialogHandler component failure"

'-- Private class type definitions
Private Type PrintDlg
    lStructSize As Long
    HWndOwner  As Long
    hDevMode   As Long
    hDevNames  As Long
    HDC        As Long
    Flags      As Long
    nFromPage  As Integer
    nToPage    As Integer
    nMinPage   As Integer
    nMaxPage   As Integer
    nCopies    As Integer
    hInstance  As Long
    lCustData  As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'-- Private class API function declarations
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Public class enums
Public Enum EPrintFlags
    ePrintFlag_AllPages = &H0
    ePrintFlag_Selection = &H1
    ePrintFlag_PageNums = &H2
    ePrintFlag_NoSelection = &H4
    ePrintFlag_NoPageNums = &H8
    ePrintFlag_Collate = &H10
    ePrintFlag_PrintToFile = &H20
    ePrintFlag_PrintSetup = &H40
    ePrintFlag_NoWarning = &H80
    ePrintFlag_ReturnDc = &H100
    ePrintFlag_ReturnIc = &H200
    ePrintFlag_ReturnDefault = &H400
    ePrintFlag_ShowHelp = &H800
    ePrintFlag_EnablePrintHook = &H1000
    ePrintFlag_EnableSetupHook = &H2000
    ePrintFlag_EnablePrintTemplate = &H4000
    ePrintFlag_EnableSetupTemplate = &H8000
    ePrintFlag_EnablePrintTemplateHandle = &H10000
    ePrintFlag_EnableSetupTemplateHandle = &H20000
    ePrintFlag_UseDevModeCopies = &H40000
    ePrintFlag_UseDevModeCopiesAndCollate = &H40000
    ePrintFlag_DisablePrintToFile = &H80000
    ePrintFlag_HidePrintToFile = &H100000
    ePrintFlag_NoNetworkButton = &H200000
End Enum

Public Enum EPrintRange
    ePrintRange_All
    ePrintRange_PageNumbers
    ePrintRange_Selection
End Enum

'-- Private variables to hold property values
Private m_eFlags As EPrintFlags
Private m_ePrintRange As EPrintRange
Private m_iCopies As Integer
Private m_iFromPage As Integer
Private m_iMax As Integer
Private m_iMin As Integer
Private m_iToPage As Integer
Private m_lhdc As Long
Private m_lHWndOwner As Long


'------------------------------------------------------------------------------
'-- Sub         : Init
'-- Notes       : Use this routine for basic object initialization.
'------------------------------------------------------------------------------
Public Sub Init(Optional HWndOwner As Long = 0, Optional Flags As EPrintFlags)
    On Error GoTo hComponentFailure

    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrPrintDialog_InvalidWindowHandle, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = HWndOwner

    m_eFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Sub


'------------------------------------------------------------------------------
'-- Property    : Copies
'-- Notes       : Returns or sets a value that determines the number of copies to be
'--               printed.
'------------------------------------------------------------------------------
Public Property Let Copies(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iCopies = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Copies() As Integer
    On Error GoTo hComponentFailure

    Copies = m_iCopies

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Flags
'-- Notes       : Returns or sets the flags specifying the options for the standard Print
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let Flags(ByVal vData As EPrintFlags)
    On Error GoTo hComponentFailure

    m_eFlags = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Flags() As EPrintFlags
    On Error GoTo hComponentFailure

    Flags = m_eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : FromPage
'-- Notes       : Returns or sets the value for the first page to be printed.
'------------------------------------------------------------------------------
Public Property Let FromPage(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iFromPage = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get FromPage() As Integer
    On Error GoTo hComponentFailure

    FromPage = m_iFromPage

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : HDC
'-- Notes       : Returns or sets a handle to the printer's device context.
'------------------------------------------------------------------------------
Public Property Let HDC(ByVal vData As Long)
    On Error GoTo hComponentFailure

    m_lhdc = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get HDC() As Long
    On Error GoTo hComponentFailure

    HDC = m_lhdc

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : HWndOwner
'-- Notes       : Returns or sets a handle to the owner window for the standard Print
'--               dialog box.
'------------------------------------------------------------------------------
Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure

    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrPrintDialog_InvalidWindowHandle, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_InvalidWindowHandle
        End If
    End If

    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure

    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Max
'-- Notes       : Returns or sets the maximum print range.
'------------------------------------------------------------------------------
Public Property Let Max(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iMax = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Max() As Integer
    On Error GoTo hComponentFailure

    Max = m_iMax

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : Min
'-- Notes       : Returns or sets the minimum print range.
'------------------------------------------------------------------------------
Public Property Let Min(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iMin = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get Min() As Integer
    On Error GoTo hComponentFailure

    Min = m_iMin

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : PrintRange
'-- Notes       : Returns or sets the print range for the standard Print dialog box.
'------------------------------------------------------------------------------
Public Property Let PrintRange(ByVal vData As EPrintRange)
    On Error GoTo hComponentFailure

    m_ePrintRange = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get PrintRange() As EPrintRange
    On Error GoTo hComponentFailure

    PrintRange = m_ePrintRange

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Property    : ToPage
'-- Notes       : Returns or sets the value for the last page to be printed.
'------------------------------------------------------------------------------
Public Property Let ToPage(ByVal vData As Integer)
    On Error GoTo hComponentFailure

    m_iToPage = vData

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property

Public Property Get ToPage() As Integer
    On Error GoTo hComponentFailure

    ToPage = m_iToPage

    Exit Property

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Property


'------------------------------------------------------------------------------
'-- Function    : Show
'-- Notes       : Displays the standard Print dialog box.
'--               Returns True if OK was selected, False otherwise.
'------------------------------------------------------------------------------
Public Function Show() As Boolean
    On Error GoTo hComponentFailure
    '------------------------------------------------------------------------------
    '-- Procedure Variables
    '------------------------------------------------------------------------------
    '-- Long
    '------------------------------------------------------------------------------
    Dim lFlags As Long
    Dim lMask  As Long
    '------------------------------------------------------------------------------
    '-- DataType
    '------------------------------------------------------------------------------
    Dim PD     As PrintDlg
    '------------------------------------------------------------------------------
    '-- End Variables
    '------------------------------------------------------------------------------




    If m_ePrintRange = ePrintRange_PageNumbers Then
        lFlags = lFlags Or ePrintFlag_PageNums
    ElseIf m_ePrintRange = ePrintRange_Selection Then
        lFlags = lFlags Or ePrintFlag_Selection
    End If

    lMask = CLng(Not (ePrintFlag_EnablePrintHook Or ePrintFlag_EnablePrintTemplate))
    lMask = lMask And CLng(Not (ePrintFlag_EnableSetupHook Or ePrintFlag_EnableSetupTemplate))

    PD.Flags = m_eFlags Or lFlags And lMask
    PD.HDC = m_lhdc
    PD.HWndOwner = m_lHWndOwner
    PD.lStructSize = Len(PD)
    PD.nCopies = m_iCopies
    PD.nFromPage = m_iFromPage
    PD.nMaxPage = m_iMax
    PD.nMinPage = m_iMin
    PD.nToPage = m_iToPage
    PD.nToPage = ToPage

    If PrintDlg(PD) Then

        Show = True

        m_iCopies = PD.nCopies
        m_iFromPage = PD.nFromPage
        m_lhdc = PD.HDC
        m_iMax = PD.nMaxPage
        m_iMin = PD.nMinPage

        If (PD.Flags And ePrintFlag_PageNums) Then
            m_ePrintRange = ePrintRange_PageNumbers
        ElseIf (PD.Flags And ePrintFlag_Selection) Then
            m_ePrintRange = ePrintRange_Selection
        Else
            m_ePrintRange = ePrintRange_All
        End If

        m_iToPage = PD.nToPage

    Else
        Show = False
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Function    : FarProc
'-- Notes       : Returns the address of the standard Print dialog box callback proc.
'--               This function is necessary because AddressOff can be used only as a
'--               parameter part, i.e. it will be used in something like
'--               m_Address = FarProc(AddressOff ProcName)
'------------------------------------------------------------------------------
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure

    FarProc = pfn

    Exit Function

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Function


'------------------------------------------------------------------------------
'-- Sub         : Class_Initialize
'-- Notes       : Class data space initialization.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure

    m_iCopies = 1
    m_eFlags = 0
    m_iFromPage = 0
    m_lHWndOwner = 0
    m_iMax = &HFFFF
    m_iMin = 1
    m_ePrintRange = ePrintRange_All
    m_iToPage = 0

    Exit Sub

hComponentFailure:
    Err.Raise eErrPrintDialog_ComponentFailure, App.EXEName & ".IAPP_PrintDialogHandler", S_ERR_ComponentFailure
End Sub
