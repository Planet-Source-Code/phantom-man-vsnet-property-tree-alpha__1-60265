VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_TextBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'pePropertyEditor
'//---------------------------------------------------------------------------------------
' Module    : IAPP_TextBoxEx
' DateTime  : 26/04/2005
' Author    : Gary Noble   ©2005 Telecom Direct Limited
' Purpose   : Simple Textbox Wrapper Functions
' Assumes   :
' Notes     :
' Revision  : v1.0
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  26/04/2005
'//---------------------------------------------------------------------------------------
Option Explicit

Implements ISubclass

' Possible performed checks on the text
Public Enum TxtExCheck
    TxtNone
    TxtAlfa
    TxtNumeric
    TxtCurrency
End Enum

Public Enum AlfaCase
    TxtUpperCase = 1
    TxtLowerCase = 2
End Enum

Public Enum TxtCntxMenu
    TxtMenuDefault
    TxtMenuNone
    TxtMenuCustom
End Enum

Private Const NUMBERS = "0123456789"
Private Const CHARACTERS_LCASE = "abcdefghijklmnopqrstvwxyz"
Private Const CHARACTERS_UCASE = "ABCDEFGHIJKLMNOPQRSTVWXYZ"

Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_UNDO As Long = &HC7
Private Const EM_SCROLLCARET As Long = &HB7
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_GETLINE As Long = &HC4
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_SETREADONLY = &HCF

Private Const GWL_WNDPROC = (-4)

Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303

Private Const GWL_STYLE = (-16)
Private Const ES_UPPERCASE = &H8&
Private Const ES_LOWERCASE = &H10&
Private Const ES_READONLY = &H800&

' NUMBERS
Private Const LOCALE_SDECIMAL As Long = &HE           'decimal separator
Private Const LOCALE_SPOSITIVESIGN As Long = &H50     'positive sign
Private Const LOCALE_SNEGATIVESIGN As Long = &H51     'negative sign

'CURRENCY
Private Const LOCALE_SCURRENCY As Long = &H14         'local symbol
Private Const LOCALE_SMONDECIMALSEP As Long = &H16    'decimal separator
Private Const LOCALE_SMONTHOUSANDSEP As Long = &H17   'thousand separator

Private WithEvents myTextBox As TextBox
Attribute myTextBox.VB_VarHelpID = -1

Private DecimalSep As String
Private NegativeSign As String
Private PositiveSign As String

Private currSymbol As String
Private ThousandCurrencySep As String
Private DecimalCurrencySep As String

Private m_bSelectOnEntry As Boolean
Private m_bEnterLoseFocus As Boolean

Private customMenu As Menu
Private m_menuType As TxtCntxMenu

Private m_bPasteEnabled As Boolean
Private m_Check As TxtExCheck
Private m_Case As AlfaCase

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" _
                                     Alias "SendMessageA" _
                                     (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" _
                                       Alias "GetLocaleInfoA" _
                                       (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
                                          (ByVal hwndLock As Long) As Long

' This has to be called before any other procedure. It sets the reference to the textbox
' used in the whole control
'
Property Set TextBoxRef(theTextBox As TextBox)
    Set myTextBox = theTextBox
End Property

Property Let SelectOnEntry(choice As Boolean)
    m_bSelectOnEntry = choice
End Property

Property Get SelectOnEntry() As Boolean
    SelectOnEntry = m_bSelectOnEntry
End Property

Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String
    Dim theString As String
    Dim retCode As Long

    retCode = GetLocaleInfo(dwLocaleID, dwLCType, theString, Len(theString))

    If retCode Then

        theString = Space$(retCode)

        retCode = GetLocaleInfo(dwLocaleID, dwLCType, theString, Len(theString))

        If retCode Then

            GetUserLocaleInfo = Left$(theString, retCode - 1)

        End If

    End If

End Function

Private Sub Class_Initialize()
    Dim LCID   As Long

    LCID = GetSystemDefaultLCID()

    'Determine number format depending on system's locale setting
    DecimalSep = GetUserLocaleInfo(LCID, LOCALE_SDECIMAL)
    NegativeSign = GetUserLocaleInfo(LCID, LOCALE_SNEGATIVESIGN)
    PositiveSign = GetUserLocaleInfo(LCID, LOCALE_SPOSITIVESIGN)

    DecimalCurrencySep = GetUserLocaleInfo(LCID, LOCALE_SMONDECIMALSEP)
    ThousandCurrencySep = GetUserLocaleInfo(LCID, LOCALE_SMONTHOUSANDSEP)
    currSymbol = GetUserLocaleInfo(LCID, LOCALE_SCURRENCY)

    m_bSelectOnEntry = False
    m_menuType = TxtMenuDefault
    m_bPasteEnabled = True
End Sub


Public Sub PerformCheck(check As TxtExCheck, Optional myCase As AlfaCase = TxtLowerCase Or TxtUpperCase)

    m_Check = check
    m_Case = myCase

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' This is used to subclass the textbox in order to change the context menu and the
' paste
    Select Case iMsg
        Case WM_CONTEXTMENU
            Debug.Print "Received WM_CONTEXTMENU at " & Now
            ' Nothing

        Case WM_PASTE
            Debug.Print "Received WM_PASTE at " & Now
            ' Nothing

    End Select

End Function

Private Sub myTextBox_GotFocus()
    If m_bSelectOnEntry Then
        myTextBox.SelStart = 0
        myTextBox.SelLength = Len(myTextBox.Text)
    End If
End Sub

' Determine if pressed key is allowed
'
Private Function IsKeyAllowed(keyStr As String) As Boolean
    Dim firstChar As Boolean

    firstChar = (Len(Trim(Left(myTextBox.Text, myTextBox.SelStart))) = 0)
    IsKeyAllowed = False

    Select Case m_Check
        Case TxtNone
            ' Everything is allowed
            IsKeyAllowed = True
            ' Nothing

        Case TxtAlfa
            If m_Case And TxtLowerCase Then
                IsKeyAllowed = (InStr(CHARACTERS_LCASE, keyStr) <> 0)
            End If
            If m_Case And TxtUpperCase Then
                IsKeyAllowed = IsKeyAllowed Or (InStr(CHARACTERS_UCASE, keyStr) <> 0)
            End If

        Case TxtNumeric
            ' Numeric formats is allowed if:
            ' - is a number
            ' - is a decimal separator (only one decimal separator is allowed in the whole textbox)
            ' - is a positive or negative sign. This is allowed only if it's the first character typed in the
            '   control.
            If InStr(NUMBERS, keyStr) <> 0 Then
                IsKeyAllowed = True

            ElseIf keyStr = DecimalSep Then
                IsKeyAllowed = (InStr(myTextBox.Text, DecimalSep) = 0)

            ElseIf keyStr = PositiveSign Or keyStr = NegativeSign Then
                IsKeyAllowed = firstChar And _
                               (InStr(myTextBox.Text, keyStr) = 0)
            End If

        Case TxtCurrency
            ' Currency format allowed if:
            ' - is a number
            ' - is a thousand separator, and is not the first char, and is not to the right of the decimal separator
            '   and is not near another thousand separator
            ' - is a decimal separator, and is not the first char, and the only decimal separator in the control
            If InStr(NUMBERS, keyStr) <> 0 Then
                IsKeyAllowed = True

            ElseIf keyStr = ThousandCurrencySep Then
                IsKeyAllowed = Not firstChar And _
                               (InStr(Left(myTextBox.Text, myTextBox.SelStart), DecimalCurrencySep) = 0) And _
                               (Right(myTextBox.Text, 1) <> ThousandCurrencySep)

            ElseIf keyStr = DecimalCurrencySep Then
                IsKeyAllowed = Not firstChar And _
                               (InStr(myTextBox.Text, DecimalCurrencySep) = 0) And _
                               (Right(myTextBox.Text, 1) <> ThousandCurrencySep)

            End If

    End Select
End Function

Private Sub myTextBox_KeyPress(KeyAscii As Integer)

    If m_bEnterLoseFocus And KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

    If KeyAscii <= 32 Then Exit Sub

    If myTextBox.MultiLine = True And KeyAscii = 13 Then Exit Sub

    If Not IsKeyAllowed(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If

End Sub

'
' Determine if the control loses focus on press of "enter" key
Property Let EnterLoseFocus(loseFocus As Boolean)
    If myTextBox.MultiLine = True Then
        m_bEnterLoseFocus = False
    Else
        m_bEnterLoseFocus = loseFocus
    End If
End Property

Property Get EnterLoseFocus() As Boolean
    EnterLoseFocus = m_bEnterLoseFocus
End Property

' Get line count
Property Get LineCount() As Long
    On Error Resume Next
    LineCount = SendMessage(myTextBox.hwnd, EM_GETLINECOUNT, 0&, ByVal 0&)
End Property

' Scroll text box until caret is visible
Public Sub ScrollCaret()
    SendMessage myTextBox.hwnd, EM_SCROLLCARET, 0, ByVal 0
End Sub

' Get current line number
Property Get CurrentLineNum() As Long
    On Error Resume Next
    CurrentLineNum = SendMessage(myTextBox.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
End Property

' Load a text file into the textbox
Public Sub LoadFile(FileName As String)
    Dim lFileLength As Long
    Dim iFileNum As Integer

    If FileName = "" Or Dir(FileName) = vbNullString Then
        Err.Raise 53                                  ' File not found
        Exit Sub
    End If

    'Get a free file number and open the file
    iFileNum = FreeFile
    Open FileName For Input As iFileNum

    'Get the length of the file and
    'read it into the text box
    lFileLength = LOF(iFileNum)
    myTextBox.Text = Input(lFileLength, #iFileNum)

    Close iFileNum

End Sub

'
'   Returns a single line in an edit control (multi-line)
'
Public Function GetLine(ByVal LineIndex As Long) As String
    Dim sBuffer As String
    Dim lReturn As Long

    sBuffer = String$(255, 0)
    Mid$(sBuffer, 1, 1) = Chr$(255)                   '// First word on the buffer is the buffer size

    GetLine = ""
    lReturn = SendMessage(myTextBox.hwnd, EM_GETLINE, LineIndex, ByVal sBuffer)
    If (lReturn > 0) Then _
       GetLine = Left$(sBuffer, lReturn)

    sBuffer = ""

End Function

'
'   Scrolls the control vertically and/or horizontally
'   by the specified amounts.
'
Public Sub LineScroll(Optional ByVal VerticalLines As Long = 0, Optional ByVal HorizontalChars As Long)

    If ((VerticalLines + HorizontalChars) > 0) Then _
       Call SendMessage(myTextBox.hwnd, EM_LINESCROLL, HorizontalChars, ByVal VerticalLines)

End Sub

'   Copies the control's current selection to
Public Sub Copy()

    Call SendMessage(myTextBox.hwnd, WM_COPY, 0, ByVal 0&)

End Sub

'   Cuts the control's current selection
Friend Sub Cut()

    Call SendMessage(myTextBox.hwnd, WM_CUT, 0, ByVal 0&)

End Sub

'   Pastes the contents of the clipboard
Friend Sub Paste()

    Call SendMessage(myTextBox.hwnd, WM_PASTE, 0, ByVal 0&)

End Sub

Friend Sub Undo()

    Call SendMessage(myTextBox.hwnd, EM_UNDO, 0, ByVal 0&)

End Sub

'   Removes all text from the control
Friend Sub Clear()

    Call SendMessage(myTextBox.hwnd, WM_CLEAR, 0, ByVal 0&)

End Sub

' With this property the control converts all inputs to UPPERCASE
'
Friend Property Get Uppercase() As Boolean

    Uppercase = (GetWindowLong(myTextBox.hwnd, GWL_STYLE) And ES_UPPERCASE)

End Property

Friend Property Let Uppercase(ByVal rValue As Boolean)

    If (rValue = True) Then
        Call SetWindowLong(myTextBox.hwnd, GWL_STYLE, GetWindowLong(myTextBox.hwnd, GWL_STYLE) Or ES_UPPERCASE)
    Else
        Call SetWindowLong(myTextBox.hwnd, GWL_STYLE, GetWindowLong(myTextBox.hwnd, GWL_STYLE) And Not ES_UPPERCASE)
    End If

End Property

' With this property the control converts all inputs to lowercase
'
Friend Property Get Lowercase() As Boolean

    Lowercase = (GetWindowLong(myTextBox.hwnd, GWL_STYLE) And ES_LOWERCASE)

End Property

Friend Property Let Lowercase(ByVal rValue As Boolean)

    If (rValue = True) Then
        Call SetWindowLong(myTextBox.hwnd, GWL_STYLE, GetWindowLong(myTextBox.hwnd, GWL_STYLE) Or ES_LOWERCASE)
    Else
        Call SetWindowLong(myTextBox.hwnd, GWL_STYLE, GetWindowLong(myTextBox.hwnd, GWL_STYLE) And Not ES_LOWERCASE)
    End If

End Property

' Make the control read-only (and not disabled!)
'
Friend Property Get ReadOnly() As Boolean

    ReadOnly = (GetWindowLong(myTextBox.hwnd, GWL_STYLE) And ES_READONLY)

End Property

Friend Property Let ReadOnly(ByVal rValue As Boolean)

    Call SendMessage(myTextBox.hwnd, EM_SETREADONLY, Abs(rValue), ByVal 0&)

End Property

' Change context menu. Available options:
' - Default menu
' - No menu
' - Custom menu
Public Sub SetContextMenu(menuType As TxtCntxMenu, Optional myMenu As Menu)

    If menuType = TxtMenuNone And m_menuType <> TxtMenuNone Then
        AttachMessage Me, myTextBox.hwnd, WM_CONTEXTMENU
    ElseIf m_menuType = TxtMenuNone And menuType <> TxtMenuNone Then
        DetachMessage Me, myTextBox.hwnd, WM_CONTEXTMENU
    End If

    m_menuType = menuType
    Select Case m_menuType
        Case TxtMenuDefault
            Set customMenu = Nothing

        Case TxtMenuNone
            '

        Case TxtMenuCustom
            Set customMenu = myMenu

    End Select
End Sub

' Use custom menu without sublclassing.
' See Microsoft Knowledge Base article Q191670
Private Sub myTextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton And m_menuType = TxtMenuCustom And Not customMenu Is Nothing Then
        LockWindowUpdate myTextBox.hwnd

        ' A disabled TextBox will not display a context menu
        myTextBox.Enabled = False

        ' Give the previous line time to complete
        DoEvents

        ' Display our own context menu
        myTextBox.Parent.PopupMenu customMenu

        ' Enable the control again
        myTextBox.Enabled = True

        ' Unlock updates
        LockWindowUpdate 0&
    End If

End Sub

' Enable / disable paste for the control
'
Property Let PasteEnabled(myEnable As Boolean)
    If myEnable = m_bPasteEnabled Then Exit Property
    m_bPasteEnabled = myEnable

    If m_bPasteEnabled Then
        DetachMessage Me, myTextBox.hwnd, WM_PASTE
    Else
        AttachMessage Me, myTextBox.hwnd, WM_PASTE
    End If
End Property

Property Get PasteEnabled() As Boolean
    PasteEnabled = m_bPasteEnabled
End Property



