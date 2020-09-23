VERSION 5.00
Begin VB.UserControl pePropertyTree 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   KeyPreview      =   -1  'True
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   Begin pePropertyEditor.isButton peClickButton 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Style           =   4
      Caption         =   "..."
      IconAlign       =   1
      iNonThemeStyle  =   4
      USeCustomColors =   -1  'True
      HighlightColor  =   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin pePropertyEditor.peColorSelector ctlDDPropSelector1 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   450
      BorderColor     =   6922635
   End
   Begin pePropertyEditor.peSpinButton CTLSpinButton1 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin pePropertyEditor.peListPicker peListPicker1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   529
      BorderColor     =   0
   End
   Begin VB.TextBox txtTextEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "pePropertyTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//---------------------------------------------------------------------------------------
'pePropertyEditor
'//---------------------------------------------------------------------------------------
' Module    : pePropertyTree
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : A Control To Simulate A Property Style Picker
' Assumes   :
' Notes     : ©2005 Gary Noble
' Revision  : 1.0
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Option Explicit

Private m_bUserMode As Boolean

'-- Api Scrollbars
'-- This Was Taken From One Of My projects - No Need For Manifest File
'-- This Class Will Auto Draw The Scrollbar Depending On The Theme
Private WithEvents m_cScrollBar As IAPP_ScrollBars
Attribute m_cScrollBar.VB_VarHelpID = -1

Private m_bDropDownPropertyShowing As Boolean
Private m_oDropDown As Object                         '-- Dropdown Object
Private m_oSplitter As IAPP_SplitterBar               '-- Splitter bar
Private m_LastIndexDrawn As Long                      '-- Sets the Lastindex Drawn
Private m_recSplitter As RECT                         '-- Splitter Rect
Private m_blnSpliiterMove As Boolean                  '-- Splitter Moving Flag
Private m_bInFocus As Boolean                         '-- Focus
Private m_lSelectedIndex As Long                      '-- Select Item Index
Private m_lLastX As Long                              '-- Mouse Move CoOrdinates
Private m_lLastY As Long                              '-- Mouse Move CoOrdinates
Private m_oTextBox As IAPP_TextBoxEx                  '-- Textbox Handler

'-- Item Collection
Dim m_ColPropertyItems As CPropertyItems              '-- Property Items
Dim m_ColListItems As ListItems                       '-- List Items Collection

'-- Property Pointers Used For Painting Display PropertyItems
Private Type tDisplayPropertyData
    ptr        As Long
    yStart     As Long
    rcData     As RECT
End Type

'-- Display nodes
Dim m_DisplayPropertyItemPTRs() As tDisplayPropertyData

'-- Redraw
Const m_def_LockWindowUpdate = True
Dim m_LockWindowUpdate As Boolean

'-- Default Left Indentation
'-- You Can Increase This In Muliples Of 2
Private Const m_lconst_PropertyItemIndentation As Long = 16

Private m_lDefSplitterLeft As Long                    '-- Default Splitter Position
Private m_lStart As Long                              '-- Item StartY Values
Private m_lItemHeight As Long                         '-- Default Item Height
Private m_oDraw As IAPP_DrawDC                        '-- Drawing DC

'-- Line color
Const m_def_LineColor = vbButtonShadow
Dim m_LineColor As OLE_COLOR

'-- Sidebar Color
Const m_def_SidebarColor = vbButtonFace
Dim m_SidebarColor As OLE_COLOR

'-- Open Close Button Colors
Const m_def_OpenCloseButtonColor = 0
Dim m_OpenCloseButtonColor As OLE_COLOR

Const m_def_BorderStyle = 1
Const m_def_PushPropertyCaptionsVisible = True
Const m_def_ShowPushedPropertyIndicators = True

Dim m_BorderStyle As Long
Dim m_PushPropertyCaptionsVisible As Boolean
Dim m_ShowPushedPropertyIndicators As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20                 '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200               '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

'-- Border Style
Public Enum EAppearanceConstants
    eacFlat
    eac3D
End Enum

Public Enum EBorderStyleConstants
    ebscNone
    ebscFixedSingle
    ebscThin
    ebscRaised
End Enum

Private m_eAppearance As EAppearanceConstants
Private m_eBorderStyle As EBorderStyleConstants

Public Event PropertyChanged(oPropItem As CPropertyItem)
Public Event PushButtonclick(oPropItem As CPropertyItem)
Public Event BeforeUpdate(oPropItem As CPropertyItem, newValue As Variant, bCancel As Boolean)
Public Event PropertySelected(oPropItem As CPropertyItem)

Private Const m_def_NormalForeColor = vbWindowText
Private Const m_def_SelectedForeColor = vbWindowText
Private Const m_def_SelectedBackColor = vbButtonFace

Dim m_NormalForeColor As OLE_COLOR
Dim m_SelectedForeColor As OLE_COLOR
Dim m_SelectedBackColor As OLE_COLOR





Public Property Get Appearance() As EAppearanceConstants
    Appearance = m_eAppearance
End Property
Public Property Let Appearance(ByVal eStyle As EAppearanceConstants)
    m_eAppearance = eStyle
    pvSetBorder
    PropertyChanged "Appearance"
End Property
Public Property Get BorderStyle() As EBorderStyleConstants
    BorderStyle = m_eBorderStyle
End Property
Public Property Let BorderStyle(ByVal eStyle As EBorderStyleConstants)
    m_eBorderStyle = eStyle
    pvSetBorder
    PropertyChanged "BorderStyle"
End Property
'//---------------------------------------------------------------------------------------
' Procedure : pvSetBorder
' Type      : Sub
' DateTime  : 26/04/2005
' Author    : Gary Noble
' Purpose   : Sets The Window Border Style
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  26/04/2005
'//---------------------------------------------------------------------------------------
Private Sub pvSetBorder()
    Dim lS     As Long

    UserControl.BorderStyle() = 0
    If m_eAppearance = eacFlat Then
        ' Flat border
        pvSetWinStyle GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
        If m_eBorderStyle > ebscNone Then
            pvSetWinStyle GWL_STYLE, WS_BORDER, 0
        Else
            pvSetWinStyle GWL_STYLE, 0, WS_BORDER
        End If
    Else
        ' 3d border
        Select Case m_eBorderStyle
            Case ebscNone
                ' No borders
                pvSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
                pvSetWinStyle GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
            Case ebscFixedSingle
                ' Default border:
                pvSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
                pvSetWinStyle GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
            Case ebscThin
                ' Thin style
                pvSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
                pvSetWinStyle GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
            Case ebscRaised
                pvSetWinStyle GWL_STYLE, WS_BORDER Or WS_THICKFRAME, 0
                pvSetWinStyle GWL_EXSTYLE, WS_EX_WINDOWEDGE, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        End Select
    End If

    pvRefreshData

End Sub
'//---------------------------------------------------------------------------------------
' Procedure : pvSetWinStyle
' Type      : Sub
' DateTime  : 26/04/2005
' Author    : Gary Noble
' Purpose   : Sets The Actual Border Style
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  26/04/2005
'//---------------------------------------------------------------------------------------
Private Sub pvSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)
    Dim lS     As Long
    Dim lhWnd  As Long
    lhWnd = UserControl.hwnd
    lS = GetWindowLong(lhWnd, lType)
    lS = lS And Not lStyleNot
    lS = lS Or lStyle
    SetWindowLong lhWnd, lType, lS
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    pvDrawItems

End Property

'//---------------------------------------------------------------------------------------
' Procedure : Clear
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Well - It Clears The Items And Created item Lists
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Sub Clear(Optional bIncludeLists As Boolean)

    If bIncludeLists Then
        If Not m_ColListItems Is Nothing Then
            Set m_ColListItems = Nothing
        End If
    End If
    
    Set m_ColListItems = Nothing

    If Not m_ColPropertyItems Is Nothing Then
        Set m_ColPropertyItems = Nothing
        Set m_ColPropertyItems = New CPropertyItems
        m_ColPropertyItems.Init ObjPtr(Me), hwnd
    End If

    Cls
    DoEvents
    m_lSelectedIndex = 0
    m_LastIndexDrawn = 0
    pvRefreshData
    m_cScrollBar.Visible(efsVertical) = False

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : ctlDDPropSelector1_Click
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Updates The Selected Property With The Colour Value
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub ctlDDPropSelector1_Click()

    pvUpDatePropertyValue vValue:=ctlDDPropSelector1.SelectedColor
    Debug.Print ctlDDPropSelector1.SelectedColor


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawItems
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws The Actual Control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawItems()

    Dim CurrentPropertyItem As CPropertyItem
    Dim CurrentPropertyItemParent As CPropertyItem
    Dim I      As Long
    Dim lY     As Long
    Dim lX     As Long
    Dim lStart As Long

    '-- Bail

    If m_LockWindowUpdate Then Exit Sub

    '-- Bail
    If m_ColPropertyItems Is Nothing Then Cls: Exit Sub

    '-- Clear
    Cls

    '-- Draw the Background
    m_oDraw.FillGradient 0, 0, ScaleWidth, ScaleHeight, UserControl.BackColor, UserControl.BackColor
    lY = 1

    '-- Draw the Sidebar
    Call pvDrawSideBar

    '-- Determine The Start Index
    If Not m_cScrollBar Is Nothing Then
        If m_cScrollBar.Visible(efsVertical) Then
            lStart = m_cScrollBar.Value(efsVertical) + 1
        Else
            lStart = 1
        End If
    Else                                              'NOT NOT...
        lStart = 1
    End If

    For I = lStart To UBound(m_DisplayPropertyItemPTRs)

        '-- Sets The Current Drawing Property
        Set CurrentPropertyItem = PropItem(m_DisplayPropertyItemPTRs(I).ptr)

        '-- Return The Parent If Any
        Set CurrentPropertyItemParent = PropItem(CurrentPropertyItem.ParentPTR)

        '-- Indentation Routine
        '-- Depending On The Property Level We Will Offset The Left
        '-- Hand Side Of The Property By The - m_lconst_PropertyItemIndentation Value
        '-- Note: Here I Am Only Indenting If The Level Is More Than One(1)
        If Not CurrentPropertyItemParent Is Nothing Then
            If CurrentPropertyItem.ChildCount = 0 And CurrentPropertyItemParent.Level > 1 Then
                lX = ((m_lconst_PropertyItemIndentation) * CurrentPropertyItem.Level)    '- (m_lconst_PropertyItemIndentation / 2)
            Else                                      'NOT CURRENTPROPERTYITEM.CHILDCOUNT...
                lX = (m_lconst_PropertyItemIndentation) * CurrentPropertyItem.Level    'm_lconst_PropertyItemIndentation * IIf(CurrentPropertyItem.ChildCount = 0, 1, 0)
            End If
        Else                                          'NOT NOT...
            lX = (m_lconst_PropertyItemIndentation) + m_lconst_PropertyItemIndentation * IIf(CurrentPropertyItem.ChildCount = 0, 1, 0)
        End If

        '-- Set The Current Property Item ID
        CurrentPropertyItem.id = I

        '-- Set The Item Rectangle
        With m_DisplayPropertyItemPTRs(I).rcData
            .Top = lY
            .Bottom = lY + (m_lItemHeight) + 2
            .Left = 0
            .Right = ScaleWidth                       '- lX
        End With                                      'M_DISPLAYPROPERTYITEMPTRS(I).RCDATA

        '-- Draw The Lines
        pvDrawItemLines CurrentPropertyItem, CurrentPropertyItemParent, lX, lY

        '-- Draw The Property Caption And Value
        pvDrawItemAndData CurrentPropertyItem, CurrentPropertyItemParent, lY, lX

        '-- Draw The On Close Button
        pvDrawOpenCloseButton CurrentPropertyItem, CurrentPropertyItemParent, lY, lX

        '-- Offset The Y Drawing Offset
        lY = lY + (m_lItemHeight) + 2

        '-- Bail
        If lY > ScaleHeight Then Exit For

    Next

    '-- Set The Last index Drawn
    '-- Used For The Scrolling On Keyup/KeyDown etc
    m_LastIndexDrawn = I

    '-- Shows The Property Item Type
    '-- IE: Colour Picker etc.
    pvShowPropertyItem

    '-- Set The Splitter Co-Ordinates
    m_recSplitter.Bottom = lY

    '-- Set the Splitter Height
    m_oSplitter.SplitterBottomOffSet = IIf(lY > ScaleHeight, ScaleHeight, lY)


    '-- Draw The control
    m_oDraw.BitBlt HDC, 0, 0, , lY, , , vbSrcCopy
    
    '-- Clean Up
    Set CurrentPropertyItem = Nothing
    Set CurrentPropertyItemParent = Nothing

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : Font
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   :
' Returns   : Font
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

    Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set UserControl.Font = New_Font
    Set CTLSpinButton1.Font = New_Font

    PropertyChanged "Font"
    m_lItemHeight = (TextHeight("',q"))
    If m_lItemHeight < 10 Then m_lItemHeight = 10
    'If m_lItemHeight < cboEdit(0).Height Then m_lItemHeight = cboEdit(0).Height - 1

    If Not m_LockWindowUpdate Then pvRefreshData

End Property

'//---------------------------------------------------------------------------------------
' Procedure : HitTest
' Type      : Function
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Determines If We Are Over an Item Or An Item Open Close Button
' Returns   : CPropertyItem
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Function HitTest(xOffset As Long, yOffset As Long, bIsChev As Boolean) As CPropertyItem

'On Error Resume Next

    Dim I      As Long
    Dim lX     As Long
    Dim lY     As Long
    Dim lStart As Long
    Dim PropertyItemParent As CPropertyItem

    '-- Start Value

    If Not m_cScrollBar Is Nothing Then
        lStart = m_cScrollBar.Value(efsVertical) + 1
    Else                                              'NOT NOT...
        lStart = 1
    End If

    '-- Loop Through The Items And Determine If We Are In The Item Or Not
    For I = lStart To UBound(m_DisplayPropertyItemPTRs)

        If PtInRect(m_DisplayPropertyItemPTRs(I).rcData, xOffset, yOffset) Then

            Set HitTest = PropItem(m_DisplayPropertyItemPTRs(I).ptr)


            
            Set PropertyItemParent = PropItem(HitTest.ParentPTR)

            '-- Are We in The Chevron
            If Not PropertyItemParent Is Nothing Then
                If HitTest.ChildCount = 0 And PropertyItemParent.Level > 1 Then
                    lX = ((m_lconst_PropertyItemIndentation) * HitTest.Level)
                Else                                  'NOT HITTEST.CHILDCOUNT...
                    lX = (m_lconst_PropertyItemIndentation) * HitTest.Level
                End If
            Else                                      'NOT NOT...
                lX = (m_lconst_PropertyItemIndentation) + m_lconst_PropertyItemIndentation * IIf(HitTest.ChildCount = 0, 1, 0)
            End If

            If m_lDefSplitterLeft < lX Then lX = m_lconst_PropertyItemIndentation

            If xOffset >= lX - 12 And xOffset <= lX - 4 Then
                bIsChev = True
                If (HitTest.ReAdjustedOpenCloseButton And HitTest.Level > 0) Then bIsChev = True
            End If

            Exit For                                  'loop varying i
        End If

    Next

    Set PropertyItemParent = Nothing

    On Error GoTo 0

End Function

Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd

End Property

'//---------------------------------------------------------------------------------------
' Procedure : LineColor
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Sets The Display Line Color
' Returns   : OLE_COLOR
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Sets The Grid Line Color"

    LineColor = m_LineColor

End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)

    m_LineColor = New_LineColor
    PropertyChanged "LineColor"
    peClickButton.HighlightColor = New_LineColor
    
    UserControl.peListPicker1.Redraw
    UserControl.ctlDDPropSelector1.Redraw

    pvDrawItems

End Property

'//---------------------------------------------------------------------------------------
' Procedure : ListItems
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Hold A List Of Drop Down Items
' Returns   : ListItems
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Get ListItems() As ListItems

    If m_ColListItems Is Nothing Then Set m_ColListItems = New ListItems
    Set ListItems = m_ColListItems

End Property

Public Property Get LockWindowUpdate() As Boolean
Attribute LockWindowUpdate.VB_Description = "Enable/Disable The Property Sheet From Drawing"

    LockWindowUpdate = m_LockWindowUpdate

End Property

'//---------------------------------------------------------------------------------------
' Procedure : LockWindowUpdate
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Locks The Window From Drawing
' Returns   : Boolean
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Let LockWindowUpdate(ByVal New_LockWindowUpdate As Boolean)

    m_LockWindowUpdate = New_LockWindowUpdate

    If New_LockWindowUpdate = False Then
        pvRefreshData
    End If

    PropertyChanged "LockWindowUpdate"

End Property

Private Sub ctlDDPropSelector1_GotFocus()

    UserControl_GotFocus

End Sub

Private Sub ctlDDPropSelector1_Hide()

    m_bDropDownPropertyShowing = False
    UserControl_MouseDown 1, 0, CSng(m_lLastX), CSng(m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top)
    Debug.Print "Hide"

End Sub

Private Sub ctlDDPropSelector1_Show()

    m_bDropDownPropertyShowing = True
    Debug.Print "Show"

End Sub

Private Sub CTLSpinButton1_Capture()
    m_bDropDownPropertyShowing = True
End Sub

Private Sub ctlSpinButton1_DownClick()

    pvUpDatePropertyValue vValue:=CTLSpinButton1.Value

End Sub

Private Sub ctlSpinButton1_GotFocus()

    m_bDropDownPropertyShowing = True
    m_bInFocus = True
    pvDrawItems

End Sub

Private Sub ctlSpinButton1_Mousemove()

    UserControl.MousePointer = vbDefault

End Sub

Private Sub CTLSpinButton1_ReleaseCapture()
    m_bDropDownPropertyShowing = False: Debug.Print "SSss"
End Sub

Private Sub ctlSpinButton1_UpClick()

    pvUpDatePropertyValue vValue:=CTLSpinButton1.Value

End Sub

Private Sub m_cScrollBar_AllowScroll(bAllow As Boolean)
    If Not m_bDropDownPropertyShowing Then
        bAllow = True
    Else
        bAllow = False
    End If


End Sub

'//---------------------------------------------------------------------------------------
' Scrollbar Events
'//---------------------------------------------------------------------------------------

Private Sub m_cScrollBar_Change(eBar As EFSScrollBarConstants)

    DoEvents
    pvDrawItems

End Sub

Private Sub m_cScrollBar_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)


    If m_cScrollBar.Visible(efsVertical) Then
        DoEvents
        pvDrawItems
    End If

End Sub

Private Sub m_cScrollBar_Scroll(eBar As EFSScrollBarConstants)
    If Not m_bDropDownPropertyShowing Then
        DoEvents
        pvDrawItems
    End If

End Sub

Private Sub m_cScrollBar_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)

    DoEvents
    pvDrawItems

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : OpenCloseButtonColor
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Sets The Open Close ButtonColor
' Returns   : OLE_COLOR
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Get OpenCloseButtonColor() As OLE_COLOR
Attribute OpenCloseButtonColor.VB_Description = "Sets The Open Closed Button Color"

    OpenCloseButtonColor = m_OpenCloseButtonColor

End Property

Public Property Let OpenCloseButtonColor(ByVal New_OpenCloseButtonColor As OLE_COLOR)

    m_OpenCloseButtonColor = New_OpenCloseButtonColor
    PropertyChanged "OpenCloseButtonColor"
    Call pvDrawItems

End Property

'//---------------------------------------------------------------------------------------
' Procedure : PropertyItems
' Type      : Property
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : PropertyItem Collection
' Returns   : CPropertyItems
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Property Get PropertyItems() As CPropertyItems

    If m_ColPropertyItems Is Nothing Then
        Set m_ColPropertyItems = New CPropertyItems
        m_ColPropertyItems.Init ObjPtr(Me), hwnd
    End If

    Set PropertyItems = m_ColPropertyItems

End Property

'//---------------------------------------------------------------------------------------
' Procedure : PropItem
' Type      : Property
' DateTime  : 04/10/2004 m_lconst_PropertyItemIndentation:43
' Author    : Gary Noble
' Purpose   : Return The PropItem Object From The Pointer
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Friend Property Get PropItem(ptr As Long) As CPropertyItem

    Dim xPropItem As CPropertyItem

    CopyMemory xPropItem, ptr, 4

    Set PropItem = xPropItem

    CopyMemory xPropItem, 0&, 4

    ZeroMemory ObjPtr(xPropItem), 4&

End Property

'//---------------------------------------------------------------------------------------
' Procedure : pvCheckDrawDC
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Makes Sure The Drawing Dc Is Enable And Of The Correct size
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvCheckDrawDC()

    If m_oDraw Is Nothing Then
        Set m_oDraw = New IAPP_DrawDC
        m_oDraw.Init ScaleWidth, ScaleHeight, UserControl.HDC
        m_oDraw.BackStyle = BS_TRANSPARENT
        m_oDraw.BackColor = UserControl.BackColor
    Else                                              'NOT M_ODRAW...
        m_oDraw.Destroy
        m_oDraw.Init ScaleWidth, ScaleHeight, UserControl.HDC
        m_oDraw.BackStyle = BS_TRANSPARENT
        m_oDraw.BackColor = UserControl.BackColor
    End If

    Set m_oDraw.Font = UserControl.Font

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvClearControls
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Clears The Controls Contained Within This Control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvClearControls()

    On Error Resume Next

    Dim ctl    As Control

    For Each ctl In UserControl.Controls
        ctl.Visible = False
    Next

    Set m_oDropDown = Nothing

    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawItemAndData
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws The Item Caption And Data Value
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawItemAndData(oProp As CPropertyItem, Parent As CPropertyItem, yOffset As Long, xOffset As Long)

    On Error Resume Next

    '-- Parent Item
    Dim PropParent As CPropertyItem

    Set Parent = PropItem(oProp.ParentPTR)

    With m_oDraw

        '-- If The Item Is A Top Level Item Then Fill It With The
        '-- Sidebar Color
        If oProp.Level = 0 Then
            .FillRect m_lconst_PropertyItemIndentation - 1, yOffset, ScaleWidth - 1, yOffset + m_lItemHeight + 1, IIf(oProp.Enabled, m_SidebarColor, vbButtonFace)
        End If

        '-- Draw The selected item
        If m_lSelectedIndex = oProp.id Then
            m_oDraw.ForeColor = IIf(Enabled, m_SelectedForeColor, vbGrayText)
            m_oDraw.FillRect m_lconst_PropertyItemIndentation + 1, yOffset, IIf((oProp.ChildCount > 0 And Not oProp.OverrideGroupHeaderHideItem) Or oProp.PropertyType = peReadOnlyPushButton, ScaleWidth - 1, m_lDefSplitterLeft), yOffset + m_lItemHeight + 1, IIf(m_bInFocus, m_SelectedBackColor, vbButtonFace)
        Else                                          'NOT M_LSELECTEDINDEX...
            m_oDraw.ForeColor = IIf(Enabled, m_NormalForeColor, vbGrayText)
            m_oDraw.Font.Bold = False
        End If
        If Not oProp.Enabled Then m_oDraw.ForeColor = vbGrayText
        
        If oProp.Level > 0 Then
            

            .DrawText oProp.Caption, IIf(xOffset > m_lDefSplitterLeft - 5, IIf(m_PushPropertyCaptionsVisible, m_lconst_PropertyItemIndentation + IIf(oProp.ChildCount > 0, m_lconst_PropertyItemIndentation - 5, 5), xOffset + 2), xOffset + 2), yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), m_lDefSplitterLeft, yOffset + m_lItemHeight + 2, DT_LEFT

            If m_ShowPushedPropertyIndicators Then
                If (xOffset > m_lDefSplitterLeft - 5) Then
                    If oProp.ChildCount = 0 Then
                        DrawHiddenPropertyArrow .HDC, 0, yOffset, m_lconst_PropertyItemIndentation, m_lItemHeight, True, BlendColor(m_SidebarColor, vbBlack, 200), True
                    End If
                End If
            End If

        Else                                          'NOT OPROP.LEVEL...
            .DrawText oProp.Caption, IIf(xOffset > m_lDefSplitterLeft - 5, m_lconst_PropertyItemIndentation + IIf(oProp.ChildCount > 0, m_lconst_PropertyItemIndentation - 5, 5), xOffset + 2), yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth, yOffset + m_lItemHeight + 2, DT_LEFT
        End If

        If oProp.ChildCount = 0 Or oProp.OverrideGroupHeaderHideItem Then

            m_oDraw.ForeColor = IIf(Enabled, vbWindowText, vbGrayText)
            
            If Not oProp.Enabled Then m_oDraw.ForeColor = vbGrayText
            
            'If Not oProp.Enabled Then
             '    m_oDraw.ForeColor = vbGrayText
            'elseIf Not Parent Is Nothing Then
             '   If Not Parent.Enabled Then m_oDraw.ForeColor = vbGrayText
            'End If

            
            Select Case oProp.PropertyType
                Case Is = peEditorType.peColorEditor
                    .FillRect m_lDefSplitterLeft + 3, yOffset + 2, m_lDefSplitterLeft + 15, yOffset + m_lItemHeight - 2, 0
                    .FillRect m_lDefSplitterLeft + 4, yOffset + 3, m_lDefSplitterLeft + 14, yOffset + m_lItemHeight - 3, oProp.Value
                    .DrawText Format(oProp.Value, oProp.DisplayFormat), m_lDefSplitterLeft + 18, yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth - 1, yOffset + m_lItemHeight + 2, DT_LEFT
                Case Is = peEditorType.pePictureEditor
                    If Not oProp.Value = Empty Then
                        .DrawText "(Picture)", m_lDefSplitterLeft + 2, yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth - 1, yOffset + m_lItemHeight + 2, DT_LEFT
                    End If
                Case Is = peEditorType.peFontEditor
                    If Not oProp.Value = Empty Then
                        .DrawText "(" & oProp.Value & " " & oProp.Value.Size & ")", m_lDefSplitterLeft + 2, yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth - 1, yOffset + m_lItemHeight + 2, DT_LEFT
                    End If
                Case Else
                    .DrawText Format(oProp.Value, oProp.DisplayFormat), m_lDefSplitterLeft + 2, yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth - 1, yOffset + m_lItemHeight + 2, DT_LEFT
            End Select

        Else                                          'NOT OPROP.CHILDCOUNT...
            If oProp.Level > 0 Then
                .DrawText oProp.Value, m_lDefSplitterLeft + 2, yOffset + ((m_lItemHeight \ 2) - (TextHeight("q,") \ 2)), ScaleWidth - 1, yOffset + m_lItemHeight + 2, DT_LEFT
            End If                                    '

        End If

    End With                                          'M_ODRAW

    Set PropParent = Nothing

    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawItemLines
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws The Propterty Item Box
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawItemLines(oProp As CPropertyItem, Parent As CPropertyItem, xOffset As Long, yOffset As Long)

    With m_oDraw

        .DrawLine m_lconst_PropertyItemIndentation + 1, yOffset + m_lItemHeight + 1, ScaleWidth, yOffset + m_lItemHeight + 1, m_LineColor
        .DrawLine m_lconst_PropertyItemIndentation, yOffset - 1, m_lconst_PropertyItemIndentation, yOffset + m_lItemHeight + 2, m_LineColor
        .DrawLine ScaleWidth - 1, yOffset - 1, ScaleWidth - 1, yOffset + m_lItemHeight + 2, m_LineColor

        If oProp.id > 1 Then
            .DrawLine m_lconst_PropertyItemIndentation + 1, yOffset - 1, ScaleWidth, yOffset - 1, m_LineColor
        Else                                          'NOT OPROP.ID...
            .DrawLine m_lconst_PropertyItemIndentation, yOffset - 1, ScaleWidth - 1, yOffset - 1, m_SidebarColor
        End If

        If oProp.Level > 0 Then
            .DrawLine m_lDefSplitterLeft, yOffset, m_lDefSplitterLeft, yOffset + m_lItemHeight + 2, m_LineColor
        End If

    End With                                          'M_ODRAW

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawOpenCloseButton
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws The Open Close Button
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawOpenCloseButton(oProp As CPropertyItem, Parent As CPropertyItem, yOffset As Long, xOffset As Long)

    On Error Resume Next

    Dim lGlymphColor As OLE_COLOR

    lGlymphColor = m_oDraw.TranslateColor(IIf(Enabled, m_OpenCloseButtonColor, vbGrayText))

    '-- Set The open Close Button On The Sidebar
    '-- If The Splitter If Less Than The xOffset

    If xOffset > m_lDefSplitterLeft - 5 Then
        oProp.ReAdjustedOpenCloseButton = True
        xOffset = m_lconst_PropertyItemIndentation
        lGlymphColor = m_oDraw.TranslateColor(IIf(Enabled, BlendColor(m_OpenCloseButtonColor, vbButtonFace, 80), vbGrayText))
    Else
        oProp.ReAdjustedOpenCloseButton = False
    End If

    With m_oDraw

        If oProp.ChildCount > 0 And oProp.Expanded Then
            .Rectangle xOffset - 13, yOffset + ((m_lItemHeight) \ 2) - 4, xOffset - 4, yOffset + ((m_lItemHeight) \ 2) + 5, , , lGlymphColor
            .DrawLine xOffset - 11, yOffset + ((m_lItemHeight) \ 2), xOffset - 6, yOffset + ((m_lItemHeight) \ 2), lGlymphColor
        ElseIf oProp.ChildCount > 0 And Not oProp.Expanded Then    'NOT OPROP.CHILDCOUNT...
            .Rectangle xOffset - 13, yOffset + ((m_lItemHeight) \ 2) - 4, xOffset - 4, yOffset + ((m_lItemHeight) \ 2) + 5, , , lGlymphColor
            .DrawLine xOffset - 11, yOffset + ((m_lItemHeight) \ 2), xOffset - 6, yOffset + ((m_lItemHeight) \ 2), lGlymphColor
            .DrawLine xOffset - 9, yOffset + ((m_lItemHeight) \ 2) - 2, xOffset - 9, yOffset + ((m_lItemHeight) \ 2) + 3, lGlymphColor
        End If

    End With                                          'M_ODRAW

    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawSideBar
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Here Were Simply Drawing The Sidebar down The Left Hand side of The Control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawSideBar()

    With m_oDraw
        .FillRect 0, 0, m_lconst_PropertyItemIndentation, ScaleHeight, m_SidebarColor
    End With                                          'M_ODRAW

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvIsOverSplitter
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Are We Over The Splitter bar
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvIsOverSplitter(Bln As Boolean, xOffset As Long, yOffset As Long)

    Dim PropertyItem As CPropertyItem
    Dim bChev  As Boolean

    Bln = PtInRect(m_recSplitter, xOffset, yOffset)

    Set PropertyItem = Nothing

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvMakePropertyItemArrayPTRs
' Type      : Sub
' DateTime  : 04/10/2004 m_lconst_PropertyItemIndentation:40
' Author    : Gary Noble
' Purpose   : The Main Point
'             For Drawing To The Screen We Make An Array Of Pointers To The node Objects.
'             This Enables us To Draw Faster And Not Take Up So Much Resource.
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Private Sub pvMakePropertyItemArrayPTRs()

    On Error Resume Next

    Dim ss     As CPropertyItem
    Erase m_DisplayPropertyItemPTRs
    ReDim m_DisplayPropertyItemPTRs(0) As tDisplayPropertyData
    m_lStart = 0

    If m_ColPropertyItems Is Nothing Then Cls: Exit Sub

    '-- Bail
    If m_ColPropertyItems.Items.Count <= 0 Then Exit Sub

    '-- Loop Through The PropertyItems And Attach The Pointer To The Array

    For Each ss In m_ColPropertyItems.Items
        ReDim Preserve m_DisplayPropertyItemPTRs(UBound(m_DisplayPropertyItemPTRs) + 1)
        ss.id = -1
        '-- Set The Pointer
        With m_DisplayPropertyItemPTRs(UBound(m_DisplayPropertyItemPTRs))
            .ptr = ObjPtr(ss)
            .yStart = m_lStart
        End With                                      'M_DISPLAYPROPERTYITEMPTRS(UBOUND(M_DISPLAYPROPERTYITEMPTRS))

        m_lStart = m_lStart + m_lItemHeight + 2

        If ss.ChildCount > 0 And ss.Expanded Then
            '-- If The Node has Children The Do The Same
            pvMakePropertyItemArrayPTRsWithChildren xNode:=ss
        End If

    Next

    pvMakeScrollVisible

    pvDrawItems
    Set ss = Nothing

    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvMakePropertyItemArrayPTRsWithChildren
' Type      : Sub
' DateTime  : 04/10/2004 m_lconst_PropertyItemIndentation:43
' Author    : Gary Noble
' Purpose   : Same As pvMakePropertyItemArrayPTRs But Used For Child Items
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Private Sub pvMakePropertyItemArrayPTRsWithChildren(xNode As CPropertyItem)

    Dim ss     As CPropertyItem

    For Each ss In xNode.ChildProperties.Items
        ss.id = -1
        ReDim Preserve m_DisplayPropertyItemPTRs(UBound(m_DisplayPropertyItemPTRs) + 1)

        With m_DisplayPropertyItemPTRs(UBound(m_DisplayPropertyItemPTRs))
            .ptr = ObjPtr(ss)
            .yStart = m_lStart
        End With                                      'M_DISPLAYPROPERTYITEMPTRS(UBOUND(M_DISPLAYPROPERTYITEMPTRS))

        m_lStart = m_lStart + m_lItemHeight + 2

        If ss.ChildCount > 0 And ss.Expanded Then
            pvMakePropertyItemArrayPTRsWithChildren xNode:=ss
        End If

    Next

    Set ss = Nothing

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvMakeScrollVisible
' Type      : Sub
' DateTime  : 04/10/2004 m_lconst_PropertyItemIndentation:36
' Author    : Gary Noble
' Purpose   : Displays Or Hides The Scrollbars
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Private Sub pvMakeScrollVisible()

    Dim bShown As Boolean
    Dim I      As Long

    On Error Resume Next

    Dim tHeight As Long
    Dim xHeight As Long

        If UBound(m_DisplayPropertyItemPTRs) - 1 < Round((ScaleHeight) \ (m_lItemHeight + 2)) Then
        m_cScrollBar.Value(efsVertical) = 0
        m_cScrollBar.Max(efsVertical) = 0
        m_cScrollBar.Min(efsVertical) = 0
        m_cScrollBar.Visible(efsVertical) = False

        m_oSplitter.Border(espbright) = 20
    Else                                              'NOT UBOUND(M_DISPLAYPROPERTYITEMPTRS)...
        m_cScrollBar.Max(efsVertical) = (UBound(m_DisplayPropertyItemPTRs)) - Round((ScaleHeight) \ (m_lItemHeight + 2)) + 1
        m_cScrollBar.Visible(efsVertical) = True
        m_cScrollBar.LargeChange(efsVertical) = 30
        m_oSplitter.Border(espbright) = 40
    End If

    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvRefreshData
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Refreshes All The Data
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvRefreshData()

'-- Make the Display nodes

    Call pvMakePropertyItemArrayPTRs

    '-- Scrollbar?
    Call pvMakeScrollVisible

    '-- Set The Drawing DC
    Call pvCheckDrawDC

    '-- Draw The Control
    Call pvDrawItems

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvShowPropertyItem
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Shows the Property Item Type
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvShowPropertyItem()

    On Error Resume Next

    m_bDropDownPropertyShowing = False

    Dim SelectedPropItem As CPropertyItem

    Call pvClearControls

    If m_lSelectedIndex > 0 Then

        If txtTextEdit.Font <> UserControl.Font Then Set txtTextEdit.Font = UserControl.Font
        
        Set SelectedPropItem = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)

        If SelectedPropItem.ChildCount > 0 Then pvClearControls

        If SelectedPropItem.ChildCount = 0 Or SelectedPropItem.OverrideGroupHeaderHideItem Then
            Select Case SelectedPropItem.PropertyType
                Case pelistEditor

                    m_bDropDownPropertyShowing = False
                    peListPicker1.Move ScaleWidth - 16, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, 15, m_lItemHeight + 1
                    txtTextEdit.Move m_lDefSplitterLeft + 2, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, ScaleWidth - m_lDefSplitterLeft - 18, m_lItemHeight
                    peListPicker1.ListDataPTR = SelectedPropItem.ListPTR
                    Set m_oDropDown = peListPicker1


                    peListPicker1.Visible = True

                Case peColorEditor

                    ctlDDPropSelector1.Move ScaleWidth - 16, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, 17, m_lItemHeight + 1
                    ctlDDPropSelector1.SelectedColor = SelectedPropItem.Value
                    Set m_oDropDown = ctlDDPropSelector1
                    ctlDDPropSelector1.Visible = True

                Case peFileEditor, peFontEditor, pePictureEditor, peReadOnlyPushButton, peDirectoryEditor
                    
                    txtTextEdit.Move m_lDefSplitterLeft + 2, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, ScaleWidth - m_lDefSplitterLeft - 18, m_lItemHeight + 1
                    
                    If SelectedPropItem.PropertyType = pePictureEditor Then
                        txtTextEdit.Visible = True
                        txtTextEdit.Locked = True
                        If SelectedPropItem.Value = Empty Then
                            txtTextEdit.Text = "(Nothing)"
                            Else
                            txtTextEdit.Text = "(Picture)"
                        End If
                    ElseIf SelectedPropItem.PropertyType = peDirectoryEditor Or peFileEditor Then
                        txtTextEdit.Visible = True
                        txtTextEdit.Locked = True
                        txtTextEdit.Text = SelectedPropItem.Value
                    End If
                    
                    peClickButton.Move ScaleWidth - 16, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, 15, m_lItemHeight + 1
                    Set m_oDropDown = peClickButton
                    peClickButton.Visible = True
                
                Case peReadOnly
                    '-- Do Nothing
                Case peTextEditor

                    Set txtTextEdit.Font = UserControl.Font
                    txtTextEdit.Move m_lDefSplitterLeft + 2, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, ScaleWidth - m_lDefSplitterLeft - 3, m_lItemHeight
                    txtTextEdit.Text = SelectedPropItem.Value
                    txtTextEdit.Visible = True
                    txtTextEdit.Locked = False

                    Set m_oDropDown = txtTextEdit

                Case peSpinButton
                    'm_bDropDownPropertyShowing = True
                    CTLSpinButton1.Move m_lDefSplitterLeft + 1, m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top, ScaleWidth - m_lDefSplitterLeft - 2, m_lItemHeight
                    CTLSpinButton1.Value = SelectedPropItem.Value
                    CTLSpinButton1.Max = SelectedPropItem.SpinButtonConstraints(Max)
                    CTLSpinButton1.Min = SelectedPropItem.SpinButtonConstraints(Min)
                    CTLSpinButton1.IncermentValue = SelectedPropItem.SpinButtonConstraints(lIncrementValue)
                    Set m_oDropDown = CTLSpinButton1

            End Select
        End If

    Else                                              'NOT M_LSELECTEDINDEX...
        pvClearControls
    End If




    If Not SelectedPropItem Is Nothing Then
        If Not SelectedPropItem.Enabled Then
            
            If SelectedPropItem.PropertyType = peReadOnlyPushButton Then
                m_oDropDown.Visible = False
            Else
                m_oDropDown.Visible = False
                Set m_oDropDown = Nothing
            End If
            
            If SelectedPropItem.ChildCount = 0 And SelectedPropItem.OverrideGroupHeaderHideItem = False And Not SelectedPropItem.PropertyType = peReadOnlyPushButton Then
                m_oDropDown.Visible = False
                Set m_oDropDown = Nothing
            End If

         ElseIf SelectedPropItem.ReadOnly Then
                pvClearControls
                m_oDropDown.Visible = False
                Set m_oDropDown = Nothing
        End If
            
        
        If Not Me.Enabled Then
            pvClearControls
            m_oDropDown.Visible = False
            Set m_oDropDown = Nothing
        End If

            
    End If
    

    
    If Not m_oDropDown Is Nothing Then
        If m_lSelectedIndex < m_cScrollBar.Value(efsVertical) + 1 Or m_lSelectedIndex > m_LastIndexDrawn Then
            m_oDropDown.Visible = False
            pvClearControls
        Else                                          'NOT M_LSELECTEDINDEX...
            m_oDropDown.Visible = True
        End If
    End If
    
    On Error GoTo 0

End Sub

Public Property Get SidebarColor() As OLE_COLOR
Attribute SidebarColor.VB_Description = "Sets The Sidebar Color"

    SidebarColor = m_SidebarColor

End Property

Public Property Let SidebarColor(ByVal New_SidebarColor As OLE_COLOR)

    m_SidebarColor = New_SidebarColor
    PropertyChanged "SidebarColor"
    UserControl.peListPicker1.Redraw
    UserControl.ctlDDPropSelector1.Redraw
    peClickButton.BackColor = m_SidebarColor
    pvDrawItems

End Property

Private Sub peClickButton_Click()

    Select Case SelectedProperty.PropertyType

        Case Is = peReadOnlyPushButton
            RaiseEvent PushButtonclick(SelectedProperty)

        Case Is = peFileEditor

            Dim oEdit As New IAPP_FOSHandler
            With oEdit
                .HWndOwner = UserControl.hwnd
                .Filter = SelectedProperty.DialogFilter
                .FileName = SelectedProperty.Value

                .Show eDialogType_OpenFile

                If .FileCount > 0 Then
                    pvUpDatePropertyValue .FileName
                End If
            End With
            Set oEdit = Nothing
        Case Is = peFontEditor
            Dim oEditF As New IAPP_FontDialogHandler

            With oEditF
                
                If IsObject(SelectedProperty.Value) Then
                    .Init IIf(Len(SelectedProperty.Value) = 0, Nothing, SelectedProperty.Value), , UserControl.hwnd
                End If
               .HWndOwner = UserControl.hwnd
               .Flags = eFontFlag_Effects + eFontFlag_ForceFontExist
            
                If .Show Then
                    If .Font <> "" Then
                        pvUpDatePropertyValue .Font
                    End If
                Else
                    Debug.Print "Cancelled"
                End If
            
            End With


            Set oEditF = Nothing

        Case Is = pePictureEditor
            Dim oEditFD As New IAPP_FOSHandler

            With oEditFD
                .HWndOwner = UserControl.hwnd
                
                If IsObject(SelectedProperty.Value) Then
                    .Init SelectedProperty.Value, , UserControl.hwnd
                End If
                .Filter = SelectedProperty.DialogFilter
                .Show eDialogType_OpenFile

                If .FileCount <> 0 Then
                    pvUpDatePropertyValue .FileName
                End If
            End With


            Set oEditFD = Nothing

        Case Is = peDirectoryEditor
            Dim oEditD As New IAPP_BOFHandler

            With oEditD
                
                .Init SelectedProperty.Value
                .HWndOwner = UserControl.hwnd
                .Browse
                If .SelectedPath <> SelectedProperty.Value Then
                    pvUpDatePropertyValue .SelectedPath
                End If
            End With
            Set oEditD = Nothing

    End Select


End Sub

Private Sub peListPicker1_Hide()

    m_bDropDownPropertyShowing = False
    UserControl_MouseDown 1, 0, CSng(m_lLastX), CSng(m_DisplayPropertyItemPTRs(m_lSelectedIndex).rcData.Top)

End Sub

Private Sub peListPicker1_Show()


    m_bDropDownPropertyShowing = True

End Sub


Private Sub peListPicker1_ItemSelected(oText As String)
    pvUpDatePropertyValue vValue:=oText
End Sub

Private Sub txtTextEdit_GotFocus()
    m_bDropDownPropertyShowing = True
End Sub

Private Sub txtTextEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 38 Then
        m_bDropDownPropertyShowing = False
        UserControl_KeyDown KeyCode, Shift
        Exit Sub
    ElseIf KeyCode = 40 Then
        m_bDropDownPropertyShowing = False
        UserControl_KeyDown KeyCode, Shift
        Exit Sub
    End If
    
    If SelectedProperty.PropertyType = pePictureEditor Then
        If KeyCode = vbKeyDelete Then
            If txtTextEdit.SelLength = Len(txtTextEdit) Then
                pvUpDatePropertyValue ""
                pvDrawItems
            End If
        End If
    End If
    
    
End Sub

Private Sub txtTextEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    If Not Me.SelectedProperty Is Nothing Then
        
         If SelectedProperty.PropertyType <> pePictureEditor And Not SelectedProperty.ReadOnly Then
           pvUpDatePropertyValue txtTextEdit.Text
        End If
        RaiseEvent PropertyChanged(SelectedProperty)
    End If
    End If
    
    If KeyCode = 13 Then pvClearControls: pvDrawItems

End Sub

Private Sub txtTextEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.MousePointer = vbDefault
End Sub

Private Sub UserControl_Click()

    m_bDropDownPropertyShowing = False
    pvDrawItems


End Sub

Private Sub UserControl_DblClick()

    Dim bChev  As Boolean
    Dim oHittest As CPropertyItem

    m_blnSpliiterMove = False

    Set oHittest = HitTest(CLng(m_lLastX), CLng(m_lLastY), bChev)

    '-- Expand Or Collapse The Selected Property
    If Not oHittest Is Nothing Then

        If oHittest.ChildCount > 0 Then
            If oHittest.Level > 1 And m_lLastX > m_lconst_PropertyItemIndentation Then
                oHittest.Expanded = Not oHittest.Expanded
            ElseIf bChev Then
                oHittest.Expanded = Not oHittest.Expanded
            ElseIf oHittest.ReAdjustedOpenCloseButton Then
                oHittest.Expanded = Not oHittest.Expanded
            ElseIf m_lLastX > m_lconst_PropertyItemIndentation Then
                If m_lLastX > m_lDefSplitterLeft + 1 Then
                    If oHittest.ChildCount > 0 And oHittest.OverrideGroupHeaderHideItem Then
                        pvSimulatePropertyItemDoubleClick
                    Else
                        oHittest.Expanded = Not oHittest.Expanded
                    End If
                Else
                    oHittest.Expanded = Not oHittest.Expanded
                End If
            End If
            pvRefreshData
        Else                                          'NOT OHITTEST.CHILDCOUNT...
            If m_lLastX > m_lDefSplitterLeft + 1 Then pvSimulatePropertyItemDoubleClick
        End If

        m_lSelectedIndex = oHittest.id
        pvDrawItems

    End If

    Set oHittest = Nothing

End Sub

Private Sub UserControl_EnterFocus()


    pvDrawItems

End Sub

Private Sub UserControl_ExitFocus()

    m_bInFocus = False

    pvDrawItems

End Sub

Private Sub UserControl_GotFocus()

    m_bDropDownPropertyShowing = False
    m_bInFocus = True
    pvDrawItems

End Sub

Private Sub UserControl_Initialize()

    m_bDropDownPropertyShowing = False

    Set m_oSplitter = New IAPP_SplitterBar
    Set m_oTextBox = New IAPP_TextBoxEx

    With m_oSplitter
        .Orientation = espVertical
        .SplitObject = Me
        .Border(espbLeft) = 20
        .Border(espbright) = 20
    End With                                          'M_OSPLITTER

End Sub

Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font

    m_LockWindowUpdate = m_def_LockWindowUpdate
    m_LineColor = m_def_LineColor
    m_SidebarColor = m_def_SidebarColor
    m_OpenCloseButtonColor = m_def_OpenCloseButtonColor

    pvDrawSampleProperties

    m_ShowPushedPropertyIndicators = m_def_ShowPushedPropertyIndicators
    m_PushPropertyCaptionsVisible = m_def_PushPropertyCaptionsVisible
    m_BorderStyle = m_def_BorderStyle
    m_SelectedBackColor = m_def_SelectedBackColor
    m_NormalForeColor = m_def_NormalForeColor
    m_SelectedForeColor = m_def_SelectedForeColor
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim PropertyItem As CPropertyItem
    Dim bChev  As Boolean
    Dim lMaxMove As Long

    '-- Set The Max Move Allowed

    If KeyCode = vbKeyEscape And m_bDropDownPropertyShowing Then
        pvSimulatePropertyItemDoubleClick
        m_bDropDownPropertyShowing = False
    End If

    If m_bDropDownPropertyShowing Then Exit Sub

    lMaxMove = Round((ScaleHeight) \ (m_lItemHeight + 2))

    If KeyCode = vbKeyUp Then
        m_lSelectedIndex = m_lSelectedIndex - 1

    ElseIf KeyCode = vbKeyDown Then                   'NOT KEYCODE...
        m_lSelectedIndex = m_lSelectedIndex + 1

    ElseIf KeyCode = 34 Then                          'NOT KEYCODE...
        m_lSelectedIndex = m_lSelectedIndex + lMaxMove

    ElseIf KeyCode = 33 Then                          'NOT KEYCODE...
        m_lSelectedIndex = m_lSelectedIndex - lMaxMove

    ElseIf KeyCode = vbKeyRight Then                  'NOT KEYCODE...
        Set PropertyItem = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)

        If Not PropertyItem Is Nothing Then
            PropertyItem.Expanded = True
            pvRefreshData
        End If

    ElseIf KeyCode = vbKeyLeft Then                   'NOT KEYCODE...
        Set PropertyItem = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)
        If Not PropertyItem Is Nothing Then
            PropertyItem.Expanded = False
            pvRefreshData
        End If

    End If

    '-- Make Sure We Can Select A Property
    If m_lSelectedIndex >= UBound(m_DisplayPropertyItemPTRs) Then
        m_lSelectedIndex = UBound(m_DisplayPropertyItemPTRs)
    ElseIf m_lSelectedIndex <= 0 Then                 'NOT M_LSELECTEDINDEX...
        m_lSelectedIndex = 1
    End If

    '-- Make It Visible If Needed
    If m_cScrollBar.Visible(efsVertical) Then
        If m_lSelectedIndex > m_cScrollBar.Value(efsVertical) + Round((ScaleHeight) \ (m_lItemHeight + 2)) Then

            If KeyCode = 33 Or KeyCode = 34 Then
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + lMaxMove
            Else                                      'NOT KEYCODE...
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + 1
            End If
        ElseIf m_lSelectedIndex - 1 < m_cScrollBar.Value(efsVertical) Then    'NOT M_LSELECTEDINDEX...
            If KeyCode = 33 Or KeyCode = 34 Then
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - lMaxMove
            Else                                      'NOT KEYCODE...
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - 1
            End If

        End If
    End If

    Set PropertyItem = Nothing

    pvDrawItems
    RaiseEvent PropertySelected(Me.SelectedProperty)
    


End Sub

Private Sub UserControl_LostFocus()

    pvClearControls
    pvDrawItems

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    Dim bChev  As Boolean
    Dim oHittest As CPropertyItem
    Dim Bln    As Boolean

    Set m_oDropDown = Nothing

    m_bDropDownPropertyShowing = False

    Set oHittest = HitTest(CLng(X), CLng(Y), bChev)

    If Not m_blnSpliiterMove Then

        '-- Check If We Are Over The Splitter
        pvIsOverSplitter Bln:=Bln, xOffset:=CLng(X), yOffset:=CLng(Y)

        If Bln Then
            pvClearControls
            UserControl.MousePointer = 9
            m_blnSpliiterMove = True
            m_oSplitter.SplitterMouseDown hwnd, X, Y
            Exit Sub                                  '---> Bottom
        Else                                          'BLN = FALSE/0
            UserControl.MousePointer = vbDefault
        End If
    End If

    If Not oHittest Is Nothing Then

        If bChev Then
            oHittest.Expanded = Not oHittest.Expanded
            pvRefreshData
        End If

        m_lSelectedIndex = oHittest.id
        pvDrawItems
        RaiseEvent PropertySelected(Me.SelectedProperty)
        
    End If

    Set oHittest = Nothing

    'pvShowPropertyItem

    On Error GoTo 0

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not m_bInFocus Then Exit Sub

    Dim Bln    As Boolean
    Dim oHittest As CPropertyItem
    Dim bChev  As Boolean

    '-- Set The Last Mouse CoOrdinates
    m_lLastY = Y
    m_lLastX = X

    If Not m_blnSpliiterMove Then
        pvIsOverSplitter Bln:=Bln, xOffset:=CLng(X), yOffset:=CLng(Y)

        If Bln Then
            UserControl.MousePointer = 9
        Else                                          'BLN = FALSE/0
            UserControl.MousePointer = vbDefault
            m_blnSpliiterMove = False
        End If

    Else                                              'NOT NOT...

        '-- Move The Splitter
        m_oSplitter.SplitterContainer_MouseMove X, Y

    End If

    Set oHittest = Nothing

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_blnSpliiterMove Then
        m_oSplitter.SplitterContainer_MouseUp X, Y

        m_recSplitter.Left = X
        m_recSplitter.Right = X + 2
        m_lDefSplitterLeft = X
    End If

    m_blnSpliiterMove = False
    UserControl.MousePointer = vbDefault
    pvDrawItems
    MousePointer = vbDefault

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    If UserControl.Ambient.UserMode Then
        m_bUserMode = True
        Set m_cScrollBar = New IAPP_ScrollBars
        m_cScrollBar.Create UserControl.hwnd
        m_cScrollBar.Orientation = efsoVertical
        m_cScrollBar.Visible(efsVertical) = False
        m_cScrollBar.Visible(efsHorizontal) = False

    End If

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_LockWindowUpdate = PropBag.ReadProperty("LockWindowUpdate", m_def_LockWindowUpdate)
    m_lDefSplitterLeft = ScaleWidth \ 2

    With m_recSplitter
        .Top = 0
        .Bottom = ScaleHeight
        .Left = ScaleWidth \ 2
        .Right = (ScaleWidth \ 2) + 2
    End With                                          'M_RECSPLITTER

    m_LineColor = PropBag.ReadProperty("LineColor", m_def_LineColor)
    m_SidebarColor = PropBag.ReadProperty("SidebarColor", m_def_SidebarColor)
    m_OpenCloseButtonColor = PropBag.ReadProperty("OpenCloseButtonColor", m_def_OpenCloseButtonColor)
    m_eBorderStyle = PropBag.ReadProperty("BorderStyle", ebscThin)
    UserControl.ctlDDPropSelector1.BorderColor = vbBlack
    Set CTLSpinButton1.Font = UserControl.Font
    
    If m_oDraw Is Nothing Then Set m_oDraw = New IAPP_DrawDC
    
        peClickButton.HighlightColor = m_oDraw.TranslateColor(m_LineColor)
    
    '-- Set The Def Item Height
    m_lItemHeight = (TextHeight("',q"))
    If m_lItemHeight < 10 Then m_lItemHeight = 10

    'If m_lItemHeight < cboEdit(0).Height Then m_lItemHeight = cboEdit(0).Height - 1

    If Not m_bUserMode Then pvDrawSampleProperties

    m_ShowPushedPropertyIndicators = PropBag.ReadProperty("ShowPushedPropertyIndicators", m_def_ShowPushedPropertyIndicators)
    m_PushPropertyCaptionsVisible = PropBag.ReadProperty("PushPropertyCaptionsVisible", m_def_PushPropertyCaptionsVisible)
    m_eBorderStyle = PropBag.ReadProperty("BorderStyle", ebscThin)
    m_eAppearance = PropBag.ReadProperty("Appearance", eac3D)
    pvSetBorder

    m_SelectedBackColor = PropBag.ReadProperty("SelectedBackColor", m_def_SelectedBackColor)
    m_NormalForeColor = PropBag.ReadProperty("NormalForeColor", m_def_NormalForeColor)
    m_SelectedForeColor = PropBag.ReadProperty("SelectedForeColor", m_def_SelectedForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

Private Sub UserControl_Resize()

    pvRefreshData
    'pvShowPropertyItem

End Sub

Private Sub UserControl_Terminate()

    Erase m_DisplayPropertyItemPTRs
    Set m_cScrollBar = Nothing
    Set m_oDraw = Nothing
    Set m_oSplitter = Nothing
    Set m_ColPropertyItems = Nothing
    Set m_ColListItems = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("LockWindowUpdate", m_LockWindowUpdate, m_def_LockWindowUpdate)
    Call PropBag.WriteProperty("LineColor", m_LineColor, m_def_LineColor)
    Call PropBag.WriteProperty("SidebarColor", m_SidebarColor, m_def_SidebarColor)
    Call PropBag.WriteProperty("OpenCloseButtonColor", m_OpenCloseButtonColor, m_def_OpenCloseButtonColor)

    Call PropBag.WriteProperty("ShowPushedPropertyIndicators", m_ShowPushedPropertyIndicators, m_def_ShowPushedPropertyIndicators)
    Call PropBag.WriteProperty("PushPropertyCaptionsVisible", m_PushPropertyCaptionsVisible, m_def_PushPropertyCaptionsVisible)
    Call PropBag.WriteProperty("BorderStyle", m_eBorderStyle, ebscThin)
    Call PropBag.WriteProperty("Appearance", m_eAppearance, eac3D)


    Call PropBag.WriteProperty("SelectedBackColor", m_SelectedBackColor, m_def_SelectedBackColor)
    Call PropBag.WriteProperty("NormalForeColor", m_NormalForeColor, m_def_NormalForeColor)
    Call PropBag.WriteProperty("SelectedForeColor", m_SelectedForeColor, m_def_SelectedForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
End Sub

Private Sub VScroll1_Change()

    pvDrawItems

End Sub

Private Sub VScroll1_Scroll()

    pvDrawItems

End Sub

Private Sub pvSimulatePropertyItemDoubleClick()

    Dim SelProp As CPropertyItem

    Set SelProp = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)

    If Not SelProp Is Nothing Then
        If Not SelProp.ReadOnly Then
            Select Case SelProp.PropertyType
                Case Is = peColorEditor
                    UserControl.ctlDDPropSelector1.SimulateDropDown
                    m_bDropDownPropertyShowing = True
                Case Is = pelistEditor
                    If GetNextListItem(lstPTr:=SelectedProperty.ListPTR) > "" Then
                        SelProp.Value = GetNextListItem(lstPTr:=SelectedProperty.ListPTR)
                        RaiseEvent PropertyChanged(SelProp)
                        pvDrawItems
                    End If
            End Select

        End If
        pvDrawItems
    End If


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawSampleProperties
' Type      : Sub
' DateTime  : 20/04/2005
' Author    : Gary Noble
' Purpose   : Draws Some Sample Properties
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  20/04/2005
'//---------------------------------------------------------------------------------------

Private Sub pvDrawSampleProperties()

    If Not m_bUserMode Then

        '-- Load Some Sample Properties

        '-- Set The Def Item Height
        m_lItemHeight = (TextHeight("',q"))
        If m_lItemHeight < 10 Then m_lItemHeight = 10
        'If m_lItemHeight < cboEdit(0).Height Then m_lItemHeight = cboEdit(0).Height - 1

        If m_lDefSplitterLeft <> UserControl.ScaleWidth \ 2 Then
            m_lDefSplitterLeft = UserControl.ScaleWidth \ 2
        End If


        Set m_oDraw = New IAPP_DrawDC
        m_oDraw.Init ScaleWidth, ScaleHeight, UserControl.HDC
        m_oDraw.BackStyle = BS_TRANSPARENT
        m_oDraw.BackColor = UserControl.BackColor

        Set m_ColPropertyItems = New CPropertyItems

        Me.LockWindowUpdate = False

        Dim oMainItem As CPropertyItem

        Dim oMainItemChild As CPropertyItem

        Set oMainItem = Me.PropertyItems.AddPropertyItem("Property Selection Control", "Main1")
        Set oMainItemChild = Me.PropertyItems.AddPropertyItem("Version", "Main2", oMainItem)
        oMainItemChild.Value = "v" & App.Major & "." & App.Minor
        oMainItem.Expanded = True

        Set oMainItem = Me.PropertyItems.AddPropertyItem("Author", "Main3")
        Set oMainItemChild = Me.PropertyItems.AddPropertyItem("Name", "Main4", oMainItem)
        oMainItemChild.Value = "Gary Noble"

        Set oMainItemChild = Me.PropertyItems.AddPropertyItem("Copyright", "Main5", oMainItem)
        oMainItemChild.Value = App.LegalCopyright

        Set oMainItemChild = Me.PropertyItems.AddPropertyItem("Email", "Main6", oMainItem)
        oMainItemChild.Value = "gwnoble@msn.com"

        oMainItem.Expanded = True

        LockWindowUpdate = False

        pvRefreshData

        pvDrawItems

    End If

End Sub

Friend Property Get PropList(ListPTR As Long) As CListItems

    Dim xPropList As CListItems

    CopyMemory xPropList, ListPTR, 4

    Set PropList = xPropList

    CopyMemory xPropList, 0&, 4

    ZeroMemory ObjPtr(xPropList), 4&

End Property

Private Sub pvUpDatePropertyValue(ByVal vValue As Variant)
    On Error Resume Next

    Dim SelProp As CPropertyItem
    Dim bCancel As Boolean
    Dim vNewValue As Variant
    
    Set SelProp = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)

    If Not SelProp Is Nothing Then
        If Not SelProp.ReadOnly Then
            If IsObject(vValue) Then
                Set vNewValue = vValue
            Else
                vNewValue = vValue
            End If
            
            RaiseEvent BeforeUpdate(SelProp, vNewValue, bCancel)
            
            If Not bCancel Then
            If SelProp.PropertyType = peFontEditor Then
                Set SelProp.Value = vValue
            ElseIf SelProp.PropertyType = pePictureEditor Then
                Set SelProp.Value = LoadPicture(vValue)
            Else
                SelProp.Value = vValue
            End If

            RaiseEvent PropertyChanged(SelProp)
            End If
            
        End If
        pvDrawItems
    End If

    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : SelectedProperty
' Type      : Property
' DateTime  : 26/04/2005
' Author    : Gary Noble
' Purpose   :
' Returns   : CPropertyItem
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  26/04/2005
'//---------------------------------------------------------------------------------------
Public Property Get SelectedProperty() As CPropertyItem
On Error Resume Next

    Dim SelProp As CPropertyItem

    Set SelProp = PropItem(m_DisplayPropertyItemPTRs(m_lSelectedIndex).ptr)

    If Not SelProp Is Nothing Then Set SelectedProperty = SelProp

    Set SelProp = Nothing

On Error GoTo 0

End Property
Public Property Get ShowPushedPropertyIndicators() As Boolean
Attribute ShowPushedPropertyIndicators.VB_Description = "Shows SideBar Pushed Indicators"
    ShowPushedPropertyIndicators = m_ShowPushedPropertyIndicators
End Property

Public Property Let ShowPushedPropertyIndicators(ByVal New_ShowPushedPropertyIndicators As Boolean)
    m_ShowPushedPropertyIndicators = New_ShowPushedPropertyIndicators
    PropertyChanged "ShowPushedPropertyIndicators"
    Call pvDrawItems
End Property

Public Property Get PushPropertyCaptionsVisible() As Boolean
Attribute PushPropertyCaptionsVisible.VB_Description = "Pushes The Captions Left Most When Not  In View"
    PushPropertyCaptionsVisible = m_PushPropertyCaptionsVisible
End Property

Public Property Let PushPropertyCaptionsVisible(ByVal New_PushPropertyCaptionsVisible As Boolean)
    m_PushPropertyCaptionsVisible = New_PushPropertyCaptionsVisible
    PropertyChanged "PushPropertyCaptionsVisible"
    pvDrawItems
End Property


Private Function GetNextListItem(ByVal lstPTr As Long) As String
    Dim oItems As CListItems
    Dim oItem  As CListItem
    Dim lCount As Long

    If Not SelectedProperty Is Nothing Then
        Set oItems = PropList(ListPTR:=lstPTr)

        If Not oItems Is Nothing Then
            If Not oItems.Count = 0 Then
                For Each oItem In oItems
                    lCount = lCount + 1
                    If oItem.DisplayText = SelectedProperty.Value Then
                        If lCount + 1 > oItems.Count Then
                            lCount = 1
                            GetNextListItem = oItems(lCount).DisplayText
                            Exit For
                        Else
                            GetNextListItem = oItems(lCount + 1).DisplayText
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    End If

    Set oItem = Nothing
    Set oItems = Nothing



End Function

Friend Sub FreeCapture()
    m_bDropDownPropertyShowing = False
    UserControl.SetFocus
    pvDrawItems
End Sub


Public Property Get SelectedBackColor() As OLE_COLOR
    SelectedBackColor = m_SelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal New_SelectedBackColor As OLE_COLOR)
    m_SelectedBackColor = New_SelectedBackColor
    PropertyChanged "SelectedBackColor"
End Property

Public Property Get NormalForeColor() As OLE_COLOR
    NormalForeColor = m_NormalForeColor
End Property

Public Property Let NormalForeColor(ByVal New_NormalForeColor As OLE_COLOR)
    m_NormalForeColor = New_NormalForeColor
    PropertyChanged "NormalForeColor"
End Property

Public Property Get SelectedForeColor() As OLE_COLOR
    SelectedForeColor = m_SelectedForeColor
End Property

Public Property Let SelectedForeColor(ByVal New_SelectedForeColor As OLE_COLOR)
    m_SelectedForeColor = New_SelectedForeColor
    PropertyChanged "SelectedForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    pvDrawItems
End Property

