VERSION 5.00
Begin VB.UserControl peListPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   KeyPreview      =   -1  'True
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   118
   Begin VB.PictureBox picDropDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   120
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   975
      Begin pePropertyEditor.CoolList CoolList1 
         Height          =   840
         Left            =   -360
         TabIndex        =   1
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1482
         Appearance      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackSelectedG2  =   255
         BoxBorder       =   -2147483646
         BoxOffset       =   0
         Focus           =   0   'False
         HoverSelection  =   -1  'True
         ItemHeight      =   14
         SelectModeStyle =   4
      End
   End
End
Attribute VB_Name = "peListPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iHook

Private m_bDropDownCap As Boolean
Private m_lMaxWidth As Long
Private m_oItems As CListItems
Private m_oItem As CListItem

Public Event ItemSelected(oText As String)
Public Event Hide()
Public Event Show()

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal HDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal HDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal HDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Private Const DI_NORMAL = &H3

Private Const BF_RECT = &HF

Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISED = &H5

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_TOPMOST = &H8
Private Const WS_EX_TOOLWINDOW = &H80

Private m_HookEx As cHook
Private m_bDropDown As Boolean
Private m_RecLine As RECT                             '-- Custom Line Params
Private m_bTop As Boolean                             '-- Dropdown Position
Private m_BillKillWindow As Boolean

Const m_def_BorderColor = vbActiveTitleBar
Const m_def_DarkFillColor = vbHighlight
Const m_def_LightFillcolor = vbInactiveTitleBar

Dim m_BorderColor As OLE_COLOR
Dim m_DarkFillColor As OLE_COLOR
Dim m_LightFillcolor As OLE_COLOR
Dim m_bKill    As Boolean
'Default Property Values:
Const m_def_DropDownHeight = 200
Const m_def_ButtonHeight = 20
'Property Variables:
Dim m_DropDownHeight As Long
Dim m_ButtonHeight As Long






Private Sub DrawRectangle(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, _
                          ByVal cx As Long, ByVal cy As Long, _
                          Optional ByVal PenColor As Long = 0, _
                          Optional ByVal BrushColor As Long = &HFFFFFF, _
                          Optional Transparent As Boolean)
    Dim hPen   As Long
    Dim hBrush As Long

    If cx < 0 Or cy < 0 Then
        Exit Sub
    End If

    hPen = SelectObject(HDC, CreatePen(0, 1, PenColor))

    If hPen Then
        If Not Transparent Then
            hBrush = SelectObject(HDC, CreateSolidBrush(BrushColor))
        End If

        Call Rectangle(HDC, X, Y, X + cx, Y + cy)

        If Not Transparent Then
            Call DeleteObject(SelectObject(HDC, hBrush))
        End If

        Call DeleteObject(SelectObject(HDC, hPen))
    End If
End Sub


Private Function VBColorToRGB(ByVal VBColor As Long) As Long
    If OleTranslateColor(VBColor, 0, VBColorToRGB) Then
        VBColorToRGB = VBColor
    End If
End Function

'//---------------------------------------------------------------------------------------
' Procedure : DropDown
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Shows The Dropdown Color Picker
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Private Sub DropDown()
    On Error Resume Next

    Dim ListTop As Single, ListLeft As Single
    Dim rc     As RECT
    Dim I As Integer, j As Integer


    Dim bFound As Boolean
    RaiseEvent Show

    picDropDown.Visible = False
    Call SetParent(picDropDown.hwnd, 0)

    Call GetWindowRect(UserControl.hwnd, rc)

    picDropDown.Move picDropDown.Left, picDropDown.Top, m_lMaxWidth, (TextHeight("Q,") * Screen.TwipsPerPixelY) * (IIf(CoolList1.ListCount + 1 > 15, 16, CoolList1.ListCount + 1)) - ((TextHeight("Q,") * Screen.TwipsPerPixelY) \ 2)

    UserControl.Refresh

    If rc.Bottom < (Screen.Height - picDropDown.Height) / Screen.TwipsPerPixelY Then
        ListTop = (rc.Bottom * Screen.TwipsPerPixelY) + 1
        m_RecLine.Top = 0
        m_RecLine.Bottom = 1
        m_RecLine.Left = 200
        m_RecLine.Right = ScaleWidth - 1
        m_bTop = False
    Else
        ListTop = (rc.Top * Screen.TwipsPerPixelY - picDropDown.Height) + (1 * Screen.TwipsPerPixelY)
        m_RecLine.Top = picDropDown.ScaleHeight - 1
        m_bTop = True
        If rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
            ListLeft = ((rc.Left) * Screen.TwipsPerPixelX)    ';- (picDropDown.Width) + (15 * Screen.TwipsPerPixelX)
        End If

    End If

    If rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        Debug.Print "2"
        ListLeft = (picDropDown.Width) + (UserControl.ScaleWidth * Screen.TwipsPerPixelX)
        m_RecLine.Left = picDropDown.ScaleWidth - UserControl.ScaleWidth
        m_RecLine.Right = picDropDown.ScaleWidth - 1
    Else
        Debug.Print "2s"
        ListLeft = (rc.Right) * Screen.TwipsPerPixelX - picDropDown.Width
        m_RecLine.Right = picDropDown.ScaleWidth - 1
        m_RecLine.Left = picDropDown.ScaleWidth - 16
    End If

    If rc.Right - ((picDropDown.Width) / Screen.TwipsPerPixelX) < 0 Then
        ListLeft = IIf(rc.Right - (rc.Right - rc.Left) - ((picDropDown.Width) / Screen.TwipsPerPixelX) > 0, rc.Left, 0)
        m_RecLine.Left = rc.Left                      '- (15)
        m_RecLine.Right = m_RecLine.Left + 15
    ElseIf rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        ListLeft = ((rc.Left) * Screen.TwipsPerPixelX) - (picDropDown.Width) + (UserControl.ScaleWidth * Screen.TwipsPerPixelX)
        m_RecLine.Left = picDropDown.ScaleWidth - UserControl.ScaleWidth
        m_RecLine.Right = picDropDown.ScaleWidth - 1
    End If


    Call SetWindowLong(picDropDown.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_TOOLWINDOW)

    picDropDown.Move ListLeft, ListTop, picDropDown.Width, picDropDown.Height
    CoolList1.Move 1, 3, (picDropDown.Width \ Screen.TwipsPerPixelX) - 2, (picDropDown.Height \ Screen.TwipsPerPixelY) - 2
    CoolList1.Visible = True

    picDropDown.Visible = True

    '-- Draw The Border Round The Dropdown Picture Box
    DrawRectangle picDropDown.HDC, 0, 0, picDropDown.ScaleWidth, picDropDown.ScaleHeight, TranslateColor(BlendColor(Parent.LineColor, vbBlack, 20)), TranslateColor(vbWindowBackground)

    '-- Draw The White Line So The DropDown Looks If Is Actually Part Of The Control
    Call DrawRectangle(picDropDown.HDC, m_RecLine.Left + 1, m_RecLine.Top, (m_RecLine.Right - m_RecLine.Left) - 1, 1, vbWhite, vbWhite)

    Call Redraw

    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText

    UserControl.Refresh

    m_HookEx.Hook Me, WH_MOUSE_LL, False

    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : Redraw
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Redraws The control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Public Sub Redraw()
    On Error Resume Next

    UserControl.Cls

    If m_bDropDown Then
        If m_bTop Then
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(BlendColor(Parent.LineColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
        Else
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight + 1, TranslateColor(BlendColor(Parent.LineColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
        End If

        Call DrawRectangle(picDropDown.HDC, m_RecLine.Left + 1, m_RecLine.Top, (m_RecLine.Right - m_RecLine.Left) - 1, 1, vbWhite, vbWhite)
    Else
        Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(Parent.SidebarColor))
    End If


    Call DrawDropDownArrow(HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText)

    UserControl.Refresh

    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText

    UserControl.Refresh

    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : DrawDropDownArrow
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draw A Dropdown Arrow
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Public Sub DrawDropDownArrow( _
       ByVal lhdc As Long, _
       ByVal lLeft As Long, _
       ByVal lTop As Long, _
       ByVal lWidth As Long, _
       ByVal lHeight As Long, _
       ByVal bEnabled As Boolean, _
       ByVal oColor As OLE_COLOR, _
       Optional ByVal bUp As Boolean = False)

    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim Color  As Long

    Color = TranslateColor(oColor)


    lCentreX = lLeft + lWidth \ 2
    lCentreY = lTop + lHeight \ 2

    If Not bUp Then
        SetPixel lhdc, lCentreX - 2, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY - 1, Color
        SetPixel lhdc, lCentreX, lCentreY - 1, Color
        SetPixel lhdc, lCentreX + 1, lCentreY - 1, Color
        SetPixel lhdc, lCentreX + 2, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY, Color
        SetPixel lhdc, lCentreX + 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color
    Else

        SetPixel lhdc, lCentreX, lCentreY - 1, Color
        SetPixel lhdc, lCentreX - 1, lCentreY, Color
        SetPixel lhdc, lCentreX + 1, lCentreY, Color
        SetPixel lhdc, lCentreX, lCentreY, Color
        SetPixel lhdc, lCentreX - 1, lCentreY + 1, Color
        SetPixel lhdc, lCentreX + 1, lCentreY + 1, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color
        SetPixel lhdc, lCentreX - 2, lCentreY + 1, Color
        SetPixel lhdc, lCentreX + 2, lCentreY + 1, Color
        SetPixel lhdc, lCentreX, lCentreY + 1, Color


    End If
End Sub




Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long
' Convert Automation color to Windows color
    Call OleTranslateColor(oClr, hPal, TranslateColor)

End Function



Private Sub CoolList1_Click()

    RaiseEvent ItemSelected(CoolList1.ItemText(CoolList1.ListIndex))
    pvKillWindow

End Sub

Private Sub iHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, nCode As WinSubHook2.eHookCode, wParam As Long, lParam As Long)


    Dim rcWindow As RECT
    Dim dat    As tMSLLHOOKSTRUCT                     'Low-level mouse data

    If nCode = HC_ACTION Then
        If bBefore Then
            'After

            dat = m_HookEx.xMSLLHOOKSTRUCT(lParam)

            If WindowFromPoint(dat.pt.X, dat.pt.Y) = UserControl.hwnd Then
                If wParam = 513 Or wParam = 516 Or wParam = 519 Then
                    m_BillKillWindow = False
                    Redraw
                    '-- Window Just Open - Do Nothing
                    Exit Sub                          '---> Bottom
                End If
            End If


            GetWindowRect picDropDown.hwnd, rcWindow

            If wParam = 513 Or wParam = 516 Or wParam = 519 Then
                If PtInRect(rcWindow, dat.pt.X, dat.pt.Y) Then
                    m_BillKillWindow = False
                Else                                  'NOT PTINRECT(RCWINDOW,...
                    m_BillKillWindow = True

                    If Not m_bDropDown Then

                        pvKillWindow
                    Else                              'NOT NOT...
                        If picDropDown.Visible Then
                            pvKillWindow
                        End If
                    End If
                End If
            ElseIf wParam = 514 Or wParam = 517 Or wParam = 520 Then    'NOT WPARAM...
                If PtInRect(rcWindow, dat.pt.X, dat.pt.Y) Then
                    m_BillKillWindow = False
                Else                                  'NOT PTINRECT(RCWINDOW,...
                    If m_BillKillWindow Then
                        pvKillWindow
                    Else                              'M_BILLKILLWINDOW = FALSE/0
                        Debug.Print "Keep Live"
                    End If
                End If
            End If

        End If
    End If

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbactivetitlebar
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Parent.LineColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Parent.LineColor = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbhighlight
Public Property Get DarkFillColor() As OLE_COLOR
    DarkFillColor = m_DarkFillColor
End Property

Public Property Let DarkFillColor(ByVal New_DarkFillColor As OLE_COLOR)
    m_DarkFillColor = New_DarkFillColor
    PropertyChanged "DarkFillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbinactivetitlebar
Public Property Get LightFillcolor() As OLE_COLOR
    LightFillcolor = Parent.SidebarColor
End Property

Public Property Let LightFillcolor(ByVal New_LightFillcolor As OLE_COLOR)
    Parent.SidebarColor = New_LightFillcolor
    PropertyChanged "LightFillcolor"
End Property

Private Sub UserControl_Click()

    If picDropDown.Visible Then
        pvKillWindow
        m_bDropDown = False
        picDropDown.Visible = False
        Call Redraw
    Else
        If Not m_bDropDownCap = True Then
            m_bDropDownCap = True
            m_bDropDown = True
            Call DropDown
            Call Redraw
            RaiseEvent Show
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
    m_bDropDownCap = False
    UserControl_Click
End Sub

Private Sub UserControl_Initialize()
    Set m_HookEx = New cHook
    m_bDropDown = False
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Parent.LineColor = m_def_BorderColor
    m_DarkFillColor = m_def_DarkFillColor
    Parent.SidebarColor = m_def_LightFillcolor
    m_ButtonHeight = m_def_ButtonHeight
    m_DropDownHeight = m_def_DropDownHeight
End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        pvKillWindow
    Else
        CoolList1.SimulateKeyPress KeyCode, Shift
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_LightFillcolor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_DarkFillColor = PropBag.ReadProperty("DarkFillColor", m_def_DarkFillColor)
    m_LightFillcolor = PropBag.ReadProperty("LightFillcolor", m_def_LightFillcolor)
    Redraw
    m_ButtonHeight = PropBag.ReadProperty("ButtonHeight", m_def_ButtonHeight)
    m_DropDownHeight = PropBag.ReadProperty("DropDownHeight", m_def_DropDownHeight)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    If Height > m_ButtonHeight * Screen.TwipsPerPixelY Then Height = m_ButtonHeight * Screen.TwipsPerPixelY

    picDropDown.Move 0, m_ButtonHeight + 1, ScaleWidth, m_DropDownHeight
    Redraw

    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    m_HookEx.UnHook
    Set m_HookEx = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("DarkFillColor", m_DarkFillColor, m_def_DarkFillColor)
    Call PropBag.WriteProperty("LightFillcolor", m_LightFillcolor, m_def_LightFillcolor)

    Call PropBag.WriteProperty("ButtonHeight", m_ButtonHeight, m_def_ButtonHeight)
    Call PropBag.WriteProperty("DropDownHeight", m_DropDownHeight, m_def_DropDownHeight)
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : UserControl_Click
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Shows Or Hides the Dropdown
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Private Sub UserControl_Clicka()

    If m_bDropDown Then
        m_bDropDown = False
        picDropDown.Visible = False
        ReleaseCapture
        Call Redraw
    Else
        m_bDropDown = True
        m_HookEx.Hook Me, WH_MOUSE_LL, False
        Call DropDown
        Call Redraw
    End If

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)



    If m_bDropDown Then m_bDropDownCap = True: pvKillWindow: Exit Sub
    m_bDropDownCap = False
    '-- Draw
    If Button <> 1 Then Exit Sub

    If m_bDropDown Then
        If m_bTop Then
            Call DrawRectangle(UserControl.HDC, 0, 1, ScaleWidth, ScaleHeight + 1, TranslateColor(Parent.LineColor), TranslateColor(vbWindowBackground))
        Else
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(vbWindowBackground))
        End If
    Else
        Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(BlendColor(Parent.SidebarColor, vbBlack)))
    End If
    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText
    UserControl.Refresh


End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then Exit Sub

    If X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Then
        '-- MouseLeave
        Call ReleaseCapture
        Call Redraw

    Else
        'MouseOver
        DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText
        UserControl.Refresh
        Redraw
    End If
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_bDropDown Then m_bDropDownCap = False
    If Button <> 1 Then Exit Sub
    Call Redraw

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,20
Public Property Get ButtonHeight() As Long
    ButtonHeight = m_ButtonHeight
End Property

Public Property Let ButtonHeight(ByVal New_ButtonHeight As Long)
    m_ButtonHeight = New_ButtonHeight
    PropertyChanged "ButtonHeight"
End Property

Public Property Get DropDownHeight() As Long
    DropDownHeight = m_DropDownHeight
End Property

Public Property Let DropDownHeight(ByVal New_DropDownHeight As Long)
    m_DropDownHeight = New_DropDownHeight
    PropertyChanged "DropDownHeight"
End Property

Private Sub pvKillWindow()

    RaiseEvent Hide
    'm_bDropDownCap = False
    m_bDropDown = False
    picDropDown.Visible = False
    CoolList1.Visible = False
    m_HookEx.UnHook
    DoEvents
    Debug.Print "Kill"
    Call Redraw

End Sub


Public Property Let ListDataPTR(LPTR As Long)

    Dim lListHeight As Long
    Dim bFound As Boolean

    Set CoolList1.Font = Parent.Font
    m_lMaxWidth = 65

    If Not CoolList1.Visible Then

        Set m_oItems = PropList(LPTR)
        Dim I  As Long

        If Not m_oItems Is Nothing Then
            CoolList1.Clear

            For Each m_oItem In m_oItems
                I = I + 1
                CoolList1.AddItem m_oItem.DisplayText
                If m_oItem.DisplayText = Parent.SelectedProperty.Value Then CoolList1.ListIndex = I - 1: bFound = True
                If TextWidth(m_oItem.DisplayText) * Screen.TwipsPerPixelX > m_lMaxWidth Then
                    m_lMaxWidth = (TextWidth(m_oItem.DisplayText) * Screen.TwipsPerPixelX) + (35 * Screen.TwipsPerPixelX)
                End If
            Next
            picDropDown.Height = (TextHeight("Q,") * Screen.TwipsPerPixelY) * (I + 1) - 10
            picDropDown.Width = m_lMaxWidth
            If Not bFound Then CoolList1.ListIndex = 0

        Else
            CoolList1.Clear
        End If

    End If


End Property

Friend Property Get PropList(ptr As Long) As CListItems

    Dim xPropList As CListItems

    CopyMemory xPropList, ptr, 4

    Set PropList = xPropList

    CopyMemory xPropList, 0&, 4

    ZeroMemory ObjPtr(xPropList), 4&

End Property

