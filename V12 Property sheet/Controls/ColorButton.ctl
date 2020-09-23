VERSION 5.00
Begin VB.UserControl peColorSelector 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   KeyPreview      =   -1  'True
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
   Begin VB.PictureBox picDropDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   0
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   154
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "peColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//---------------------------------------------------------------------------------------
'pePropertyEditor
'//---------------------------------------------------------------------------------------
' Module    : Dropdown Color Selector
' DateTime  : 19/04/2005
' Author    : Gary Noble   ©2005
' Purpose   : Simulates A Colour DropDown Selector
' Assumes   :
' Notes     :
' Revision  : 1.0
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------
'Event
'---------------------------------------------------------
Public Event Click()
Public Event Hide()
Public Event Show()
Private m_bShowing As Boolean

'---------------------------------------------------------
'Declares
'---------------------------------------------------------
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal HDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal HDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal HDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Const DT_CENTER As Integer = &H1
Private Const DT_SINGLELINE As Integer = &H20
Private Const DT_VCENTER As Integer = &H4

Private Const DI_NORMAL As Integer = &H3

Private Const BF_RECT As Integer = &HF

Private Const BDR_SUNKENOUTER As Integer = &H2
Private Const BDR_RAISEDINNER As Integer = &H4
Private Const BDR_RAISED As Integer = &H5

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_TOPMOST As Integer = &H8
Private Const WS_EX_TOOLWINDOW As Integer = &H80
Private Const WS_EX_WINDOWEDGE As Integer = &H100
Private Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or _
                                     WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
'---------------------------------------------------------
'Private variable
'---------------------------------------------------------
Private m_iXIndex As Integer                          '__ Selection Control
Private m_iYIndex As Integer                          '
Private m_nFillColor As Long                          '-- Button Fill Color Normal
Private m_nDarkFillColor As Long                      '-- Button Fill Color
Private m_nShadowColor As Long                        '-- Shadow
Private m_nSelectedColor As Long                      '-- Selected Color
Private m_arrColor(4, 7) As Long                      '-- Color Array
Private m_nHoverColor As Long                         '-- Hovering Color
Private m_RecLine As RECT                             '-- Custom Line Params
Private m_bInCustomColor As Boolean                   '-- Custom Color Flag
Private m_bDropDown As Boolean                        '-- Is the DropDown Shown
Private bTop   As Boolean                             '-- Dropdown Position
Dim m_nBorderColor As OLE_COLOR                       '-- border Color

'---------------------------------------------------------
'Defaults
'---------------------------------------------------------
Const m_def_BorderColor = 0
Const m_def_SelectedColor = 0

'//---------------------------------------------------------------------------------------
' Procedure : DrawRectangle
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : simply Draws A Rectangle
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Public Sub DrawRectangle(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, _
                         ByVal cx As Long, ByVal cy As Long, _
                         Optional ByVal PenColor As Long = 0, _
                         Optional ByVal BrushColor As Long = &HFFFFFF, _
                         Optional Transparent As Boolean)

    Dim hPen   As Long
    Dim hBrush As Long

    If cx < 0 Or cy < 0 Then
        Exit Sub                                      '---> Bottom
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



Private Sub picDropDown_KeyDown(KeyCode As Integer, Shift As Integer)

    Debug.Print KeyCode

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

Private Sub UserControl_Click()

    If m_bDropDown Then
        RaiseEvent Hide
        m_bDropDown = False
        Call ReleaseCapture
        picDropDown.Visible = False

        Call Redraw
    Else                                              'M_BDROPDOWN = FALSE/0
        RaiseEvent Show
        m_bDropDown = True
        Call ReleaseCapture
        SetCapture picDropDown.hwnd
        Call DropDown
        Call Redraw
    End If

End Sub

Private Sub UserControl_Initialize()

'-- Initialize color Array

    m_arrColor(0, 0) = 0
    m_arrColor(0, 1) = 13209
    m_arrColor(0, 2) = 13107
    m_arrColor(0, 3) = 13056
    m_arrColor(0, 4) = 6697728
    m_arrColor(0, 5) = 8388608
    m_arrColor(0, 6) = 10040115
    m_arrColor(0, 7) = 3355443
    m_arrColor(1, 0) = 128
    m_arrColor(1, 1) = 26367
    m_arrColor(1, 2) = 32896
    m_arrColor(1, 3) = 32768
    m_arrColor(1, 4) = 8421376
    m_arrColor(1, 5) = 16711680
    m_arrColor(1, 6) = 10053222
    m_arrColor(1, 7) = 8421504
    m_arrColor(2, 0) = 255
    m_arrColor(2, 1) = 39423
    m_arrColor(2, 2) = 52377
    m_arrColor(2, 3) = 6723891
    m_arrColor(2, 4) = 13421619
    m_arrColor(2, 5) = 16737843
    m_arrColor(2, 6) = 8388736
    m_arrColor(2, 7) = 9868950
    m_arrColor(3, 0) = 16711935
    m_arrColor(3, 1) = 52479
    m_arrColor(3, 2) = 65535
    m_arrColor(3, 3) = 65280
    m_arrColor(3, 4) = 16776960
    m_arrColor(3, 5) = 16763904
    m_arrColor(3, 6) = 6697881
    m_arrColor(3, 7) = 12632256
    m_arrColor(4, 0) = 13408767
    m_arrColor(4, 1) = 10079487
    m_arrColor(4, 2) = 10092543
    m_arrColor(4, 3) = 13434828
    m_arrColor(4, 4) = 16777164
    m_arrColor(4, 5) = 16764057
    m_arrColor(4, 6) = 16751052
    m_arrColor(4, 7) = 16777215

    m_nBorderColor = TranslateColor(vbActiveTitleBar)
    m_nFillColor = TranslateColor(vbInactiveTitleBar)
    m_nDarkFillColor = TranslateColor(vbActiveTitleBar)
    m_nShadowColor = TranslateColor(vbButtonShadow)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        UserControl_Click
    Else
        KeyCode = 0
    End If
    RaiseEvent Hide

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'-- Draw
    If Button <> 1 Then Exit Sub


    If m_bDropDown Then
        If bTop Then
            Call DrawRectangle(UserControl.HDC, 0, 1, ScaleWidth, ScaleHeight + 1, TranslateColor(BlendColor(m_nBorderColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
        Else
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(BlendColor(m_nBorderColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
        End If
    Else
        Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(BlendColor(Parent.SidebarColor, vbBlack)))
    End If

    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText
    UserControl.Refresh


End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 0 Then
        If m_bDropDown Then
            If bTop Then
                Call DrawRectangle(UserControl.HDC, 0, 1, ScaleWidth, ScaleHeight + 1, TranslateColor(BlendColor(m_nBorderColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
            Else
                Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(BlendColor(m_nBorderColor, vbBlack, 20)), TranslateColor(vbWindowBackground))
            End If
        Else
            '    Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(vbHighlight), TranslateColor(m_nDarkFillColor))
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(BlendColor(Parent.SidebarColor, vbBlack)))

        End If
        DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText
        Exit Sub
    End If

    If X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Then
        '-- MouseLeave
        Call ReleaseCapture
        Call Redraw

    Else                                              'NOT X...
        'MouseOver
        Call SetCapture(UserControl.hwnd)
        'Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 195))))
        Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(Parent.SidebarColor))
        DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText
        UserControl.Refresh

    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 1 Then Exit Sub
    Call Redraw

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_nHoverColor = -1

    Value = PropBag.ReadProperty("Value", &H0)

    m_nSelectedColor = TranslateColor(PropBag.ReadProperty("SelectedColor", m_def_SelectedColor))
    m_nBorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)

    m_nBorderColor = TranslateColor(vbActiveTitleBar)
    m_nFillColor = TranslateColor(vbInactiveTitleBar)
    m_nDarkFillColor = TranslateColor(vbActiveTitleBar)
    m_nShadowColor = TranslateColor(vbButtonShadow)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Value", Value, &H0

    Call PropBag.WriteProperty("SelectedColor", m_nSelectedColor, m_def_SelectedColor)
    Call PropBag.WriteProperty("BorderColor", m_nBorderColor, m_def_BorderColor)

End Sub

Private Sub UserControl_Resize()

    UserControl.Width = 15 * Screen.TwipsPerPixelX
    Call Redraw

End Sub

Private Sub UserControl_ExitFocus()

'Hide dropdown window

    If picDropDown.Visible Then
        picDropDown.Visible = False
        RaiseEvent Hide
        m_bDropDown = False
        Call ReleaseCapture
    End If

    Call Redraw

End Sub

Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim rc     As RECT
    Dim I      As Integer
    Dim j      As Integer

    '-- Draws The Selected Colur And The Hover Color If Neccessary
    '-- Also Selects The Hover Item Colour

    Call DrawAllColors

    If X < 0 Or X > picDropDown.ScaleWidth Or Y < 0 Or Y > picDropDown.ScaleHeight Then
        'Outside of dropdown window
        m_nHoverColor = -1
        Call ReleaseCapture

        picDropDown.Visible = False
        RaiseEvent Hide
        m_bDropDown = False
        Call Redraw

    Else                                              'NOT X...
        If Button = 1 Then

            If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 20 Then
                'User-defined color
                m_bInCustomColor = True
                Call ReleaseCapture
                ' Call SetCapture(picDropDown.Hwnd)
                Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3)
                Call DrawEdge(picDropDown.HDC, rc, BDR_SUNKENOUTER, BF_RECT)
                picDropDown.Refresh
            Else                                      'NOT X...
                m_bInCustomColor = False
                'Other colors
                Call SetCapture(picDropDown.hwnd)
                If GetColorIndex(X, Y, I, j) Then
                    m_nHoverColor = m_arrColor(j, I)
                    Call SetRect(rc, 8 + I * 18 - 3, 32 + j * 18 - 3, 8 + I * 18 + 12 + 3, 32 + j * 18 + 12 + 3)
                    If m_nSelectedColor <> m_arrColor(j, I) Then
                        DrawRectangle picDropDown.HDC, rc.Left + 1, rc.Top, (rc.Right - rc.Left), (rc.Bottom - rc.Top), TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                        DrawRectangle picDropDown.HDC, rc.Left + 5, rc.Top + 4, (rc.Right - rc.Left) - 6, (rc.Bottom - rc.Top) - 6, &H808080, TranslateColor(m_arrColor(j, I))
                    End If
                    picDropDown.Refresh
                End If
            End If
        End If

    End If

End Sub

Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim rc     As RECT
    Dim I      As Integer
    Dim j      As Integer

    '-- Draw the control
    '-- With The Selected Colours

    Call DrawAllColors

    If X < 0 Or Y < 0 Or X > picDropDown.ScaleWidth Or Y > picDropDown.Height Then
        'do nothing
        If m_bShowing Then ReleaseCapture
    Else                                              'NOT X...
        Call SetCapture(picDropDown.hwnd)

        If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 18 Then
            'User-defined color
            If Not m_bShowing Then

                Call SetRect(rc, 8, 8, 2 + 136, 18)
                If Button = 0 Then
                    DrawRectangle picDropDown.HDC, rc.Left, rc.Top, rc.Right, rc.Bottom, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                Else
                    DrawRectangle picDropDown.HDC, rc.Left, rc.Top, rc.Right, rc.Bottom, TranslateColor(vbActiveTitleBar), TranslateColor(vbActiveTitleBar)
                End If
                rc.Top = rc.Top + 4
                rc.Bottom = rc.Bottom + 4
                rc.Left = rc.Left + 8
                Call DrawText(picDropDown.HDC, "User-Defined Color", 18, rc, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE)
                picDropDown.Refresh

                If Button <> 1 Then Exit Sub

                Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138, 8 + 18 + 3)
                DrawRectangle picDropDown.HDC, rc.Left, rc.Top, rc.Right, rc.Bottom, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                UserControl.Refresh

            End If

        Else                                          'NOT X...
            'Other colors
            If GetColorIndex(X, Y, I, j) Then
                Call SetRect(rc, 8 + I * 18 - 3, 32 + j * 18 - 3, 8 + I * 18 + 12 + 3, 32 + j * 18 + 12 + 3)
                If m_iXIndex = I And m_iYIndex = j Then
                    ' Do Nothing
                Else                                  'NOT M_IXINDEX...
                    If Button = 0 Then
                        If m_nSelectedColor <> m_arrColor(j, I) Then
                            DrawRectangle picDropDown.HDC, rc.Left, rc.Top, (rc.Right - rc.Left), (rc.Bottom - rc.Top), TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                            DrawRectangle picDropDown.HDC, rc.Left + 3, rc.Top + 3, (rc.Right - rc.Left) - 6, (rc.Bottom - rc.Top) - 6, &H808080, TranslateColor(m_arrColor(j, I))
                        End If
                    ElseIf Button = 1 Then            'NOT BUTTON...

                        '-- Only Draw The Hovering Colour If the Hovering Colour <> Than The Selected Colour

                        If m_nHoverColor = -1 Then
                            If m_arrColor(j, I) <> m_nSelectedColor Then
                                DrawRectangle picDropDown.HDC, rc.Left + 1, rc.Top, (rc.Right - rc.Left), (rc.Bottom - rc.Top), TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                                DrawRectangle picDropDown.HDC, rc.Left + 3, rc.Top + 3, (rc.Right - rc.Left) - 6, (rc.Bottom - rc.Top) - 6, &H808080, m_arrColor(j, I)
                            End If
                        ElseIf m_nHoverColor = m_arrColor(j, I) Then    'NOT M_NHOVERCOLOR...
                            DrawRectangle picDropDown.HDC, rc.Left + 1, rc.Top, (rc.Right - rc.Left), (rc.Bottom - rc.Top), TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))
                            DrawRectangle picDropDown.HDC, rc.Left + 5, rc.Top + 4, (rc.Right - rc.Left) - 6, (rc.Bottom - rc.Top) - 6, &H808080, m_nHoverColor
                        End If

                    End If
                End If
                picDropDown.Refresh
            End If

        End If
    End If

End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim I As Integer, j As Integer
    Dim bSelect As Boolean
    Dim oColorDiag As IAPP_ClrDialogHandler

    If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 18 Then

        '-- Show The Custom Color Dialoge
        If m_bInCustomColor Then
            m_iXIndex = -1
            m_iYIndex = -1

            On Error Resume Next

            m_bShowing = True
            m_bDropDown = False
            Redraw

            Set oColorDiag = New IAPP_ClrDialogHandler

            With oColorDiag

                .Init , , UserControl.hwnd, eColourFlag_FullOpen

                .HWndOwner = UserControl.hwnd
                .InitialColour = m_nSelectedColor


                .Show

                If .SelectedColour > -1 Then
                    m_nSelectedColor = .SelectedColour
                    RaiseEvent Click
                End If
            End With

            m_bShowing = False

            bSelect = True
            Call ReleaseCapture
            picDropDown.Visible = False
            'RaiseEvent Hide
            m_bDropDown = False
            m_bDropDown = False
            Redraw
        End If

        Set oColorDiag = Nothing
    Else                                              'NOT X...
        'Other colors
        If GetColorIndex(X, Y, I, j) Then

            '-- If The Hovering Color Is <> than The Selected Colour

            If m_nHoverColor <> m_arrColor(j, I) Then Redraw: GoTo CleanExit
            m_iXIndex = I
            m_iYIndex = j
            m_nSelectedColor = m_arrColor(j, I)
            Debug.Print "ll: " & m_nSelectedColor
            Call ReleaseCapture
            Call Redraw

            m_bDropDown = False
            picDropDown.Visible = False
            RaiseEvent Click
            RaiseEvent Hide
            Call Redraw
            bSelect = True

        End If

    End If

CleanExit:

    Call Redraw
    If Not bSelect Then Call SetCapture(picDropDown.hwnd)
    m_bInCustomColor = False
    m_nHoverColor = -1
    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : DrawAllColors
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws All The Defined Colours And Buttons
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub DrawAllColors()

    Dim I As Integer, j As Integer
    Dim rc     As RECT

    picDropDown.Cls

    '-- Draw The Border Round The Dropdown Picture Box
    DrawRectangle picDropDown.HDC, 0, 0, picDropDown.ScaleWidth, picDropDown.ScaleHeight, BlendColor(TranslateColor(BlendColor(vbActiveTitleBar, vbBlack, 2)), TranslateColor(vbActiveTitleBar), 200)

    '-- Draw The White Line So The DropDown Looks If Is Actually Part Of The Control
    Call DrawRectangle(picDropDown.HDC, m_RecLine.Left, m_RecLine.Top, m_RecLine.Right - m_RecLine.Left, 1, vbWhite, vbWhite)

    Dim hBrush As Long

    '-- Draw The User-defined color
    If m_iXIndex >= 0 And m_iXIndex <= 7 And m_iYIndex >= 0 And m_iYIndex <= 4 Then
        Call DrawRectangle(picDropDown.HDC, 8, 8, 138, 18, &H808080)
    End If

    '-- Draw The Selected Color At The Top
    Call DrawRectangle(picDropDown.HDC, 8, 3, 138, 4, TranslateColor(vbActiveTitleBar), TranslateColor(Me.SelectedColor))

    Call SetRect(rc, 8, 8, 138 + 8, 20 + 8)
    Call DrawText(picDropDown.HDC, "User-Defined Color", 18, rc, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE)

    '-- Draw The SelectionBox Round the Selected Colour
    If m_iXIndex >= 0 And m_iXIndex <= 7 And m_iYIndex >= 0 And m_iYIndex <= 4 Then
        Call DrawSelectedColorBackground(picDropDown.HDC, 8 + m_iXIndex * 18 - 3, 32 + m_iYIndex * 18 - 3, 12 + 6, 12 + 6)

    End If

    '-- Draws All Other colors
    For I = 0 To 4
        For j = 0 To 7

            If m_nHoverColor > -1 Then
                If m_arrColor(I, j) = m_nHoverColor Then
                    If m_arrColor(I, j) <> m_nSelectedColor Then
                        DrawRectangle picDropDown.HDC, (8 + j * 18) - 2, (32 + I * 18) - 3, 18, 18, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 100)))
                        Call DrawRectangle(picDropDown.HDC, (8 + j * 18) + 2, (32 + I * 18) + 1, 12, 12, &H808080, m_arrColor(I, j))
                    Else                              'NOT M_ARRCOLOR(I,...
                        Call DrawRectangle(picDropDown.HDC, 8 + j * 18, 32 + I * 18, 12, 12, &H808080, m_arrColor(I, j))
                    End If
                Else                                  'NOT M_ARRCOLOR(I,...
                    Call DrawRectangle(picDropDown.HDC, 8 + j * 18, 32 + I * 18, 12, 12, &H808080, m_arrColor(I, j))
                End If
            Else                                      'NOT M_NHOVERCOLOR...
                Call DrawRectangle(picDropDown.HDC, 8 + j * 18, 32 + I * 18, 12, 12, &H808080, m_arrColor(I, j))
            End If

        Next j
    Next I

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : TranslateColor
' Type      : Function
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Translates a Colour To A System usable One
' Returns   : Long
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = 0
    End If

End Function

'//---------------------------------------------------------------------------------------
' Procedure : DrawSelectedColorBackground
' Type      : Sub
' DateTime  : 19/04/2005
' Author    : Gary Noble
' Purpose   : Draws The Border Round The Selected Colour
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------

Private Sub DrawSelectedColorBackground(HDC As Long, X As Long, Y As Long, cx As Long, cy As Long)

    Dim I As Long, j As Long
    Dim rc     As RECT
    Dim hBrush As Long

    DrawRectangle HDC, X, Y, cx, cy, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 50)))

End Sub

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

    m_bShowing = False

    If picDropDown.Visible Then
        picDropDown.Visible = False
        RaiseEvent Hide
        m_bDropDown = False
        Call ReleaseCapture
        Redraw
        Exit Sub                                      '---> Bottom
    End If

    '-- Get m_ixIndex and m_iyIndex
    m_iXIndex = -1
    m_iYIndex = -1
    For I = 0 To 4
        For j = 0 To 7
            If m_arrColor(I, j) = m_nSelectedColor Then
                m_iXIndex = j
                m_iYIndex = I
                Exit For                              'loop varying j
            End If
        Next j
        If m_iXIndex > -1 Or m_iYIndex > -1 Then
            Exit For                                  'loop varying i
        End If
    Next I

    Call SetParent(picDropDown.hwnd, 0)

    Call GetWindowRect(UserControl.hwnd, rc)

    If rc.Bottom < (Screen.Height - picDropDown.Height) / Screen.TwipsPerPixelY Then
        ListTop = rc.Bottom * Screen.TwipsPerPixelY
        m_RecLine.Top = 0
        m_RecLine.Bottom = 1
        m_RecLine.Left = 1
        m_RecLine.Right = ScaleWidth - 1
        bTop = False
    Else                                              'NOT RC.BOTTOM...
        ListTop = (rc.Top * Screen.TwipsPerPixelY - picDropDown.Height) + (1 * Screen.TwipsPerPixelY)
        m_RecLine.Top = picDropDown.ScaleHeight - 1
        bTop = True
    End If

    If rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        ListLeft = ((rc.Left) * Screen.TwipsPerPixelX) - (picDropDown.Width) + (15 * Screen.TwipsPerPixelX)
        m_RecLine.Left = picDropDown.ScaleWidth - 15
        m_RecLine.Right = picDropDown.ScaleWidth - 1
    Else                                              'NOT RC.RIGHT...
        ListLeft = (rc.Right) * Screen.TwipsPerPixelX - picDropDown.Width
        m_RecLine.Right = picDropDown.ScaleWidth - 1
        m_RecLine.Left = picDropDown.ScaleWidth - 14
    End If

    If ListLeft < 1 Then
        ListLeft = 16 * Screen.TwipsPerPixelX
        m_RecLine.Left = rc.Left - (15)
        m_RecLine.Right = m_RecLine.Left + 13
    ElseIf rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then    'NOT LISTLEFT...
        ListLeft = ((rc.Left) * Screen.TwipsPerPixelX) - (picDropDown.Width) + (15 * Screen.TwipsPerPixelX)
        m_RecLine.Left = picDropDown.ScaleWidth - 14
        m_RecLine.Right = picDropDown.ScaleWidth - 1
    End If

    Call SetWindowLong(picDropDown.hwnd, GWL_EXSTYLE, WS_EX_PALETTEWINDOW)    ' WS_EX_TOPMOST Or WS_EX_TOOLWINDOW)

    picDropDown.Move ListLeft, ListTop, picDropDown.Width, picDropDown.Height

    picDropDown.Visible = True

    Call DrawAllColors

    Call SetCapture(picDropDown.hwnd)

    Call Redraw

    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText

    UserControl.Refresh

    On Error GoTo 0

End Sub

Private Function GetColorIndex(ByVal X As Single, ByVal Y As Single, ByRef XIndex As Integer, ByRef YIndex As Integer) As Boolean

    Dim I      As Integer
    Dim j      As Integer

    For I = 0 To 7
        If X >= 8 + I * 18 - 3 And X <= 8 + I * 18 + 12 + 3 Then
            Exit For                                  'loop varying i
        End If
    Next

    For j = 0 To 4
        If Y >= 32 + j * 18 - 3 And Y <= 32 + j * 18 + 12 + 3 Then
            Exit For
        End If
    Next

    If I >= 8 Or j >= 5 Then
        GetColorIndex = False
    Else                                              'NOT I...
        XIndex = I
        YIndex = j
        GetColorIndex = True
    End If

End Function

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

        If bTop Then
            Call DrawRectangle(UserControl.HDC, 0, -1, ScaleWidth, IIf(m_bDropDown, ScaleHeight + 1, ScaleHeight), BlendColor(TranslateColor(vbActiveTitleBar), vbBlack, 2), TranslateColor(vbWindowBackground))
        Else                                          'BTOP = FALSE/0
            Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight + 1, TranslateColor(BlendColor(vbActiveTitleBar, vbBlack, 2)), TranslateColor(vbWindowBackground))
        End If
    Else                                              'M_BDROPDOWN = FALSE/0

        'Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(vbActiveTitleBar), TranslateColor(IIf(AppThemed, vbInactiveTitleBar, BlendColor(vbActiveTitleBar, vbWhite, 250))))
        Call DrawRectangle(UserControl.HDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(Parent.LineColor), TranslateColor(Parent.SidebarColor))

    End If

    Call DrawDropDownArrow(HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText)

    Call DrawAllColors
    UserControl.Refresh

    DrawDropDownArrow HDC, 0, 0, ScaleWidth, ScaleHeight, True, vbWindowText

    UserControl.Refresh

    On Error GoTo 0
End Sub

Public Property Get Value() As Long

    Value = m_nSelectedColor

End Property

Public Property Let Value(ByVal newValue As Long)

    m_nSelectedColor = newValue
    Call Redraw
    PropertyChanged ("Value")

End Property

Public Property Get SelectedColor() As OLE_COLOR

    SelectedColor = TranslateColor(m_nSelectedColor)

End Property

Public Property Let SelectedColor(ByVal New_SelectedColor As OLE_COLOR)

    m_nSelectedColor = TranslateColor(New_SelectedColor)
    PropertyChanged "SelectedColor"

End Property

Private Sub UserControl_InitProperties()

    m_nSelectedColor = TranslateColor(m_def_SelectedColor)
    m_nBorderColor = m_def_BorderColor

End Sub

Public Property Get BorderColor() As OLE_COLOR

    BorderColor = m_nBorderColor

End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)

    m_nBorderColor = TranslateColor(New_BorderColor)
    PropertyChanged "BorderColor"
    Redraw

End Property

Public Sub SimulateDropDown()

    UserControl_Click
    RaiseEvent Show

End Sub

