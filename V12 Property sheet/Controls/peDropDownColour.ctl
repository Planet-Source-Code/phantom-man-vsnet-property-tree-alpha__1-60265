VERSION 5.00
Begin VB.UserControl peDropDownColour 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picDropDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   0
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   154
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox CommonDialog1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin VB.Image imgDropDown 
      Height          =   45
      Left            =   480
      Picture         =   "peDropDownColour.ctx":0000
      Top             =   2520
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "peDropDownColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
'Excel Style Color Picker
'Version 1.0
'
'Copyright © 2004 by Grid2000.com. All Rights Reserved.
'---------------------------------------------------------

Option Explicit

'---------------------------------------------------------
'Event
'---------------------------------------------------------
Public Event Click()

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
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


'---------------------------------------------------------
'Private variable
'---------------------------------------------------------
Private m_iXIndex As Integer
Private m_iYIndex As Integer
Private m_nBorderColor As Long
Private m_nFillColor As Long
Private m_nDarkFillColor As Long
Private m_nShadowColor As Long
Private m_nSelectedColor As Long
Private m_arrColor(4, 7) As Long



Public Sub DrawRectangle(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
                         ByVal cx As Long, ByVal cy As Long, _
                         Optional ByVal PenColor As Long = 0, _
                         Optional ByVal BrushColor As Long = &HFFFFFF, _
                         Optional Transparent As Boolean)
    Dim hPen As Long
    Dim hBrush As Long
    
    If cx < 0 Or cy < 0 Then
        Exit Sub
    End If
    
    hPen = SelectObject(hDC, CreatePen(0, 1, PenColor))
    
    If hPen Then
        If Not Transparent Then
            hBrush = SelectObject(hDC, CreateSolidBrush(BrushColor))
        End If
        
        Call Rectangle(hDC, x, y, x + cx, y + cy)
        
        If Not Transparent Then
            Call DeleteObject(SelectObject(hDC, hBrush))
        End If
        
        Call DeleteObject(SelectObject(hDC, hPen))
    End If
End Sub


Public Function VBColorToRGB(ByVal VBColor As Long) As Long
    If OleTranslateColor(VBColor, 0, VBColorToRGB) Then
        VBColorToRGB = VBColor
    End If
End Function

Private Sub UserControl_Initialize()
    'Initialize color
    
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
    
    m_nBorderColor = RGB(8, 36, 107)
    m_nFillColor = RGB(181, 190, 214)
    m_nDarkFillColor = RGB(132, 146, 181)
    m_nShadowColor = VBColorToRGB(vbButtonShadow)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    If x > 0 And x < 23 And y > 0 And y < 23 Then
        'Draw Icon
        Call DrawRectangle(UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nDarkFillColor)
        Call DrawSelectedColor
        UserControl.Refresh
    
    ElseIf x > 23 And x < 35 And y > 0 And y < 23 Then
        'Draw dropDown arrow
        Call DrawRectangle(UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor)
        Call DrawSelectedColor
        UserControl.Refresh
        
    End If
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then Exit Sub
    
    If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
        'MouseLeave
        Call ReleaseCapture
        Call Redraw
        
    Else
        'MouseOver
        Call SetCapture(UserControl.hWnd)
        Call DrawRectangle(UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor)
        Call DrawRectangle(UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nFillColor)
        Call DrawSelectedColor
        UserControl.Refresh
        
    End If
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    If x > 0 And x < 23 And y > 0 And y < 23 Then
        'Click Icon
        RaiseEvent Click
        
    ElseIf x > 23 And x < 35 And y > 0 And y < 23 Then
        'Click dropdown arrow
        Call ReleaseCapture
        Call DropDown
        
    End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", &H0)
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Value", Value, &H0
    
End Sub


Private Sub UserControl_Resize()
    UserControl.Width = 35 * Screen.TwipsPerPixelX
    UserControl.Height = 22 * Screen.TwipsPerPixelY
    Call Redraw
End Sub


Private Sub UserControl_ExitFocus()
    'Hide dropdown window
    If picDropDown.Visible Then
        picDropDown.Visible = False
        Call ReleaseCapture
    End If
        
    Call Redraw
End Sub


Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rc As RECT
    Dim i As Integer
    Dim j As Integer
    
    Call DrawAllColors
    
    If x < 0 Or x > picDropDown.ScaleWidth Or y < 0 Or y > picDropDown.ScaleHeight Then
        'Outside of dropdown window
        Call Redraw
        Call ReleaseCapture
        picDropDown.Visible = False
        
    Else
        If Button = 1 Then
            If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 Then
                'User-defined color
                Call SetCapture(picDropDown.hWnd)
                Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3)
                Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                picDropDown.Refresh
            Else
                'Other colors
                Call SetCapture(picDropDown.hWnd)
                If GetColorIndex(x, y, i, j) Then
                    Call SetRect(rc, 8 + i * 18 - 3, 32 + j * 18 - 3, 8 + i * 18 + 12 + 3, 32 + j * 18 + 12 + 3)
                    Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                    picDropDown.Refresh
                End If
            End If
        End If
        
    End If
End Sub


Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rc As RECT
    Dim i As Integer
    Dim j As Integer
    
    Call DrawAllColors
    
    If x < 0 Or y < 0 Or x > picDropDown.ScaleWidth Or y > picDropDown.Height Then
        'do nothing
    Else
        Call SetCapture(picDropDown.hWnd)
        
        If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 Then
            'User-defined color
            Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3)
            Call DrawEdge(picDropDown.hDC, rc, BDR_RAISEDINNER, BF_RECT)
            picDropDown.Refresh
            
        Else
            'Other colors
            If GetColorIndex(x, y, i, j) Then
                Call SetRect(rc, 8 + i * 18 - 3, 32 + j * 18 - 3, 8 + i * 18 + 12 + 3, 32 + j * 18 + 12 + 3)
                If m_iXIndex = i And m_iYIndex = j Then
                    Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                Else
                    If Button = 0 Then
                        Call DrawEdge(picDropDown.hDC, rc, BDR_RAISEDINNER, BF_RECT)
                    ElseIf Button = 1 Then
                        Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                    End If
                End If
                picDropDown.Refresh
            End If
            
        End If
    End If
End Sub


Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer, j As Integer
    
    If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 Then
        'User-defined color
        m_iXIndex = -1
        m_iYIndex = -1
        
        Call ReleaseCapture
        picDropDown.Visible = False
        
        On Error Resume Next
        
        'CommonDialog1.CancelError = True
        'CommonDialog1.ShowColor
        
        If Err.Number = 0 Then
            'm_nSelectedColor = CommonDialog1.color
            RaiseEvent Click
        End If
        
        Call Redraw
        
    Else
        'Other colors
        If GetColorIndex(x, y, i, j) Then
            m_iXIndex = i
            m_iYIndex = j
            m_nSelectedColor = m_arrColor(j, i)
            
            Call ReleaseCapture
            Call Redraw
            
            picDropDown.Visible = False
            
            RaiseEvent Click
        End If
        
    End If
End Sub


Private Sub DrawAllColors()
    Dim i As Integer, j As Integer
    Dim rc As RECT
    
    picDropDown.Cls
    
    'User-defined color
    Call DrawRectangle(picDropDown.hDC, 8, 8, 138, 18, &H808080, , True)
    Call SetRect(rc, 8, 8, 138 + 8, 18 + 8)
    Call DrawText(picDropDown.hDC, "User-Defined Color", 18, rc, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE)
    
    'Selected color
    If m_iXIndex >= 0 And m_iXIndex <= 7 And m_iYIndex >= 0 And m_iYIndex <= 4 Then
        Call DrawSelectedColorBackground(picDropDown.hDC, 8 + m_iXIndex * 18 - 3, 32 + m_iYIndex * 18 - 3, 12 + 6, 12 + 6)
    End If
    
    'Other colors
    For i = 0 To 4
        For j = 0 To 7
            Call DrawRectangle(picDropDown.hDC, 8 + j * 18, 32 + i * 18, 12, 12, &H808080, m_arrColor(i, j))
        Next j
    Next i
    
    'Window border
    Call SetRect(rc, 0, 0, picDropDown.ScaleWidth, picDropDown.ScaleHeight)
    Call DrawEdge(picDropDown.hDC, rc, BDR_RAISED, BF_RECT)
End Sub


Private Sub DrawSelectedColor()
    Call DrawRectangle(UserControl.hDC, 3, 15, 16, 3, VBColorToRGB(m_nSelectedColor), VBColorToRGB(m_nSelectedColor))
    
    UserControl.PaintPicture imgDropDown.Picture, 26, 10, 5, 3, 0, 0, 5, 3, vbSrcAnd
End Sub


Private Sub DrawSelectedColorBackground(hDC As Long, x As Long, y As Long, cx As Long, cy As Long)
    Dim i As Long, j As Long
    Dim rc As RECT
    Dim hBrush As Long
    
    hBrush = CreateSolidBrush(&HFFFFFF)
    
    For i = x To x + cx - 1
        If i Mod 2 = 0 Then
            For j = y + 1 To y + cy - 1 Step 2
                Call SetRect(rc, i, j, i + 1, j + 1)
                Call FillRect(hDC, rc, hBrush)
            Next
        Else
            For j = y To y + cy - 1 Step 2
                Call SetRect(rc, i, j, i + 1, j + 1)
                Call FillRect(hDC, rc, hBrush)
            Next
        End If
    Next
    
    Call DeleteObject(hBrush)
    
    Call SetRect(rc, x, y, x + cx, y + cy)
    Call DrawEdge(hDC, rc, BDR_SUNKENOUTER, BF_RECT)
End Sub


Private Sub DropDown()
    Dim ListTop As Single, ListLeft As Single
    Dim rc As RECT
    Dim i As Integer, j As Integer
    
    'Get m_ixIndex and m_iyIndex
    m_iXIndex = -1
    m_iYIndex = -1
    For i = 0 To 4
        For j = 0 To 7
            If m_arrColor(i, j) = m_nSelectedColor Then
                m_iXIndex = j
                m_iYIndex = i
                Exit For
            End If
        Next j
        If m_iXIndex > -1 Or m_iYIndex > -1 Then
            Exit For
        End If
    Next i
    
    '
    Call GetWindowRect(UserControl.hWnd, rc)
    
    If rc.Bottom < (Screen.Height - picDropDown.Height) / Screen.TwipsPerPixelY Then
        ListTop = rc.Bottom * Screen.TwipsPerPixelY
    Else
        ListTop = rc.Top * Screen.TwipsPerPixelY - picDropDown.Height
    End If
    
    If rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        ListLeft = (rc.Left - 1) * Screen.TwipsPerPixelX
    Else
        ListLeft = (rc.Right + 1) * Screen.TwipsPerPixelX - picDropDown.Width
    End If
    
    '
    Call SetWindowLong(picDropDown.hWnd, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_TOOLWINDOW)
    Call SetParent(picDropDown.hWnd, 0)
    picDropDown.Move ListLeft, ListTop, picDropDown.Width, picDropDown.Height
    
    '
    picDropDown.Visible = True
    Call DrawAllColors
    Call SetCapture(picDropDown.hWnd)
    
    '
    Call DrawRectangle(UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor)
    Call DrawRectangle(UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor)
    Call DrawSelectedColor
    UserControl.Refresh
End Sub


Private Function GetColorIndex(ByVal x As Single, ByVal y As Single, ByRef XIndex As Integer, ByRef YIndex As Integer) As Boolean
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 7
        If x >= 8 + i * 18 - 3 And x <= 8 + i * 18 + 12 + 3 Then
            Exit For
        End If
    Next
    
    For j = 0 To 4
        If y >= 32 + j * 18 - 3 And y <= 32 + j * 18 + 12 + 3 Then
            Exit For
        End If
    Next
    
    If i >= 8 Or j >= 5 Then
        GetColorIndex = False
    Else
        XIndex = i
        YIndex = j
        GetColorIndex = True
    End If
End Function


Private Sub Redraw()
    UserControl.Cls
    Call DrawSelectedColor
    UserControl.Refresh
End Sub


'Value property
Public Property Get Value() As Long
    Value = m_nSelectedColor
End Property

Public Property Let Value(ByVal NewValue As Long)
    m_nSelectedColor = NewValue
    Call Redraw
    PropertyChanged ("Value")
End Property




