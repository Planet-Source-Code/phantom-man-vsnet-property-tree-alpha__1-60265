VERSION 5.00
Object = "*\A..\pePropertySheet.vbp"
Begin VB.Form Form1 
   Caption         =   "Property Sheet Sample (Alpha)"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin pePropertyEditor.pePropertyTree pePropertyTree1 
      Align           =   3  'Align Left
      Height          =   4005
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7064
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockWindowUpdate=   0   'False
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   5400
      ScaleHeight     =   1755
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Set Label Font"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reload"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Enable/Disable"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Push Property Captions"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Push Property Indicators"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Change Appearance"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTest 
      Caption         =   "Test Properties Sample"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRed As Integer
Dim lGreen As Integer
Dim lBlue As Integer


Private Sub Command1_Click()

    Me.pePropertyTree1.Clear

End Sub

Private Sub Command2_Click()

    Me.pePropertyTree1.LockWindowUpdate = True

    Me.pePropertyTree1.LineColor = &HFF8080
    Me.pePropertyTree1.SidebarColor = &HFFC0C0
    Me.pePropertyTree1.OpenCloseButtonColor = &HFF0000
    Me.pePropertyTree1.SelectedBackColor = &HFFC0C0
    Me.pePropertyTree1.LockWindowUpdate = False

End Sub

Private Sub Command3_Click()
    Me.pePropertyTree1.ShowPushedPropertyIndicators = Not Me.pePropertyTree1.ShowPushedPropertyIndicators
    
    
End Sub

Private Sub Command4_Click()

    Me.pePropertyTree1.PushPropertyCaptionsVisible = Not Me.pePropertyTree1.PushPropertyCaptionsVisible


End Sub

Private Sub Command5_Click()
Me.pePropertyTree1.Enabled = Not Me.pePropertyTree1.Enabled

End Sub

Private Sub Command6_Click()
    
    Form_Load
    
End Sub

Private Sub Command7_Click()


    lblTest.Font.Name = Me.pePropertyTree1.SelectedProperty.Value.Name
    lblTest.Font.Bold = Me.pePropertyTree1.SelectedProperty.Value.Bold
    lblTest.Font.Size = Me.pePropertyTree1.SelectedProperty.Value.Size
    
    '-- or
    '-- Set lblTest.Font = Me.pePropertyTree1.SelectedProperty.Value
    
End Sub

Private Sub Form_DblClick()

   
'   Set Command1.Font = Me.pePropertyTree1.SelectedProperty.Value
'    Me.Picture = pePropertyTree1.SelectedProperty.Value
    
    

End Sub

Private Sub Form_Load()

  Dim i As Long
  Dim oListBoolean      As CListItems
  Dim oListWhatever     As CListItems
  Dim pMain As CPropertyItem
  Dim pMainProp As CPropertyItem
  Dim pSubMainProp As CPropertyItem
  


    '-- Add Some Properties
    With Me.pePropertyTree1

        '-- Lock Update
        .LockWindowUpdate = True
        
        '-- Clear Window
        .Clear True
        
        '-- Add Boolean ListItems
        Set oListBoolean = .ListItems.Add("Boolean", peSimpleDropDown)
                oListBoolean.ListItemsAdd "True", "True", "True"
                oListBoolean.ListItemsAdd "False", "False", "False"
        
        '-- Add Other List Items
        Set oListWhatever = .ListItems.Add("WhatEver", peSimpleDropDown)
      
            For i = 1 To 20
                
                If i = 10 Then
                    '-- Test List Width
                    oListWhatever.ListItemsAdd i & ") Item " & i, "Item  " & i & " With A Longer Text String", i & ") Item " & i
                Else
                    oListWhatever.ListItemsAdd i & ") Item " & i, "Item " & i, i & ") Item " & i
                End If
           Next
              
        
        '-- Start
        '-- Font Property
        Set pMain = .PropertyItems.AddPropertyItem("Fonts", "FontsMain")
            pMain.Expanded = True
            
        Set pMainProp = .PropertyItems.AddPropertyItem("Font", "Font", pMain, peFontEditor)
            Set pMainProp.Value = Me.Font
            pMainProp.OverrideGroupHeaderHideItem = True
            pMainProp.Enabled = True
            pMainProp.Expanded = True
            
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Name", "Name", pMainProp)
            pSubMainProp.ReadOnly = True
            pSubMainProp.Value = Me.Font.Name
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Size", "Size", pMainProp)
            pSubMainProp.ReadOnly = True
            pSubMainProp.Value = Me.Font.Size
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Bold", "Bold", pMainProp)
            pSubMainProp.ReadOnly = True
            pSubMainProp.Value = Me.Font.Bold
        
        
        '-- Colour Selections
        Set pMain = .PropertyItems.AddPropertyItem("Colours", "Colours")
            pMain.Expanded = True
                        
            
        Set pMainProp = .PropertyItems.AddPropertyItem("Label Colour", "ActualValue", pMain, peColorEditor)
            pMainProp.Value = vbRed
            pMainProp.OverrideGroupHeaderHideItem = True
            
            
            ColorCodeToRGB pMainProp.Value, lRed, lGreen, lBlue
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("RGB Value", "RGBValue", pMainProp)
            pSubMainProp.Value = "RGB (" & lRed & "," & lGreen & "," & lBlue & ")"
            pSubMainProp.ReadOnly = True
            
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Hex Value", "Hex", pMainProp)
            pSubMainProp.Value = "Hex (" & Hex(RGB(lRed, lGreen, lBlue)) & ")"
            pSubMainProp.ReadOnly = True
            
            
        '-- List Selections
        Set pMain = .PropertyItems.AddPropertyItem("Lists", "Lists")
      
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Boolean Test", "Boolean Test", pMain)
            pSubMainProp.ReadOnly = True
      
        Set pMainProp = .PropertyItems.AddPropertyItem("Enable Label", "Label Enabled", pSubMainProp, pelistEditor, oListBoolean)
            pMainProp.Value = True
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Whatever Test", "Whatever Test", pMain)
            pSubMainProp.ReadOnly = True
      
        Set pMainProp = .PropertyItems.AddPropertyItem("Label Caption", "Value", pSubMainProp, pelistEditor, oListWhatever)
            pMainProp.Value = "Item 5"
            
            
        '-- List Selections
        Set pMain = .PropertyItems.AddPropertyItem("Dialogs", "Dialogs")
      
        Set pMainProp = .PropertyItems.AddPropertyItem("Dialog Types", "Types", pMain)
            pMainProp.ReadOnly = True
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Pictures", "Pictures", pMainProp)
            pSubMainProp.ReadOnly = True
            
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Picture", "Picture", pSubMainProp, pePictureEditor)
            pSubMainProp.DialogFilter = "Bitmaps|*.bmp|Jpeg|*.jpeg|Gif|*.gif"
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("File Select", "FileSelect", pMainProp)
            pSubMainProp.ReadOnly = True
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Open File", "Open", pSubMainProp, peFileEditor)
            pSubMainProp.DialogFilter = "Text Files|*.txt|"
            
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Folder Browser", "Browser", pMainProp)
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Browse For Folder", "BFF", pSubMainProp, peDirectoryEditor)
                  
        '-- Buttons
        Set pMain = .PropertyItems.AddPropertyItem("Buttons", "Buttons")
      
        Set pMainProp = .PropertyItems.AddPropertyItem("Button Types", "BTypes", pMain)
            pMainProp.ReadOnly = True
            
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("About", "About", pMainProp, peReadOnlyPushButton)
            pSubMainProp.Value = "(About)"
            
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Your Own", "YourOwn", pMainProp, peReadOnlyPushButton)
            pSubMainProp.Value = "(Custom)"
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Name", "Name", pMainProp, peReadOnly)
            pSubMainProp.ReadOnly = True
            pSubMainProp.Value = "Non Edit"
        
        
        Set pSubMainProp = .PropertyItems.AddPropertyItem("Spin Button", "Spin", pMainProp, peSpinButton)
            With pSubMainProp
                .Value = 20
                .SpinButtonConstraints(lIncrementValue) = 10
                .SpinButtonConstraints(Max) = 200
                .SpinButtonConstraints(Min) = -100
            End With
            
        '-- Simple Text
        Set pMain = .PropertyItems.AddPropertyItem("Simple Text", "SimpleText")
        Set pMainProp = .PropertyItems.AddPropertyItem("No Format", "No Format", pMain, peTextEditor)
        
        Set pMainProp = .PropertyItems.AddPropertyItem("Read Only", "Read Only", pMain, peTextEditor)
            pMainProp.ReadOnly = True
            
            pMainProp.ReadOnly = True
            pMainProp.Value = "Text Test"
        
        Set pMainProp = .PropertyItems.AddPropertyItem("Formatted", "Currency", pMain, peTextEditor)
            pMainProp.Value = "200.24"
            pMainProp.DisplayFormat = "£###0.00"
        
        Set pMainProp = .PropertyItems.AddPropertyItem("Date", "Date", pMain, peTextEditor)
            pMainProp.Value = Date
            pMainProp.DisplayFormat = "dd mmmm yyyy"
        
            
        .LockWindowUpdate = False

    End With 'ME.PEPROPERTYTREE1


End Sub



Private Sub pePropertyTree1_BeforeUpdate(oPropItem As pePropertyEditor.CPropertyItem, NewValue As Variant, bCancel As Boolean)
       
       '-- Validate
       If oPropItem.Key = "Date" Then
            If Not IsDate(NewValue) Then
                MsgBox "Value Must Be Of A Date Value"
                bCancel = True
            End If
        ElseIf oPropItem.Key = "Currency" Then
            If Not IsNumeric(NewValue) Then
                MsgBox "Value Must Be Of A Currency Value"
                bCancel = True
            End If
            
        End If
        
End Sub

Private Sub pePropertyTree1_PropertyChanged(oPropItem As pePropertyEditor.CPropertyItem)
    
    '-- Here You would Save Your Property Data
    Debug.Print "Property Changed: " & oPropItem.Value
    
    Dim p As CPropertyItem
    
    If oPropItem.PropertyType = peFontEditor Then
            Set p = oPropItem.ChildProperties.Items(1): p.Value = oPropItem.Value.Name
            Set p = oPropItem.ChildProperties.Items(2): p.Value = oPropItem.Value.Size
            Set p = oPropItem.ChildProperties.Items(3): p.Value = oPropItem.Value.Bold
                
    ElseIf oPropItem.PropertyType = peColorEditor Then
            
            ColorCodeToRGB oPropItem.Value, lRed, lGreen, lBlue
            
            Set p = oPropItem.ChildProperties.Items(1): p.Value = "RGB (" & lRed & "," & lGreen & "," & lBlue & ")"
            Set p = oPropItem.ChildProperties.Items(2): p.Value = "Hex (" & Hex(RGB(lRed, lGreen, lBlue)) & ")"
            lblTest.ForeColor = oPropItem.Value
    
    ElseIf oPropItem.PropertyType = pelistEditor Then
        Select Case oPropItem.Key
            Case "Label Enabled"
                lblTest.Enabled = oPropItem.Value
            Case Else
                lblTest.Caption = oPropItem.Value
        End Select
    ElseIf oPropItem.PropertyType = pePictureEditor Then
        
        Me.Picture1.Picture = oPropItem.Value
    
    End If
    
End Sub

Private Sub pePropertyTree1_PropertySelected(oPropItem As pePropertyEditor.CPropertyItem)

    If oPropItem.Key = "Font" Then
        Command7.Enabled = True
    Else
        Command7.Enabled = False
    End If

End Sub

Private Sub pePropertyTree1_PushButtonclick(oPropItem As pePropertyEditor.CPropertyItem)
    
    If oPropItem.Key = "About" Then
        MsgBox vbCrLf & "Property Tree V0.1" & vbCrLf & _
                "Copyright 2005 Gary Noble" & vbCrLf, vbInformation, "About"
    Else
        
        MsgBox "Push Button Clicked For: " & oPropItem.Caption
    
    End If
    
End Sub


Private Sub ColorCodeToRGB(lColorCode As Long, iRed As Integer, iGreen As Integer, iBlue As Integer)
    
    Dim lColor As Long
    lColor = lColorCode
    iRed = lColor Mod &H100
    lColor = lColor \ &H100
    iGreen = lColor Mod &H100
    lColor = lColor \ &H100
    iBlue = lColor Mod &H100
    
End Sub
