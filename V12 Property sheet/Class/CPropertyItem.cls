VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPropertyItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'pePropertyTreeX
'//---------------------------------------------------------------------------------------
' Module    : CPropertyItem
' DateTime  : 04/10/2004 15:51
' Author    : Gary Noble
' Purpose   : Item Object
' Assumes   : CPropertyItems.cls
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private m_oSpinButtonConstraints As tSpinButtonConstraints

Private mvarKey As String
Private mvarCaption As String
Private mvarParent As CPropertyItem
Private mvarChildCount As Long
Private mvarChildProperties As CPropertyItems         ' Used For Adding Child Properties
Private mvarExpanded As Boolean
Private mvarLevel As Long
Private mvarRectData As RECT
Private mvarBold As Boolean
Private mvarID As Long
Private mvarItemFont As StdFont
Private mvarCheVronRect As RECT
Private mvarForeColor As OLE_COLOR
Private mvarValue As Variant
Private mvarValueColor As OLE_COLOR
Private mvarValueBold As Boolean
Private mvarParentPTR As Long
Private mvarIconIndex As Long
Private mvarPropertyType As peEditorType
Private mvarListPTR As Long
Private mvarTextFormat As String
Private mvarReadOnly As Boolean
Private mvarBackColor As Long
Private mvarHelpString As String
Private mvarFont As StdFont
Private mvarStartY As Long
Private mvarItemHeight As Long
Private mvarEnabled As Boolean
Private m_lOwner As Long
Private m_hWnd As Long
Private m_sComboKey As String
Private m_lComboID As Long
Private m_ctlCombo As ComboBox
Private mvarOverrideGroupHeaderHideItem As Boolean 'local copy
Private mvarReAdjustedOpenCloseButton As Boolean 'local copy
Private mvarDialogFilter As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDisplayFormat As String 'local copy
Public Property Let DisplayFormat(ByVal vData As String)
    mvarDisplayFormat = vData
End Property


Public Property Get DisplayFormat() As String
    DisplayFormat = mvarDisplayFormat
End Property


Public Property Let DialogFilter(ByVal sFilter As String)
    mvarDialogFilter = sFilter
End Property

Public Property Get DialogFilter() As String
    DialogFilter = mvarDialogFilter
End Property

Friend Property Let ReAdjustedOpenCloseButton(ByVal vData As Boolean)
    mvarReAdjustedOpenCloseButton = vData
End Property


Friend Property Get ReAdjustedOpenCloseButton() As Boolean
    ReAdjustedOpenCloseButton = mvarReAdjustedOpenCloseButton
End Property

Public Property Let OverrideGroupHeaderHideItem(ByVal vData As Boolean)
    mvarOverrideGroupHeaderHideItem = vData
End Property


Public Property Get OverrideGroupHeaderHideItem() As Boolean
    OverrideGroupHeaderHideItem = mvarOverrideGroupHeaderHideItem
End Property




Public Property Let Enabled(ByVal vData As Boolean)

    mvarEnabled = vData

End Property

Public Property Get Enabled() As Boolean

    Enabled = mvarEnabled

End Property

Friend Sub Init(ByVal LPTR As Long, _
                ByVal lhWnd As Long)

    m_lOwner = LPTR
    m_hWnd = lhWnd

End Sub

Private Property Get PropertyTreeObject() As pePropertyTree

    Dim oT     As Object

    If IsWindow(m_hWnd) Then
        If Not (m_lOwner = 0) Then
            CopyMemory oT, m_lOwner, 4&
            Set PropertyTreeObject = oT
            CopyMemory oT, 0&, 4&
        End If
    End If

End Property

Friend Property Let StartY(lStart As Long)

    mvarStartY = lStart

End Property

Friend Property Get StartY() As Long

    StartY = mvarStartY

End Property

Friend Property Get ItemHeight() As Long

    ItemHeight = mvarItemHeight

End Property

Public Property Let HelpString(ByVal vData As String)

    mvarHelpString = vData

End Property

Public Property Get HelpString() As String

    HelpString = mvarHelpString

End Property

Public Property Let BackColor(ByVal vData As Long)

    mvarBackColor = vData

End Property

Public Property Get BackColor() As Long

    BackColor = mvarBackColor

End Property

Public Property Let ReadOnly(ByVal vData As Boolean)

    mvarReadOnly = vData

End Property

Public Property Get ReadOnly() As Boolean

    ReadOnly = mvarReadOnly

End Property

Public Property Let TextFormat(ByVal vData As String)

    mvarTextFormat = vData

End Property

Public Property Get TextFormat() As String

    TextFormat = mvarTextFormat

End Property

Public Property Let SetListData(ListData As CListItems)

    ListPTR = ObjPtr(ListData)

End Property

Friend Property Let ListPTR(ByVal vData As Long)

    mvarListPTR = vData

End Property

Friend Property Get ListPTR() As Long

    ListPTR = mvarListPTR

End Property

Public Property Let PropertyType(ByVal vData As peEditorType)

    Let mvarPropertyType = vData

    Set m_ctlCombo = Nothing

End Property

Friend Property Get ComboID() As Long

    ComboID = m_lComboID

End Property

Friend Property Get ComboKey() As String

    ComboKey = m_sComboKey

End Property

Public Property Get PropertyType() As peEditorType

    Let PropertyType = mvarPropertyType

End Property

Friend Property Let ParentPTR(ByVal vData As Long)

    mvarParentPTR = vData

End Property

Friend Property Get ParentPTR() As Long

    ParentPTR = mvarParentPTR

End Property


Public Property Set Value(ByVal vData As Variant)

    Set mvarValue = vData

End Property

Public Property Let Value(ByVal vData As Variant)
On Error Resume Next

    mvarValue = vData

On Error GoTo 0
End Property


Public Property Get Value() As Variant
    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else
        Value = mvarValue
    End If
End Property



Friend Property Let id(ByVal vData As Long)

    mvarID = vData

End Property

Friend Property Get id() As Long

    id = mvarID

End Property

Friend Sub SetRect(X As Long, Y As Long, Right As Long, Bottom As Long)

    mvarRectData.Top = Y
    mvarRectData.Left = X
    mvarRectData.Right = Right
    mvarRectData.Bottom = Bottom

End Sub

Friend Property Get RectData() As RECT

    Let RectData = mvarRectData

End Property

Friend Property Let Level(ByVal vData As Long)

    mvarLevel = vData

End Property

Friend Property Get Level() As Long

    Level = mvarLevel

End Property

Public Property Let Expanded(ByVal vData As Boolean)

    mvarExpanded = vData

End Property

Public Property Get Expanded() As Boolean

    Let Expanded = mvarExpanded

End Property

Public Property Get ChildProperties() As CPropertyItems

    If mvarChildProperties Is Nothing Then
        Set mvarChildProperties = New CPropertyItems

    End If

    Set ChildProperties = mvarChildProperties
    ChildProperties.Init m_lOwner, m_hWnd

End Property

Public Property Let ChildCount(ByVal vData As Long)

    mvarChildCount = vData

End Property

Public Property Get ChildCount() As Long

    ChildCount = mvarChildCount

End Property

Public Property Set Parent(ByVal vData As CPropertyItem)

    Set mvarParent = vData

End Property

Public Property Get Parent() As CPropertyItem

    Set Parent = mvarParent

End Property

Public Property Let Caption(ByVal vData As String)

    mvarCaption = vData

End Property

Public Property Get Caption() As String

    Caption = mvarCaption

End Property

Friend Property Let Key(ByVal vData As String)

    mvarKey = vData

End Property

Public Property Get Key() As String

    Let Key = mvarKey

End Property

Private Sub Class_Initialize()

    mvarEnabled = True
    Set mvarChildProperties = New CPropertyItems

End Sub

Private Sub Class_Terminate()

    Set mvarChildProperties = Nothing

End Sub

'-- Spin Button Values And Params

Public Property Let SpinButtonConstraints(oType As eSpinButtonConstraints, ByVal Val As Long)

    If oType = eSpinButtonConstraints.Min Then
        m_oSpinButtonConstraints.lMin = Val
    ElseIf oType = eSpinButtonConstraints.Max Then    'NOT OTYPE...
        m_oSpinButtonConstraints.lMax = Val
    ElseIf oType = eSpinButtonConstraints.lIncrementValue Then    'NOT OTYPE...
        m_oSpinButtonConstraints.lIncrementValue = Val
    End If

End Property

Public Property Get SpinButtonConstraints(oType As eSpinButtonConstraints) As Long

    If oType = eSpinButtonConstraints.Min Then
        SpinButtonConstraints = m_oSpinButtonConstraints.lMin
    ElseIf oType = eSpinButtonConstraints.Max Then    'NOT OTYPE...
        SpinButtonConstraints = m_oSpinButtonConstraints.lMax
    ElseIf oType = eSpinButtonConstraints.lIncrementValue Then    'NOT OTYPE...
        SpinButtonConstraints = m_oSpinButtonConstraints.lIncrementValue
    End If

End Property

