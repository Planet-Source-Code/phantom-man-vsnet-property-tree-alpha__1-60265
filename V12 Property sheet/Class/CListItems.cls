VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CListItem"
Attribute VB_Ext_KEY = "Member0" ,"CListItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mCol   As Collection
'local variable(s) to hold property value(s)
Private mvarListType As peListDataType                'local copy

Friend Property Let ListType(ByVal vData As peListDataType)

    Let mvarListType = vData

End Property

Public Property Get ListType() As peListDataType

    Let ListType = mvarListType

End Property

Public Function ListItemsAdd(Value As Variant, DisplayText As String, Optional sKey As String) As CListItem

    On Error GoTo ehHandleError_ListItemsAdd

    'create a new object
    Dim objNewMember As CListItem

    Set objNewMember = New CListItem

    'set the properties passed into the method
    If IsObject(Value) Then
        Set objNewMember.Value = Value
    Else                                              'ISOBJECT(VALUE) = FALSE/0
        objNewMember.Value = Value
    End If
    If IsObject(DisplayText) Then
        Let objNewMember.DisplayText = DisplayText
    Else                                              'ISOBJECT(DISPLAYTEXT) = FALSE/0
        objNewMember.DisplayText = DisplayText
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else                                              'NOT LEN(SKEY)...
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set ListItemsAdd = objNewMember
    Set objNewMember = Nothing

CleanExit:

    On Error GoTo 0

    Exit Function

ehHandleError_ListItemsAdd:

    Err.Raise Err.Number + vbObjectError, Err.Source, Err.Description

    Resume CleanExit

End Function

Public Property Get ListItem(vntIndexKey As Variant) As CListItem
Attribute ListItem.VB_UserMemId = 0

    On Error GoTo ehHandleError_ListItem
    Set ListItem = mCol(vntIndexKey)

CleanExit:

    On Error GoTo 0

    Exit Property

ehHandleError_ListItem:

    Resume CleanExit

End Property

Public Property Get Count() As Long

    Count = mCol.Count

End Property

Public Sub Remove(vntIndexKey As Variant)

    mCol.Remove vntIndexKey

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mCol.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()

    Set mCol = Nothing

End Sub

