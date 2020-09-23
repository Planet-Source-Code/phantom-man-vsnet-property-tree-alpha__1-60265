VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Class1"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'pePropertyEditor
'//---------------------------------------------------------------------------------------
' Module    : ListItems
' DateTime  : 19/04/2005
' Author    : Gary Noble   ©2005 Telecom Direct Limited
' Purpose   : Holds The Collection Of Custom List Items
' Assumes   :
' Notes     :
' Revision  :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  19/04/2005
'//---------------------------------------------------------------------------------------
Option Explicit

Private mCol   As Collection

Public Function Add(Key As Variant, eType As peListDataType) As CListItems

    On Error GoTo ehHandleError_Add

    'create a new object
    Dim objNewMember As CListItems

    Set objNewMember = New CListItems

    objNewMember.ListType = eType

    If Len(Key) = 0 Then
        mCol.Add objNewMember
    Else                                              'NOT LEN(KEY)...
        mCol.Add objNewMember, Key
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

CleanExit:

    On Error GoTo 0

    Exit Function

ehHandleError_Add:

    Err.Raise Err.Number + vbObjectError, Err.Source, Err.Description

    Resume CleanExit

End Function

Public Property Get Item(vntIndexKey As Variant) As CListItems

    On Error GoTo ehHandleError_Item
    Set Item = mCol(vntIndexKey)

CleanExit:

    On Error GoTo 0

    Exit Property

ehHandleError_Item:

    Err.Raise Err.Number + vbObjectError, Err.Source, Err.Description

    Resume CleanExit

End Property

Public Property Get Count() As Long

    Count = mCol.Count

End Property

Public Sub Remove(vntIndexKey As Variant)

    On Error GoTo ehHandleError_Remove
    mCol.Remove vntIndexKey

CleanExit:

    On Error GoTo 0

    Exit Sub

ehHandleError_Remove:

    Resume CleanExit

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

