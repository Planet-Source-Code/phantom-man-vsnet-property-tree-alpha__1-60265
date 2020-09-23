VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarValue As Variant
Private mvarDisplayText As String

Public Property Let DisplayText(ByVal vData As String)

    mvarDisplayText = vData

End Property

Public Property Get DisplayText() As String

    Let DisplayText = mvarDisplayText

End Property

Public Property Let Value(ByVal vData As Variant)

    mvarValue = vData

End Property

Public Property Set Value(ByVal vData As Variant)

    Set mvarValue = vData

End Property

Public Property Get Value() As Variant

    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else                                              'ISOBJECT(MVARVALUE) = FALSE/0
        Value = mvarValue
    End If

End Property

