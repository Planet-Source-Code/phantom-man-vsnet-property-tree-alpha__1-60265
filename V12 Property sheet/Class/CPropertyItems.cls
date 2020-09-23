VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPropertyItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mvarItems As Collection
Private mvarTotalItemCount As Long
Private m_lOwner As Long
Private m_hWnd As Long
Private m_DefFont As StdFont

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private m_Tree As pePropertyTree

Friend Sub Init(ByVal LPTR As Long, _
                ByVal lhWnd As Long)

    m_lOwner = LPTR
    m_hWnd = lhWnd
    Set m_DefFont = New StdFont
    
    With m_DefFont
        .Name = "Arial"
        .Size = 8
        .Bold = False
    End With
    
        
        
    
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

Friend Property Let TotalItemCount(ByVal vData As Long)

    mvarTotalItemCount = vData

End Property

Friend Property Get TotalItemCount() As Long

    TotalItemCount = mvarTotalItemCount

End Property

Public Property Set Items(ByVal vData As Collection)

    Set mvarItems = vData

End Property

Public Property Get Items() As Collection

    Set Items = mvarItems

End Property

'//---------------------------------------------------------------------------------------
' Procedure : AddPropertyItem
' Type      : Function
' DateTime  : 04/10/2004 15:53
' Author    : Gary Noble
' Purpose   : Adds A Item To The Collection
' Returns   : CPropertyItem
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Public Function AddPropertyItem(ByVal Caption As String, ByVal Key As String, Optional ByRef Parent As CPropertyItem, Optional PropType As peEditorType = peTextEditor, Optional ListData As CListItems = Nothing) As CPropertyItem

    On Error GoTo errAdd

    Dim X      As CPropertyItem

    Set X = New CPropertyItem

    '-- bail
    If mvarItems Is Nothing Then Set mvarItems = New Collection

    '-- If The Parent is Not Nothing The We Add To The Parent Collection
    If Not Parent Is Nothing Then

        rParentPropertyItem(ObjPtr(Parent)).ChildCount = rParentPropertyItem(ObjPtr(Parent)).ChildCount + 1
        Set X = Parent.ChildProperties.AddPropertyItem(Caption, Key, , PropType, ListData)
        X.ListPTR = ObjPtr(ListData)
        X.Level = Parent.Level + 1
        X.ParentPTR = ObjPtr(Parent)

    Else                                              'NOT NOT...

        '-- Add new Top Level Item
        With X

            .Init m_lOwner, m_hWnd
            .Caption = Caption
            '-- Add It To The collection
            .ListPTR = ObjPtr(ListData)
            .PropertyType = PropType

            mvarItems.Add X, Key

        End With                                      'X
    End If

    TotalItemCount = TotalItemCount + 1

    X.Key = Key

    '  CopyMemory AddPropertyItem, x, 4
    '  CopyMemory AddPropertyItem, 0&, 4

CleanExit:

    '-- Return the Item
    Set AddPropertyItem = X

    Exit Function

errAdd:

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Set X = Nothing
    Resume CleanExit

End Function

Private Sub Class_Initialize()

    Set mvarItems = New Collection

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    Set mvarItems = Nothing
    Set m_DefFont = Nothing
    
    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : rParentPropertyItem
' Type      : Property
' DateTime  : 04/10/2004 15:55
' Author    : Gary Noble
' Purpose   : Returns A Item Object From its Pointer
' Returns   : CPropertyItem
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------

Friend Property Get rParentPropertyItem(ptr As Long) As CPropertyItem

    Dim xPropertyItem As CPropertyItem

    On Error GoTo ehHandleError_rParentPropertyItem
    CopyMemory xPropertyItem, ptr, 4

    Set rParentPropertyItem = xPropertyItem

    CopyMemory xPropertyItem, 0&, 4

    ZeroMemory ObjPtr(xPropertyItem), 4&

CleanExit:

    On Error GoTo 0

    Exit Property

ehHandleError_rParentPropertyItem:

    Resume CleanExit

End Property

