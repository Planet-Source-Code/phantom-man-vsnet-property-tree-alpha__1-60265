VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub AttachMessage(iwp As ISubclass, _
                         ByVal lngHwnd As Long, _
                         ByVal iMsg As Long)


    MSubclass.AttachMessage iwp, lngHwnd, iMsg

End Sub

Public Function CallOldWindowProc(ByVal lngHwnd As Long, _
                                  ByVal iMsg As Long, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) As Long


    CallOldWindowProc = MSubclass.CallOldWindowProc(lngHwnd, iMsg, wParam, lParam)

End Function

Public Property Get CurrentMessage() As Long

    CurrentMessage = MSubclass.CurrentMessage

End Property

Public Sub DetachMessage(iwp As ISubclass, _
                         ByVal lngHwnd As Long, _
                         ByVal iMsg As Long)


    MSubclass.DetachMessage iwp, lngHwnd, iMsg

End Sub


