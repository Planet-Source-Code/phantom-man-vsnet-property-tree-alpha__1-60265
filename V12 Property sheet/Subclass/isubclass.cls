VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum EMsgResponse
    emrConsume                                        '-- Process instead of original WindowProc
    emrPostProcess                                    '-- Process after original WindowProc
    emrPreprocess                                     '-- Process before original WindowProc
End Enum
#If False Then
    Private emrConsume, emrPostProcess, emrPreprocess
#End If
Public MsgResponse As EMsgResponse

Public Function WindowProc(ByVal lngHwnd As Long, _
                           ByVal iMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

'<STUB> Reason: Interface procedure used by a class used with 'Implements'


End Function

'


