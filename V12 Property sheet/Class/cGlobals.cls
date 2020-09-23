VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlobals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum peEditorType
    peTextEditor = 2                                  '(default) Editor will be a TextBox.
    peNumberEditor = 3                                'Editor will be a TextBox which will accept only numbers.
    peFontEditor = 5                                  'Editor will be a font selection dialog.
    peFileEditor = 7                                  'Editor will be a file selection dialog.
    peDirectoryEditor = 8                             'Editor will be a folder selection dialog.
    peListEditor = 9                                  'Editor will be a drop down listbox.
    peColorEditor = 10                                'Editor will be drop-down color picker.
    peSpinButton = 12                                 '-- Spin button Control
    pePictureEditor = 13
    peReadOnlyPushButton = 14
    peReadOnly = 15
End Enum

Public Enum peListDataType
    peColorList = 0
    peSimpleDropDown = 1
End Enum

Public Type tSpinButtonConstraints
    lMin       As Long
    lMax       As Long
    lIncrementValue As Long
End Type



Public Enum eSpinButtonConstraints
    Min = 0
    Max = 1
    lIncrementValue = 2
End Enum

Public Enum eBorderStyle
    Sunken = 0
    Thin = 1
    None = 2
End Enum

