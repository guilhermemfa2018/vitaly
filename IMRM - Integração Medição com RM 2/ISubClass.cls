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
   emrConsume
   emrPostProcess
   emrPreprocess
End Enum
Public MsgResponse As EMsgResponse
Function WindowProc(ByVal Hwnd As Long, _
                    ByVal iMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
End Function
