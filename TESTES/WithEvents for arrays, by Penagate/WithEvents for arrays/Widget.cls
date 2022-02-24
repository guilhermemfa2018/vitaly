VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Widget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Force the callback class to implement the interface defining the events.
Private mCallback As IWidgetEvents
'

' For test purposes:
Sub FireDummyEvent(ByVal Index As Long)
    ' Raise an event, passing a parameter
    If (Not mCallback Is Nothing) Then _
        mCallback.DummyEvent Index
End Sub
'

' Allow the callback object to be set. Very important.
Property Set Callback(ByRef newObj As IWidgetEvents)
    Set mCallback = newObj
End Property

Property Get Callback() As IWidgetEvents
    Set Callback = mCallback
End Property
    
