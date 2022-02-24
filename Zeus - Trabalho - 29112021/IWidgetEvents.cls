VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWidgetEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This interface class defines all the events that the Widget can raise.
' All events defined here must be blank, and all must be implemented
' by the handling class.
'
' It is not necessary to use an interface to define events, but if you
' do not, and set the Callback property of the object to an object that
' does not implement all event handler methods, you will run into an
' error if you then try to call an event handler that is not implemented.
' Using the interface forces the event handling class to implement all
' event handler methods and thus avoids any possible errors. Additionally,
' using an interface ensures strong early binding of the object reference.
' This makes for faster, more efficient code.
'

' Dummy event with a parameter.
Public Sub DummyEvent(ByVal Index As Long)
End Sub
