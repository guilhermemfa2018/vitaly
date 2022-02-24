VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Callback As Class1
Private m_Index    As Long

Public Sub SomeEvent(ByVal Index As Long, ByVal Param As Long, ByRef RetVal As String)

End Sub

'The Friend scope modifier prevents the following properties and methods from
'being Implemented yet still allows them to be accessible within the project

Friend Property Get Callback(Optional ByVal Index As Long) As Class1
    Set Callback = m_Callback
End Property
Friend Property Set Callback(ByVal Index As Long, ByRef RHS As Class1)
    m_Index = Index
    Set m_Callback = RHS
End Property

Friend Sub TriggerAnEvent(ByVal Param As Long)
    Dim Frm As VB.Form, RetVal As String

    Debug.Assert Not m_Callback Is Nothing      'If code stops here, the Callback property must be set first!
    m_Callback.SomeEvent m_Index, Param, RetVal

    Set Frm = m_Callback
    Frm.Print RetVal
End Sub
