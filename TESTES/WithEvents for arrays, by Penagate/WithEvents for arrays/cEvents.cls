VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public WithEvents Button As CommandButton
Attribute Button.VB_VarHelpID = -1
 
Private m_idx As Long
 
Public Function Add_CommandButton(p_Btn As CommandButton, p_idx As Long)
    m_idx = p_idx
    Set Button = p_Btn
End Function
 
Private Sub Button_Click()
    Button.Parent.ButtonClick m_idx
End Sub
 
Private Sub Class_Terminate()
    Set Button = Nothing
End Sub
