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
 
Public WithEvents Img As Image
Attribute Img.VB_VarHelpID = -1
 
Private m_idx As Long
 
Public Function Add_Image(p_Btn As Image, p_idx As Long)
    m_idx = p_idx
    Set Img = p_Btn
End Function
 
Private Sub Img_Click()
    Img.Parent.ImgClick m_idx
End Sub
 
Private Sub Class_Terminate()
    Set Img = Nothing
End Sub

'Private Sub Img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'construirBotoes 1, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg", 360, 120, 615, 615, "Novo"
'End Sub

'Private Sub Img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'construirBotoes 1, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg", 360, 120, 615, 615, "Novo"
'End Sub
