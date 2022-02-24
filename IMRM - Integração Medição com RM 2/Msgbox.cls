VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Msgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Formato
    Nenhum = 0
    critico = 1
    exclamacao = 2
    informacao = 3
    pergunta = 4
End Enum

Public Enum Contexto
    Ok = 0
    OKCancel = 1
    YesNo = 2
    YesNoCancel = 3
    AbortRetryIgnore = 4
    RetryCancel = 5
End Enum

Public Sub Skin(ByVal CaMinho As String)
    Onde = CaMinho
End Sub

Public Sub Abrir(Mensagem As String, Optional Tipo As Contexto = Ok, Optional Icon As Formato = Nenhum, Optional Titulo As String)
On Error GoTo Err
    Load MsgMs

    'If Icon = critico Then MsgMs.critica.Visible = True
    'If Icon = informacao Then MsgMs.informacao.Visible = True
    'If Icon = informacao Then MsgMs.informacao.Visible = True
    'If Icon = pergunta Then MsgMs.pergunta.Visible = True

    MsgMs.Label1.Caption = Mensagem
    MsgMs.Caption = Titulo
    MsgMs.Forma
    
    If Tipo = Ok Then
        MsgMs.Command1.Caption = "Ok"
        MsgMs.Command2.Visible = False
        MsgMs.Command3.Visible = False
        MsgMs.Command1.Left = MsgMs.Width / 2 - MsgMs.Command1.Width / 2

    ElseIf Tipo = OKCancel Then
        MsgMs.Command1.Caption = "Ok"
        MsgMs.Command2.Caption = "Cancelar"
        MsgMs.Command3.Visible = False
        MsgMs.Command1.Left = MsgMs.Width / 2 - MsgMs.Command1.Width - 50
        MsgMs.Command2.Left = MsgMs.Width / 2 + 50
        
    ElseIf Tipo = YesNo Then
        MsgMs.Command1.Caption = "Sim"
        MsgMs.Command2.Caption = "Não"
        MsgMs.Command3.Visible = False
        MsgMs.Command1.Left = MsgMs.Width / 2 - MsgMs.Command1.Width - 50
        MsgMs.Command2.Left = MsgMs.Width / 2 + 50

    ElseIf Tipo = YesNoCancel Then
        MsgMs.Command1.Caption = "Sim"
        MsgMs.Command2.Caption = "Não"
        MsgMs.Command3.Caption = "Cancelar"

        MsgMs.Command2.Left = MsgMs.Width / 2 - MsgMs.Command2.Width / 2
        MsgMs.Command3.Left = MsgMs.Command2.Left + MsgMs.Command2.Width + 50
        MsgMs.Command1.Left = MsgMs.Command2.Left - MsgMs.Command2.Width - 50
        
    ElseIf Tipo = AbortRetryIgnore Then
        MsgMs.Command1.Caption = "Abortar"
        MsgMs.Command2.Caption = "Repetir"
        MsgMs.Command3.Caption = "Ignorar"
        MsgMs.Command2.Left = MsgMs.Width / 2 - MsgMs.Command2.Width / 2
        MsgMs.Command3.Left = MsgMs.Command2.Left + MsgMs.Command2.Width + 50
        MsgMs.Command1.Left = MsgMs.Command2.Left - MsgMs.Command2.Width - 50

    ElseIf Tipo = RetryCancel Then
        MsgMs.Command1.Caption = "Repetir"
        MsgMs.Command2.Caption = "Cancelar"
        MsgMs.Command3.Visible = False
        MsgMs.Command1.Left = MsgMs.Width / 2 - MsgMs.Command1.Width - 50
        MsgMs.Command2.Left = MsgMs.Width / 2 + 50
    End If
    MsgMs.Show 1
Err:
    Resume Next
End Sub

Public Property Get Resposta() As Integer
    Resposta = Tp
End Property

Private Sub Class_Initialize()
    Onde = Empty
    Onde1 = Empty
End Sub
