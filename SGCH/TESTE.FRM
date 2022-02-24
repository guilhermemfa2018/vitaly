VERSION 5.00
Begin VB.Form frmCalc1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prototipo para calculo de área de materiais"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Operações:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCalc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private X As Double, Y As Double
Private calcula As Double
Private Conta As Integer
'Private Calc1 As String
'Private Calc2 As String
Private Operador As String

Private Sub Command2_Click()
    Y = X
    X = Len(Text1.Text) - 1
    If Conta > 1 Then
        If Operador = "*" Then
            calcula = calcula * Val(Mid$(Text1.Text, Y + 2, Len(Text1.Text) - Y))
        End If
        If Operador = "/" Then
            calcula = calcula / Val(Mid$(Text1.Text, Y + 2, Len(Text1.Text) - Y))
        End If
    ElseIf Conta = 1 Then
        If Operador = "*" Then
            calcula = Val(Mid$(Text1.Text, 1, X - 1)) * Val(Mid$(Text1.Text, Y + 2, Len(Text1.Text) - Y))
        End If
        If Operador = "/" Then
            calcula = Val(Mid$(Text1.Text, 1, X - 1)) / Val(Mid$(Text1.Text, Y + 2, Len(Text1.Text) - Y))
        End If
    End If
    Text2.Text = Format(calcula, "#,##0.00;(#,##0.00")
    Form_Load
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim vAr0(4) As String
    vAr0(1) = "a"
    vAr0(2) = "b"
    vAr0(3) = "c"
    vAr0(4) = "d"
    For X = 1 To 4
        MsgBox vAr0(X)
    Next
    
    
    
    Text1.Text = ""
    X = 1
    Y = 1
    calcula = 1
    Conta = 0
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Conta > 1 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Command2_Click
        ElseIf KeyCode = &H6A Then ' Tecla de Multiplicar no teclado Alfa-numerico
            Operacao
            Operador = "*"
        ElseIf KeyCode = &H6F Then ' Tecla de dividir no tecla do Alfa-numerico
            Operacao
            Operador = "/"
        End If
    ElseIf Conta = 1 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Command2_Click
        ElseIf KeyCode = &H6A Then ' Tecla de Multiplicar no teclado Alfa-numerico
            Operacao
            Operador = "*"
        ElseIf KeyCode = &H6F Then ' Tecla de dividir no teclado Alfa-numerico
            Operacao
            Operador = "/"
        End If
    ElseIf Conta = 0 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Conta = Conta + 1
            Command2_Click
        ElseIf KeyCode = &H6A Then ' Tecla de Multiplicar no teclado Alfa-numerico
            Operador = "*"
            Operacao
        ElseIf KeyCode = &H6F Then ' Tecla de dividir no teclado Alfa-numerico
            Operador = "/"
            Operacao
        End If
    End If
End Sub

Private Sub Operacao()
    If Conta > 1 Then
        Select Case Operador
        Case "/"
            calcula = calcula / Val(Mid$(Text1.Text, X + 2, Len(Text1.Text) - X))
        Case "*"
            calcula = calcula * Val(Mid$(Text1.Text, X + 2, Len(Text1.Text) - X))
        End Select
    ElseIf Conta = 1 Then
        Select Case Operador
        Case "/"
            calcula = Val(Mid$(Text1.Text, Y, X)) / Val(Mid$(Text1.Text, X + 2, Len(Text1.Text) - X))
        Case "*"
            calcula = Val(Mid$(Text1.Text, Y, X)) * Val(Mid$(Text1.Text, X + 2, Len(Text1.Text) - X))
        End Select
    End If
    Y = X
    X = Len(Text1.Text)
    Conta = Conta + 1
    Text2.Text = Format(calcula, "#,##0.00;(#,##0.00")
End Sub

