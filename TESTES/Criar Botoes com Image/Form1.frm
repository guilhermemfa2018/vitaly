VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18030
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   18030
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   2400
      TabIndex        =   0
      Top             =   3720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7858
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   10800
      Picture         =   "Form1.frx":0054
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   9960
      Picture         =   "Form1.frx":0D1E
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   9120
      Picture         =   "Form1.frx":19E8
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   8400
      Picture         =   "Form1.frx":26B2
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   7560
      Picture         =   "Form1.frx":337C
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   6840
      Picture         =   "Form1.frx":4046
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3240
      Picture         =   "Form1.frx":4D10
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2520
      Picture         =   "Form1.frx":59DA
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1800
      Picture         =   "Form1.frx":66A4
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "Form1.frx":736E
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Image1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
    Image2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
    Image3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
    Image4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
    Image5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
    Image6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
    Image7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
    Image8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
    Image9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
    Image10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")

End Sub

Private Sub Image1_Click()
    'MsgBox ""
End Sub

'MouseDown
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
End Sub
    
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
End Sub
    
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
End Sub
    
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")

End Sub

'MouseUP
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_UP.jpg")
End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_UP.jpg")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_UP.jpg")
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_UP.jpg")
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_UP.jpg")
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_UP.jpg")
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_UP.jpg")
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_UP.jpg")

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    MsgBox SSTab1.Caption
End Sub

