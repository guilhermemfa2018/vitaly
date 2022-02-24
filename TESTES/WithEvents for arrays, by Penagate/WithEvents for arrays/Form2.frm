VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private WithEvents ctlDynamic As VBControlExtender
Option Explicit
Dim mo_Events As Collection
'
'Private objText(19, 15) As TextBox
'Private objFrame(19, 15) As Frame
'Private objCombo(19, 15) As ComboBox
'Private objLabel(19, 15) As Label
'Private objListview(19, 15) As MSComctlLib.Listview
''Private objButton1(19, 15) As VBControlExtender
''Private objButton(19, 15) As VBControlExtender
'Private objPicture(19, 15) As PictureBox
'
Private objImage As Image
'
''Novo
'Private WithEvents objImageNovo0 As Image, WithEvents objImageNovo1 As Image, WithEvents objImageNovo2 As Image, WithEvents objImageNovo3 As Image, WithEvents objImageNovo4 As Image, WithEvents objImageNovo5 As Image, WithEvents objImageNovo6 As Image, WithEvents objImageNovo7 As Image, WithEvents objImageNovo8 As Image, WithEvents objImageNovo9 As Image, WithEvents objImageNovo10 As Image, WithEvents objImageNovo11 As Image, WithEvents objImageNovo12 As Image, WithEvents objImageNovo13 As Image, WithEvents objImageNovo14 As Image, WithEvents objImageNovo15 As Image
''Editar
'Private WithEvents objImageEdita0 As Image, WithEvents objImageEdita1 As Image, WithEvents objImageEdita2 As Image, WithEvents objImageEdita3 As Image, WithEvents objImageEdita4 As Image, WithEvents objImageEdita5 As Image, WithEvents objImageEdita6 As Image, WithEvents objImageEdita7 As Image, WithEvents objImageEdita8 As Image, WithEvents objImageEdita9 As Image, WithEvents objImageEdita10 As Image, WithEvents objImageEdita11 As Image, WithEvents objImageEdita12 As Image, WithEvents objImageEdita13 As Image, WithEvents objImageEdita14 As Image, WithEvents objImageEdita15 As Image
''Excluir
'Private WithEvents objImageExclui0 As Image, WithEvents objImageExclui1 As Image, WithEvents objImageExclui2 As Image, WithEvents objImageExclui3 As Image, WithEvents objImageExclui4 As Image, WithEvents objImageExclui5 As Image, WithEvents objImageExclui6 As Image, WithEvents objImageExclui7 As Image, WithEvents objImageExclui8 As Image, WithEvents objImageExclui9 As Image, WithEvents objImageExclui10 As Image, WithEvents objImageExclui11 As Image, WithEvents objImageExclui12 As Image, WithEvents objImageExclui13 As Image, WithEvents objImageExclui14 As Image, WithEvents objImageExclui15 As Image
''Sair
'Private WithEvents objImageSai0 As Image, WithEvents objImageSai1 As Image, WithEvents objImageSai2 As Image, WithEvents objImageSai3 As Image, WithEvents objImageSai4 As Image, WithEvents objImageSai5 As Image, WithEvents objImageSai6 As Image, WithEvents objImageSai7 As Image, WithEvents objImageSai8 As Image, WithEvents objImageSai9 As Image, WithEvents objImageSai10 As Image, WithEvents objImageSai11 As Image, WithEvents objImageSai12 As Image, WithEvents objImageSai13 As Image, WithEvents objImageSai14 As Image, WithEvents objImageSai15 As Image
''Admitir
'Private WithEvents objImageAdmite0 As Image, WithEvents objImageAdmite1 As Image, WithEvents objImageAdmite2 As Image, WithEvents objImageAdmite3 As Image, WithEvents objImageAdmite4 As Image, WithEvents objImageAdmite5 As Image, WithEvents objImageAdmite6 As Image, WithEvents objImageAdmite7 As Image, WithEvents objImageAdmite8 As Image, WithEvents objImageAdmite9 As Image, WithEvents objImageAdmite10 As Image, WithEvents objImageAdmite11 As Image, WithEvents objImageAdmite12 As Image, WithEvents objImageAdmite13 As Image, WithEvents objImageAdmite14 As Image, WithEvents objImageAdmite15 As Image
''Filtro
'Private WithEvents objImageFiltra0 As Image, WithEvents objImageFiltra1 As Image, WithEvents objImageFiltra2 As Image, WithEvents objImageFiltra3 As Image, WithEvents objImageFiltra4 As Image, WithEvents objImageFiltra5 As Image, WithEvents objImageFiltra6 As Image, WithEvents objImageFiltra7 As Image, WithEvents objImageFiltra8 As Image, WithEvents objImageFiltra9 As Image, WithEvents objImageFiltra10 As Image, WithEvents objImageFiltra11 As Image, WithEvents objImageFiltra12 As Image, WithEvents objImageFiltra13 As Image, WithEvents objImageFiltra14 As Image, WithEvents objImageFiltra15 As Image
''Imprimir
'Private WithEvents objImageImprime0 As Image, WithEvents objImageImprime1 As Image, WithEvents objImageImprime2 As Image, WithEvents objImageImprime3 As Image, WithEvents objImageImprime4 As Image, WithEvents objImageImprime5 As Image, WithEvents objImageImprime6 As Image, WithEvents objImageImprime7 As Image, WithEvents objImageImprime8 As Image, WithEvents objImageImprime9 As Image, WithEvents objImageImprime10 As Image, WithEvents objImageImprime11 As Image, WithEvents objImageImprime12 As Image, WithEvents objImageImprime13 As Image, WithEvents objImageImprime14 As Image, WithEvents objImageImprime15 As Image
''Atualizar
'Private WithEvents objImageAtualiza0 As Image, WithEvents objImageAtualiza1 As Image, WithEvents objImageAtualiza2 As Image, WithEvents objImageAtualiza3 As Image, WithEvents objImageAtualiza4 As Image, WithEvents objImageAtualiza5 As Image, WithEvents objImageAtualiza6 As Image, WithEvents objImageAtualiza7 As Image, WithEvents objImageAtualiza8 As Image, WithEvents objImageAtualiza9 As Image, WithEvents objImageAtualiza10 As Image, WithEvents objImageAtualiza11 As Image, WithEvents objImageAtualiza12 As Image, WithEvents objImageAtualiza13 As Image, WithEvents objImageAtualiza14 As Image, WithEvents objImageAtualiza15 As Image
''Afastar
'Private WithEvents objImageAfasta0 As Image, WithEvents objImageAfasta1 As Image, WithEvents objImageAfasta2 As Image, WithEvents objImageAfasta3 As Image, WithEvents objImageAfasta4 As Image, WithEvents objImageAfasta5 As Image, WithEvents objImageAfasta6 As Image, WithEvents objImageAfasta7 As Image, WithEvents objImageAfasta8 As Image, WithEvents objImageAfasta9 As Image, WithEvents objImageAfasta10 As Image, WithEvents objImageAfasta11 As Image, WithEvents objImageAfasta12 As Image, WithEvents objImageAfasta13 As Image, WithEvents objImageAfasta14 As Image, WithEvents objImageAfasta15 As Image
''Programar
'Private WithEvents objImagePrograma0 As Image, WithEvents objImagePrograma1 As Image, WithEvents objImagePrograma2 As Image, WithEvents objImagePrograma3 As Image, WithEvents objImagePrograma4 As Image, WithEvents objImagePrograma5 As Image, WithEvents objImagePrograma6 As Image, WithEvents objImagePrograma7 As Image, WithEvents objImagePrograma8 As Image, WithEvents objImagePrograma9 As Image, WithEvents objImagePrograma10 As Image, WithEvents objImagePrograma11 As Image, WithEvents objImagePrograma12 As Image, WithEvents objImagePrograma13 As Image, WithEvents objImagePrograma14 As Image, WithEvents objImagePrograma15 As Image
'
'Private WithEvents objTeste As VBControlExtender
'
'
'Private Sub chameleonButton11_Click()
'    Msgbox "chameleonButton11"
'End Sub
'
'Private Sub cmdconsulta4_Click()
'    Msgbox "cmdconsulta4"
'End Sub
'
Private Sub Command1_Click()
    Dim vProximaTab As Integer, x As Integer
    x = 11
'    For vProximaTab = 0 To X
'        If SSTab1.TabVisible(vProximaTab) = False Then
'            Exit For
'        Else
'        End If
'    Next
'    If vProximaTab <= 11 Then
'        SSTab1.TabVisible(vProximaTab) = True
'        SSTab1.Tab = vProximaTab
'        construirControles vProximaTab
'    End If
    For vProximaTab = 0 To x
        If SSTab1.TabVisible(vProximaTab) = False Then
            Exit For
        Else
        End If
    Next
    If vProximaTab <= 10 Then
        SSTab1.TabVisible(vProximaTab) = True
        SSTab1.Tab = vProximaTab
        construirControles vProximaTab
        construirBotoes vProximaTab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg", 360, 120, 615, 615, "Novo"
        construirBotoes vProximaTab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg", 360, 720, 615, 615, "Editar"
        construirBotoes vProximaTab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_UP.jpg", 360, 1320, 615, 615, "Excluir"
        construirBotoes vProximaTab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_UP.jpg", 360, 1920, 615, 615, "Sair"
        construirBotoes vProximaTab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_UP.jpg", 360, 8040, 615, 615, "Admitir"
        construirBotoes vProximaTab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_UP.jpg", 360, 8640, 615, 615, "Filtrar"
        construirBotoes vProximaTab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_UP.jpg", 360, 9240, 615, 615, "Imprimir"
        construirBotoes vProximaTab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_UP.jpg", 360, 9840, 615, 615, "Atualizar"
        construirBotoes vProximaTab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_UP.jpg", 360, 10440, 615, 615, "Afastamento"
        construirBotoes vProximaTab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_UP.jpg", 360, 11040, 615, 615, "Programação"
    End If
    statusDados vProximaTab, True
End Sub

Private Sub Command2_Click()
    'descontruirControles SSTab1.Tab
    statusDados SSTab1.Tab, False
    SSTab1.TabVisible(SSTab1.Tab) = False
End Sub

Private Function statusDados(vTabAtiva As Integer, VouF As Boolean)
    If vTabAtiva = 0 Then
        'Frame1.Visible = VouF
    End If
    If vTabAtiva = 1 Then
        'Frame4.Visible = VouF
    End If
    If vTabAtiva = 2 Then
        'Frame7.Visible = VouF
    End If
    If vTabAtiva = 3 Then
        'Frame10.Visible = VouF
    End If
    If vTabAtiva = 4 Then
        'Frame13.Visible = VouF
    End If
    If vTabAtiva = 5 Then
        'Frame16.Visible = VouF
    End If
    If vTabAtiva = 6 Then
        'Frame19.Visible = VouF
    End If
    If vTabAtiva = 7 Then
        'Frame22.Visible = VouF
    End If
    If vTabAtiva = 8 Then
        'Frame25.Visible = VouF
    End If
    If vTabAtiva = 9 Then
        'Frame28.Visible = VouF
    End If
    If vTabAtiva = 10 Then
        'Frame31.Visible = VouF
    End If
End Function
'
'Private Function propImageNovo(imgNovo As Object)
'    With imgNovo
'        .Visible = True
'        .Top = 360
'        .Left = 120
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'        .Tag = "Novo"
'        .ToolTipText = "Novo"
'    End With
'End Function
'
'Private Function propImageEditar(imgEdita As Object)
'    With imgEdita
'        .Visible = True
'        .Top = 360
'        .Left = 720
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'        .Tag = "Editar"
'        .ToolTipText = "Editar"
'    End With
'End Function
'
'Private Function propImageExclui(imgExclui As Object)
'    With imgExclui
'        .Visible = True
'        .Top = 360
'        .Left = 1320
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_UP.jpg")
'        .Tag = "Excluir"
'        .ToolTipText = "Excluir"
'    End With
'End Function
'
'Private Function propImageSai(imgSai As Object)
'    With imgSai
'        .Visible = True
'        .Top = 360
'        .Left = 1920
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_UP.jpg")
'        .Tag = "Sair"
'        .ToolTipText = "Sair"
'    End With
'End Function
'
'Private Function propImageAdmite(imgAdmite As Object)
'    With imgAdmite
'        .Visible = True
'        .Top = 360
'        .Left = 8040
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_UP.jpg")
'        .Tag = "Admitir"
'        .ToolTipText = "Admitir"
'    End With
'End Function
'
'Private Function propImageFiltro(imgFiltro As Object)
'    With imgFiltro
'        .Visible = True
'        .Top = 360
'        .Left = 8640
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_UP.jpg")
'        .Tag = "Filtrar"
'        .ToolTipText = "Filtrar"
'    End With
'End Function
'
'Private Function propImageImprime(imgImprime As Object)
'    With imgImprime
'        .Visible = True
'        .Top = 360
'        .Left = 9240
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_UP.jpg")
'        .Tag = "Imprimir"
'        .ToolTipText = "Imprimir"
'    End With
'End Function
'
'Private Function propImageAtualiza(imgAtualiza As Object)
'    With imgAtualiza
'        .Visible = True
'        .Top = 360
'        .Left = 9840
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_UP.jpg")
'        .Tag = "Atualizar"
'        .ToolTipText = "Atualizar"
'    End With
'End Function
'
'Private Function propImageAfasta(imgAfasta As Object)
'    With imgAfasta
'        .Visible = True
'        .Top = 360
'        .Left = 10440
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_UP.jpg")
'        .Tag = "Afastamento"
'        .ToolTipText = "Afastamento"
'    End With
'End Function
'
'Private Function propImagePrograma(imgPrograma As Object)
'    With imgPrograma
'        .Visible = True
'        .Top = 360
'        .Left = 11040
'        .Width = 615
'        .Height = 615
'        .Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_UP.jpg")
'        .Tag = "Programação"
'        .ToolTipText = "Programação"
'    End With
'End Function
'
Private Function construirControles(vTab As Integer)
    Set objFrame(vTab, 0) = Controls.Add("VB.Frame", "Frame1" + Trim(Str(vTab)), SSTab1)
    With objFrame(vTab, 0)
        .Visible = True
        .Top = 360
        .Left = 120
        .Width = 16695
        .Height = 9015
        .Caption = "Informações"
    End With

    Set objFrame(vTab, 1) = Controls.Add("VB.Frame", "Frame0" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 1)
        .Visible = True
        .Top = 240
        .Left = 2760
        .Width = 5175
        .Height = 735
        .Caption = "Pesquisa"
    End With
    
    Set objPicture(vTab, 0) = Controls.Add("VB.PictureBox", "picBg" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objPicture(vTab, 0)
        .Visible = False
        .Top = 360
        .Left = 15600
        .Width = 855
        .Height = 495
    End With


    Set objFrame(vTab, 2) = Controls.Add("VB.Frame", "Frame3" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 2)
        .Visible = True
        .Top = 120
        .Left = 12360
        .Width = 3975
        .Height = 855
        .Caption = "Filtro "
        .Appearance = 0
        .BackColor = &H8000000F
    End With

    Set objLabel(vTab, 0) = Controls.Add("VB.Label", "Label1" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 120
        .Width = 735
        .Height = 255
        .Caption = "Status: "
    End With

    Set objLabel(vTab, 1) = Controls.Add("VB.Label", "Label3" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 1)
        .Visible = True
        .Top = 480
        .Left = 120
        .Width = 855
        .Height = 255
        .Caption = "Período: "
    End With

    Set objLabel(vTab, 2) = Controls.Add("VB.Label", "Label2" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 2)
        .Visible = True
        .Top = 240
        .Left = 960
        .Width = 2055
        .Height = 255
        .Caption = "-"
    End With

    Set objLabel(vTab, 3) = Controls.Add("VB.Label", "Label4" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 3)
        .Visible = True
        .Top = 480
        .Left = 960
        .Width = 2055
        .Height = 255
        .Caption = "-"
    End With

    Set objListview(vTab, 0) = Controls.Add("MSComctlLib.ListViewCtrl.2", "Listview2" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objListview(vTab, 0)
        .Visible = True
        .Top = 1080
        .Left = 120
        .Width = 16455
        .Height = 7695
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .LabelWrap = True
        .SortKey = 0
        .SortOrder = lvwAscending
        .View = lvwReport
        .BackColor = &H80000018
        .ForeColor = &H800000
    End With
    
End Function

Private Function construirBotoes(vTab As Integer, vBotao As Integer, vCaminho As String, vTop As Integer, vLeft As Integer, vWidth As Integer, vHeight As Integer, vTag As String)
On Error Resume Next
    Set objImage = Me.Controls.Add("VB.Image", "objImage" & vTab & vBotao, objFrame(vTab, 0))
    With objImage
        .Visible = True
        .Top = vTop
        .Left = vLeft
        .Width = vWidth
        .Height = vHeight
        .Picture = LoadPicture(vCaminho)
        .Tag = vTag
        .ToolTipText = vTag & vTab & vBotao
    End With
    mo_Events.Add New cEvents
    mo_Events(Val(vTab & vBotao)).Add_Image objImage, Val(vTab & vBotao)
End Function
'
'Private Function descontruirControles(vTab As Integer)
'On Error Resume Next
'    Dim i As Long
'    For i = 0 To 15
'        Me.Controls.Remove objFrame(vTab, i).Name
'        Me.Controls.Remove objText(vTab, i).Name
'        Me.Controls.Remove objButton(vTab, i).Name
'        Me.Controls.Remove objCombo(vTab, i).Name
'        Me.Controls.Remove objLabel(vTab, i).Name
'        Me.Controls.Remove objListview(vTab, i).Name
'        Me.Controls.Remove objButton1(vTab, i).Name
'        Me.Controls.Remove objPicture(vTab, i).Name
'    Next
'End Function
'
Private Function desconstroiTabs()
    Dim i As Long
    For i = 0 To 10
        SSTab1.TabVisible(i) = False
    Next
End Function

Private Sub Form_Load()
'    vPosAtual = 1
    Set mo_Events = New Collection
    'AplicarSkin Me, Principal.Skin1
    'NewColorDBGrid Me
    'On Error GoTo ErrHandler
    'MudaPropPicture 'Configura Picture para colorir as linhas do listview de acordo com o Tipo de FCE
    'configControles

'DimensionaLV "Métodos e Processos"
    'desconstroiTabs
    Exit Sub
'ErrHandler:
'    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub
'
'
''--------- CLIQUES ---------------------------
'
'Private Sub objImageNovo0_Click()
'    Msgbox objImageNovo0.Tag & " 0"
'End Sub
'
'Private Sub objImageNovo1_Click()
'    Msgbox objImageNovo1.Tag & " 0"
'End Sub
'
'Private Sub objImageNovo2_Click()
'    Msgbox objImageNovo2.Tag & " 2"
'End Sub
'
'Private Sub objImageNovo3_Click()
'    Msgbox objImageNovo3.Tag & " 3"
'End Sub
'
'Private Sub objImageNovo4_Click()
'    Msgbox objImageNovo4.Tag + 4
'End Sub
'
'' NOVO MouseDown
'Private Sub objImageNovo0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'Private Sub objImageNovo14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'Private Sub objImageNovo15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg")
'End Sub
'
'' EDITAR MouseDown
'Private Sub objImageEdita0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'Private Sub objImageEdita14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'Private Sub objImageEdita15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_DOWN.jpg")
'End Sub
'
'
'' EXCLUIR MouseDown
'Private Sub objImageExclui0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'Private Sub objImageExclui14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'Private Sub objImageExclui15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageExclui15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_DOWN.jpg")
'End Sub
'
'' SAIR MouseDown
'Private Sub objImageSai0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'Private Sub objImageSai14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'Private Sub objImageSai15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageSai15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_DOWN.jpg")
'End Sub
'
'' ADMITIR MouseDown
'Private Sub objImageAdmite0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'Private Sub objImageAdmite14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'Private Sub objImageAdmite15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAdmite15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_DOWN.jpg")
'End Sub
'
'' FILTRO MouseDown
'Private Sub objImageFiltra0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'Private Sub objImageFiltra14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'Private Sub objImageFiltra15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageFiltra15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_DOWN.jpg")
'End Sub
'
'' IMPRIMIR MouseDown
'Private Sub objImageImprime0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'Private Sub objImageImprime14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'Private Sub objImageImprime15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageImprime15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_DOWN.jpg")
'End Sub
'
'' ATUALIZAR MouseDown
'Private Sub objImageAtualiza0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'Private Sub objImageAtualiza14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'Private Sub objImageAtualiza15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAtualiza15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_DOWN.jpg")
'End Sub
'
'' AFASTAMENTO MouseDown
'Private Sub objImageAfasta0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'Private Sub objImageAfasta14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'Private Sub objImageAfasta15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageAfasta15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_DOWN.jpg")
'End Sub
'
'' PROGRAMACAO MouseDown
'Private Sub objImagePrograma0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'Private Sub objImagePrograma14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'Private Sub objImagePrograma15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImagePrograma15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_DOWN.jpg")
'End Sub
'
'
'
'
'
'
'' NOVO MouseUp
'Private Sub objImageNovo0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'Private Sub objImageNovo15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageNovo15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg")
'End Sub
'
'
'' EDITAR MouseUp
'Private Sub objImageEdita0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita0.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita1.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita2.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita3.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita4.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita5.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita6.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita7.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita8.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita9.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita10.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita11.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita12.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita13.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita14.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'Private Sub objImageEdita15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    objImageEdita15.Picture = LoadPicture("D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg")
'End Sub
'
'
'
'
'
'
'

Public Sub ImgClick(p_idx As Long)
    MsgBox "Button is clicked # " & p_idx

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mo_Events = Nothing
End Sub

