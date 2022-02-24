VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgramacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PPS - Plano de Programação Semanal (Operações em aberto)"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   22425
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   22425
   Begin VB.Frame Frame3 
      Caption         =   "Programar para a semana: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   615
         Left            =   1680
         Picture         =   "frmProgramacao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   17
         Text            =   "-"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   0
      Left            =   1320
      Picture         =   "frmProgramacao.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Imprimir PPS"
      ToolTipText     =   "Imprimir PPS"
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   720
      Picture         =   "frmProgramacao.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7560
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   120
      Picture         =   "frmProgramacao.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "Salvar PPS"
      ToolTipText     =   "Salvar PPS"
      Top             =   7560
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Configurações - OS nº: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   3255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmProgramacao.frx":3328
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Text            =   "-"
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":33A0
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Text            =   "1"
         Top             =   1320
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":340C
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   109707265
         CurrentDate     =   42500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Programação"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   109707265
         CurrentDate     =   41950
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   3255
      Begin VB.OptionButton optButton 
         Caption         =   "Extra PPS"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Atraso + PPS"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Não Programadas"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Programadas p/ semana:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgramacao.frx":3470
            Key             =   "linha"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmProgramacao.frx":394A
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmProgramacao.frx":39A4
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmProgramacao.frx":39FE
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":3A58
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":3ABE
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":3B3A
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8175
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   14420
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   8388608
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   2640
      OleObjectBlob   =   "frmProgramacao.frx":3BB4
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   8655
   End
End
Attribute VB_Name = "frmProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vQtdCols As Integer
Private vSomaMinutosOS As Double

Private vGuardaLegenda As String
Private vTotal3101SC01 As String
Private vTotal3101SC02 As String
Private vTotal3101SC03 As String
Private vTotal3101SC04 As String
Private vTotal3101SC05 As String
Private vTotal3101SC06 As String
Private vTotal3101SC07 As String
Private vTotal3101SC08 As String
Private vTotal3101SC09 As String
Private vTotal3101SC10 As String
Private vTotal3101SC12 As String
Private vTotal3102SC01 As String
Private vTotal3102SC02 As String
Private vTotal3103SC01 As String
Private vTotal3103SC02 As String
Private vTotal3104SC01 As String
Private vTotal3104SC02 As String
Private vTotal3105SC01 As String
Private vTotal3105SC02 As String
Private vTotal3106SC01 As String
Private vTotal4101SC03 As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        vGuardaLegenda = Principal.StatusBar1.Panels(3).Text
        TransfereDados
        FCRProgramacao.Show 1
    Case 12
        If GravaProgramacao = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
        Else
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ZEUS"
        End If
    Case 13
        Unload Me
    End Select
End Sub

Private Sub Command1_Click()
    InsereSemana
    SomaMinutos
End Sub

Private Sub DTPicker2_CloseUp()
    Text5.Text = DatePart("ww", DTPicker2.Value, vbMonday, vbFirstFourDays)
End Sub

Private Sub Form_Load()
    listview_cabecalho
    CompoeLV 0
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_Resize()
    frmProgramacao.Top = 0
    frmProgramacao.Left = 0
'    DimensionaPPS
'    listview_cabecalho
'    CompoeLV
End Sub

Private Sub listview_cabecalho()
    Dim rsCompoeCC As New ADODB.Recordset
    Dim SqlCompoeCC As String
    Dim Y As Integer
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "S.Prog.", ListView1.Width / 24
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 18
    ListView1.ColumnHeaders.Add , , "Total", ListView1.Width / 20
    ListView1.ColumnHeaders.Add , , "FCE", ListView1.Width / 28
    ListView1.ColumnHeaders.Add , , "Desenhos", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "OS", ListView1.Width / 21
    ListView1.ColumnHeaders.Add , , "Rev.", ListView1.Width / 32
    ListView1.ColumnHeaders.Add , , "TamLinha", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "DataProgramada", ListView1.Width / 10000
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview

    SqlCompoeCC = "select substring(codreduzido,6,10) from tbFormula where codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03') group by codreduzido order by codreduzido"
    rsCompoeCC.Open SqlCompoeCC, cnBanco, adOpenKeyset, adLockReadOnly
    While Not rsCompoeCC.EOF
        ListView1.ColumnHeaders.Add , , rsCompoeCC.Fields(0), ListView1.Width / 15
        rsCompoeCC.MoveNext
    Wend
    
    ListView1.ColumnHeaders.Add , , "S.Termino.", ListView1.Width / 24
    ListView1.ColumnHeaders.Add , , "Data Termino.", ListView1.Width / 24
    
    vQtdCols = ListView1.ColumnHeaders.Count
    rsCompoeCC.Close
    Set rsCompoeCC = Nothing
End Sub

Private Sub CompoeLV(vIndex As Integer)
    Dim rsCompoeLVCC As New ADODB.Recordset
    Dim SqlCompoeLVCC As String
    
    Dim rsDesenhoOS As New ADODB.Recordset
    Dim SqlDesenhoOS As String
    
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer, r As Integer, vOS As Integer, vRevisao As Integer
    Dim vCC As String, vSemanaProg As String
    ListView1.ListItems.Clear
    ListView1.Sorted = False

    SqlDesenhoOS = "Delete from tbDesenhosOS"
    rsDesenhoOS.Open SqlDesenhoOS, cnBanco

    SqlDesenhoOS = "insert into tbdesenhosOS SELECT a.idos,a.revisaoos,max(d.fce) as fce,min(g.desenho) as desenhos FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto and d.fce > 2000 " & _
                   "left join tbitemlm as f on SUBSTRING(a.desenhos,1,2) = f.codlm and replace(SUBSTRING(a.desenhos,3,4),';','') = f.codseq and d.fce = f.fce left join tbDesenhos as g on f.codigodes = g.iddesenho group by a.idos,a.revisaoos order by a.idos,a.revisaoos"
    rsDesenhoOS.Open SqlDesenhoOS, cnBanco

    'Compoe o lado esquerdo do ListView
    'Filtro - Todas
    If vIndex = 0 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,a.idos,right('00' + rtrim(a.revisaoos),2) as revisaoos,max(a.idoperacao) as operacao,max(d.fce) as fce,max(DATEPART(WK,a.dataprogramacao)) as SemanaPlanejada,max(a.dataprevista),MAX(g.desenho),max(DATEPART(WK,a.datatermino)) as SemanaTermino,max(a.datatermino) FROM tbMPItens as a " & _
                        "Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "left join tbdesenhosos as g on d.fce = g.fce and a.idos=g.idos and a.revisaoos = g.revisaoos where a.idos <> 0 and a.status <> 3 group by a.idos,a.revisaoos,a.dataprevista order by a.idos,a.revisaoos"
    'Filtro - Programadas para a semana: ??
    ElseIf vIndex = 1 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,a.idos,right('00' + rtrim(a.revisaoos),2) as revisaoos,max(a.idoperacao) as operacao,max(d.fce) as fce,max(DATEPART(WK,a.dataprogramacao)) as SemanaPlanejada,max(a.dataprevista),MAX(g.desenho),max(DATEPART(WK,a.datatermino)) as SemanaTermino,max(a.datatermino) FROM tbMPItens as a " & _
                        "Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "left join tbdesenhosos as g on d.fce = g.fce and a.idos=g.idos and a.revisaoos = g.revisaoos where a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) = '" & Val(Text1.Text) & "' group by a.idos,a.revisaoos,a.dataprevista order by a.idos,a.revisaoos"
    'Filtro - Não programadas
    ElseIf vIndex = 2 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,a.idos,right('00' + rtrim(a.revisaoos),2) as revisaoos,max(a.idoperacao) as operacao,max(d.fce) as fce,max(DATEPART(WK,a.dataprogramacao)) as SemanaPlanejada,max(a.dataprevista),MAX(g.desenho),max(DATEPART(WK,a.datatermino)) as SemanaTermino,max(a.datatermino) FROM tbMPItens as a " & _
                        "Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "left join tbdesenhosos as g on d.fce = g.fce and a.idos=g.idos and a.revisaoos = g.revisaoos where a.idos <> 0 and a.status <> 3 and a.dataprevista is null group by a.idos,a.revisaoos,a.dataprevista order by a.idos,a.revisaoos"
    'Filtro - Atraso + PPS
    ElseIf vIndex = 3 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,a.idos,right('00' + rtrim(a.revisaoos),2) as revisaoos,max(a.idoperacao) as operacao,max(d.fce) as fce,max(DATEPART(WK,a.dataprogramacao)) as SemanaPlanejada,max(a.dataprevista), " & _
                        "MAX(g.desenho),max(DATEPART(WK,a.datatermino)) as SemanaTermino,max(a.datatermino) FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao left join tbdesenhosos as g on d.fce = g.fce and a.idos=g.idos and a.revisaoos = g.revisaoos " & _
                        "where a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) <> DATEPART(WK,a.dataprogramacao) and DATEPART(WK,a.dataprevista) < '" & DatePart("ww", vDataDoBanco) & "' or " & _
                        "a.idos <> 0 and DATEPART(WK,a.dataprevista) < '" & DatePart("ww", vDataDoBanco) & "' and DATEPART(WK,a.dataprevista) = DATEPART(WK,a.dataprogramacao) or a.idos <> 0 and a.status <> 3 and a.dataprogramacao is null group by a.idos,a.revisaoos,a.dataprevista order by a.idos,a.revisaoos"
    'Filtro  - Extra PPS
    ElseIf vIndex = 4 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,a.idos,right('00' + rtrim(a.revisaoos),2) as revisaoos,max(a.idoperacao) as operacao,max(d.fce) as fce,max(DATEPART(WK,a.dataprogramacao)) as SemanaPlanejada,max(a.dataprevista), " & _
                        "MAX(g.desenho),max(DATEPART(WK,a.datatermino)) as SemanaTermino,max(a.datatermino) FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao left join tbdesenhosos as g on d.fce = g.fce and a.idos=g.idos and a.revisaoos = g.revisaoos " & _
                        "where a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) = DATEPART(WK,a.dataprogramacao) and DATEPART(WK,a.dataprevista) >= '" & DatePart("ww", vDataDoBanco) & "' group by a.idos,a.revisaoos,a.dataprevista order by a.idos,a.revisaoos"
    End If
    rsCompoeLVCC.Open SqlCompoeLVCC, cnBanco, adOpenKeyset, adLockReadOnly
    While Not rsCompoeLVCC.EOF
        If IsNull(rsCompoeLVCC.Fields(0)) Then
            Set ItemLst = ListView1.ListItems.Add(, , "-") 'Semana Programada
        Else
'            Set ItemLst = ListView1.ListItems.Add(, , rsCompoeLVCC.Fields(0)) 'Semana Programada
            Set ItemLst = ListView1.ListItems.Add(, , DatePart("ww", rsCompoeLVCC.Fields(6), vbMonday, vbFirstFourDays)) 'Semana Programada
        End If
        If rsCompoeLVCC.Fields(0) < DatePart("ww", vDataDoBanco, vbMonday, vbFirstFourDays) And rsCompoeLVCC.Fields(0) <> rsCompoeLVCC.Fields(5) Or rsCompoeLVCC.Fields(0) < DatePart("ww", vDataDoBanco, vbMonday, vbFirstFourDays) And IsNull(rsCompoeLVCC.Fields(5)) Then
            ItemLst.SubItems(1) = "Atraso" ' Status
        ElseIf rsCompoeLVCC.Fields(0) < DatePart("ww", vDataDoBanco, vbMonday, vbFirstFourDays) And rsCompoeLVCC.Fields(0) = rsCompoeLVCC.Fields(5) Then
            ItemLst.SubItems(1) = "Atraso/Extra PPS" ' Status
        ElseIf rsCompoeLVCC.Fields(0) = rsCompoeLVCC.Fields(5) And rsCompoeLVCC.Fields(0) >= DatePart("ww", vDataDoBanco, vbMonday, vbFirstFourDays) Then
            ItemLst.SubItems(1) = "Extra PPS" ' Status
        ElseIf IsNull(rsCompoeLVCC.Fields(0)) Then
            ItemLst.SubItems(1) = "-" ' Status
        Else
            ItemLst.SubItems(1) = "PPS" ' Status
        End If
        ItemLst.SubItems(2) = "-" ' Total
        ItemLst.SubItems(3) = rsCompoeLVCC.Fields(4) ' FCE
        If Not IsNull(rsCompoeLVCC.Fields(7)) Then ItemLst.SubItems(4) = Right(rsCompoeLVCC.Fields(7), 13) Else ItemLst.SubItems(4) = "-" ' Desenhos
        ItemLst.SubItems(5) = Format(rsCompoeLVCC.Fields(1), "000000") ' OS
        ItemLst.SubItems(6) = rsCompoeLVCC.Fields(2) ' Revisão OS
        If Not IsNull(rsCompoeLVCC.Fields(0)) Then ItemLst.SubItems(8) = rsCompoeLVCC.Fields(6) Else ItemLst.SubItems(8) = "null" ' Data Programada
        For r = 9 To vQtdCols - 1
            ItemLst.SubItems(r) = "-"
        Next
        If IsNull(rsCompoeLVCC.Fields(8)) Then
            ItemLst.SubItems(30) = "-" ' Semana Termino
        Else
            ItemLst.SubItems(30) = DatePart("ww", rsCompoeLVCC.Fields(9), vbMonday, vbFirstFourDays) ' Semana Termino
        End If
        If IsNull(rsCompoeLVCC.Fields(9)) Then
            ItemLst.SubItems(31) = "-" ' Semana Termino
        Else
            ItemLst.SubItems(31) = rsCompoeLVCC.Fields(9) ' Semana Termino
        End If
        
        
        rsCompoeLVCC.MoveNext
    Wend
    Me.ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    rsCompoeLVCC.Close
    Set rsCompoeLVCC = Nothing

    Dim vTempoCC As String
    Dim vHPlanejadas As String
    Dim vHProgramadas As String
    Dim vHAtraso As String
    
    'Compoe o lado direito do ListView
    If vIndex = 0 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.idos,a.revisaoos,max(a.idoperacao) as operacao,right('000000' + rtrim(a.idos),6)+right('000' + rtrim(a.revisaoos),3) +right('000' + rtrim(a.idoperacao),3) as OSREVOP,a.idcc,max(d.fce) as fce " & _
                        "FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "where a.idos <> 0 and a.status <> 3 group by a.tempocalc,a.idos,a.revisaoos,a.idoperacao,a.idcc order by " & _
                        "a.idos,a.revisaoos,SemanaProgramada,a.idoperacao,a.idcc"
    ElseIf vIndex = 1 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.idos,a.revisaoos,max(a.idoperacao) as operacao,right('000000' + rtrim(a.idos),6)+right('000' + rtrim(a.revisaoos),3) +right('000' + rtrim(a.idoperacao),3) as OSREVOP,a.idcc,max(d.fce) as fce " & _
                        "FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "where a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) = '" & Val(Text1.Text) & "' group by a.tempocalc,a.idos,a.revisaoos,a.idoperacao,a.idcc order by " & _
                        "a.idos,a.revisaoos,SemanaProgramada,a.idoperacao,a.idcc"
    ElseIf vIndex = 2 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.idos,a.revisaoos,max(a.idoperacao) as operacao,right('000000' + rtrim(a.idos),6)+right('000' + rtrim(a.revisaoos),3) +right('000' + rtrim(a.idoperacao),3) as OSREVOP,a.idcc,max(d.fce) as fce " & _
                        "FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "where a.idos <> 0 and a.status <> 3 and a.dataprevista is null group by a.tempocalc,a.idos,a.revisaoos,a.idoperacao,a.idcc order by a.idos,a.revisaoos,SemanaProgramada,a.idoperacao,a.idcc"
    
    ElseIf vIndex = 3 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.idos,a.revisaoos,max(a.idoperacao) as operacao,right('000000' + rtrim(a.idos),6)+right('000' + rtrim(a.revisaoos),3) +right('000' + rtrim(a.idoperacao),3) as OSREVOP,a.idcc,max(d.fce) as fce " & _
                        "FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao where a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) <> DATEPART(WK,a.dataprogramacao) and DATEPART(WK,a.dataprevista) < '" & DatePart("ww", vDataDoBanco) & "' or " & _
                        "a.idos <> 0 and DATEPART(WK,a.dataprevista) < '" & DatePart("ww", vDataDoBanco) & "' and DATEPART(WK,a.dataprevista) = DATEPART(WK,a.dataprogramacao) or a.idos <> 0 and a.status <> 3 and a.dataprogramacao is null group by a.tempocalc,a.idos,a.revisaoos,a.idoperacao,a.idcc order by a.idos,a.revisaoos,SemanaProgramada,a.idoperacao,a.idcc"
    ElseIf vIndex = 4 Then
        SqlCompoeLVCC = "Set datefirst 1 SELECT max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.idos,a.revisaoos,max(a.idoperacao) as operacao,right('000000' + rtrim(a.idos),6)+right('000' + rtrim(a.revisaoos),3) +right('000' + rtrim(a.idoperacao),3) as OSREVOP,a.idcc,max(d.fce) as fce " & _
                        "FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao where  a.idos <> 0 and a.status <> 3 and DATEPART(WK,a.dataprevista) = DATEPART(WK,a.dataprogramacao) and DATEPART(WK,a.dataprevista) >= '" & DatePart("ww", vDataDoBanco) & "' group by a.tempocalc,a.idos,a.revisaoos,a.idoperacao,a.idcc order by a.idos,a.revisaoos"
    End If
    rsCompoeLVCC.Open SqlCompoeLVCC, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCompoeLVCC.RecordCount = 0 Then
        mobjMsg.Abrir "Filtro não encontrado", Ok, critico, "Atenção"
        optButton(1).Value = False
        Exit Sub
    End If
    Y = ListView1.ListItems.Count
    vHPlanejadas = "0000:00"
    vHProgramadas = "0000:00"
    vHAtraso = "0000:00"
    rsCompoeLVCC.MoveFirst
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If Not IsNull(rsCompoeLVCC.Fields(0)) Then vSemanaProg = rsCompoeLVCC.Fields(0) Else vSemanaProg = "-"
        vOS = rsCompoeLVCC.Fields(2)
        If rsCompoeLVCC.Fields(3) <> "" Then vRevisao = rsCompoeLVCC.Fields(3) Else vRevisao = 0
        vTempoCC = "0000:00"
'        Do While vOS = Val(ListView1.SelectedItem.ListSubItems.Item(5)) And vRevisao = Val(ListView1.SelectedItem.ListSubItems.Item(6)) And Not rsCompoeLVCC.EOF
        Do While vSemanaProg = ListView1.ListItems.Item(X) And vOS = Val(ListView1.SelectedItem.ListSubItems.Item(5)) And vRevisao = Val(ListView1.SelectedItem.ListSubItems.Item(6)) And Not rsCompoeLVCC.EOF
            vCC = Mid$(rsCompoeLVCC.Fields(6), 6, 10)
             For r = 10 To vQtdCols
                If vCC = ListView1.ColumnHeaders.Item(r).Text Then
                    If rsCompoeLVCC.Fields(1) <> "" And rsCompoeLVCC.Fields(1) <> " " Then
                        ListView1.SelectedItem.ListSubItems.Item(r - 1) = rsCompoeLVCC.Fields(1)
                        somaTempoCCPPS rsCompoeLVCC.Fields(1), vTempoCC
                    End If
                End If
            Next
            rsCompoeLVCC.MoveNext
            If Not rsCompoeLVCC.EOF Then
                If Not IsNull(rsCompoeLVCC.Fields(0)) Then vSemanaProg = rsCompoeLVCC.Fields(0) Else vSemanaProg = "-"
                vOS = rsCompoeLVCC.Fields(2)
                vRevisao = rsCompoeLVCC.Fields(3)
            Else
                ListView1.SelectedItem.ListSubItems.Item(2) = vTempoCC
                somaTempoCCPPS vTempoCC, vHPlanejadas 'Calcula horas planejadas
                If ListView1.SelectedItem.ListSubItems.Item(1) <> "Atraso" Then
                    somaTempoCCPPS vTempoCC, vHProgramadas 'Calcula horas programadas
                End If
                If ListView1.SelectedItem.ListSubItems.Item(1) = "Atraso" Then
                    somaTempoCCPPS vTempoCC, vHAtraso 'Calcula horas de atraso
                End If
                SkinLabel4.Caption = vHPlanejadas
                SkinLabel5.Caption = vHProgramadas
                SkinLabel6.Caption = vHAtraso
                rsCompoeLVCC.Close
                Set rsCompoeLVCC = Nothing
                SomaCCs ListView1
                Exit Sub
            End If
        Loop
        ListView1.SelectedItem.ListSubItems.Item(2) = vTempoCC
        somaTempoCCPPS vTempoCC, vHPlanejadas 'Calcula horas planejadas
        If ListView1.SelectedItem.ListSubItems.Item(1) <> "Atraso" Then
            somaTempoCCPPS vTempoCC, vHProgramadas 'Calcula horas programadas
        End If
        If ListView1.SelectedItem.ListSubItems.Item(1) = "Atraso" Then
            somaTempoCCPPS vTempoCC, vHAtraso 'Calcula horas de atraso
        End If
    Next
End Sub

Private Function somaTempoCCPPS(vTempo, vOndeAcumula As String)
    If vTempo = "" Or vTempo = " " Then vTempo = "00:00"
    Dim seg As Long, min As Long, hora As Long
    Dim tempo As Long
    Dim matriz2

    matriz2 = Split(vTempo, ":")
    tempo = tempo + (CLng(matriz2(0)) * 3600)
    tempo = tempo + (CLng(matriz2(1)) * 60)
    
    If vOndeAcumula <> "" Then
        matriz2 = Split(vOndeAcumula, ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
    End If
    
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    
    vOndeAcumula = Format(hora, "0000") & ":" & Format(min, "00")
    somaTempoPPSAtraso = vOndeAcumula
End Function

Private Sub TransfereDados()
    Dim rsDeletaTabProg As New ADODB.Recordset
    Dim SqlDeletaTabProg As String
    
    Dim rsInsereDadoProg As New ADODB.Recordset
    Dim SqlInsereDadoProg As String
    
    Dim X As Integer, Y As Integer, r As Integer
    SqlDeletaTabProg = "Delete from tbPrintProgramacao"
    rsDeletaTabProg.Open SqlDeletaTabProg, cnBanco
    
    Y = ListView1.ListItems.Count
    If Y > 0 Then Principal.ProgressBar1.Max = Y
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Transferindo dados para tabela temporária"
    For X = 2 To Y
        Principal.ProgressBar1.Value = vProgress
        ListView1.ListItems.Item(X).Selected = True
        For r = 10 To vQtdCols
            SqlInsereDadoProg = "Insert into tbPrintProgramacao(" & _
                                "OS,revisao,semprog,status,fce,total,desenhos,hplanejadas,hprogramadas,hatraso,idcc,tempocc,total3101SC01,total3101SC02,total3101SC03,total3101SC04,total3101SC05,total3101SC06,total3101SC07,total3101SC08,total3101SC09,total3101SC10,total3101SC12,total3102SC01,total3102SC02,total3103SC01,total3103SC02,total3104SC01,total3104SC02,total3105SC01,total3105SC02,total3106SC01,total4101SC03) " & _
                                "values(" & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(5) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(6) & "', " & _
                                "'" & ListView1.ListItems.Item(X) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(1) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(3) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(2) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(4) & "', " & _
                                "'" & SkinLabel4.Caption & "','" & SkinLabel5.Caption & "', " & _
                                "'" & SkinLabel6.Caption & "','" & ListView1.ColumnHeaders.Item(r).Text & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems(r - 1).Text & "', " & _
                                "'" & vTotal3101SC01 & "','" & vTotal3101SC02 & "', " & _
                                "'" & vTotal3101SC03 & "','" & vTotal3101SC04 & "', " & _
                                "'" & vTotal3101SC05 & "','" & vTotal3101SC06 & "', " & _
                                "'" & vTotal3101SC07 & "','" & vTotal3101SC08 & "', " & _
                                "'" & vTotal3101SC09 & "','" & vTotal3101SC10 & "', " & _
                                "'" & vTotal3101SC12 & "','" & vTotal3102SC01 & "', " & _
                                "'" & vTotal3102SC02 & "','" & vTotal3103SC01 & "', " & _
                                "'" & vTotal3103SC02 & "','" & vTotal3104SC01 & "', " & _
                                "'" & vTotal3104SC02 & "','" & vTotal3105SC01 & "', " & _
                                "'" & vTotal3105SC02 & "', " & _
                                "'" & vTotal3106SC01 & "', " & _
                                "'" & vTotal4101SC03 & "')"
            rsInsereDadoProg.Open SqlInsereDadoProg, cnBanco
        Next
        vProgress = vProgress + 1
    Next
    Principal.ProgressBar1.Value = 0
    Legenda = vGuardaLegenda
    Principal.StatusBar1.Panels(3).Text = Legenda
    'FCRProgramacao.Show 1
End Sub

'SOMA COLUNA DE UM LISTVIEW EM HORAS
Private Sub SomaCCs(LV As ListView)
    'On Error Resume Next
    Dim X As Integer, Y As Integer, F As Integer
    Dim ItemLst As ListItem
    Y = LV.ListItems.Count
   
    vTotal3101SC01 = "0000:00"
    vTotal3101SC02 = "0000:00"
    vTotal3101SC03 = "0000:00"
    vTotal3101SC04 = "0000:00"
    vTotal3101SC05 = "0000:00"
    vTotal3101SC06 = "0000:00"
    vTotal3101SC07 = "0000:00"
    vTotal3101SC08 = "0000:00"
    vTotal3101SC09 = "0000:00"
    vTotal3101SC10 = "0000:00"
    vTotal3101SC12 = "0000:00"
    vTotal3102SC01 = "0000:00"
    vTotal3102SC02 = "0000:00"
    vTotal3103SC01 = "0000:00"
    vTotal3103SC02 = "0000:00"
    vTotal3104SC01 = "0000:00"
    vTotal3104SC02 = "0000:00"
    vTotal3105SC01 = "0000:00"
    vTotal3105SC02 = "0000:00"
    vTotal3106SC01 = "0000:00"
    vTotal4101SC03 = "0000:00"
    
    'GRAVA POSIÇÃO ATUAL
    'For X = 1 To Y
    '    If LV.ListItems.Item(X).Selected = True Then F = X
    'Next
    ListView1.Sorted = True
    ListView1.SortKey = 5
    ListView1.SortOrder = lvwAscending
    
    
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        If LV.SelectedItem.ListSubItems.Item(9) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(9), vTotal3101SC01
        If LV.SelectedItem.ListSubItems.Item(10) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(10), vTotal3101SC02
        If LV.SelectedItem.ListSubItems.Item(11) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(11), vTotal3101SC03
        If LV.SelectedItem.ListSubItems.Item(12) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(12), vTotal3101SC04
        If LV.SelectedItem.ListSubItems.Item(13) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(13), vTotal3101SC05
        If LV.SelectedItem.ListSubItems.Item(14) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(14), vTotal3101SC06
        If LV.SelectedItem.ListSubItems.Item(15) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(15), vTotal3101SC07
        If LV.SelectedItem.ListSubItems.Item(16) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(16), vTotal3101SC08
        If LV.SelectedItem.ListSubItems.Item(17) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(17), vTotal3101SC09
        If LV.SelectedItem.ListSubItems.Item(18) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(18), vTotal3101SC10
        If LV.SelectedItem.ListSubItems.Item(19) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(19), vTotal3101SC12
        If LV.SelectedItem.ListSubItems.Item(20) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(20), vTotal3102SC01
        If LV.SelectedItem.ListSubItems.Item(21) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(21), vTotal3102SC02
        If LV.SelectedItem.ListSubItems.Item(22) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(22), vTotal3103SC01
        If LV.SelectedItem.ListSubItems.Item(23) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(23), vTotal3103SC02
        If LV.SelectedItem.ListSubItems.Item(24) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(24), vTotal3104SC01
        If LV.SelectedItem.ListSubItems.Item(25) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(25), vTotal3104SC02
        If LV.SelectedItem.ListSubItems.Item(26) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(26), vTotal3105SC01
        If LV.SelectedItem.ListSubItems.Item(27) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(27), vTotal3105SC02
        If LV.SelectedItem.ListSubItems.Item(28) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(28), vTotal3106SC01
        If LV.SelectedItem.ListSubItems.Item(29) <> "-" Then somaTempoCCPPS LV.SelectedItem.ListSubItems.Item(29), vTotal4101SC03
    Next
    'RESTAURA POSIÇÃO ATUAL
    LV.ListItems.Item(1).Selected = True
    
    Set ItemLst = LV.ListItems.Add(, , "-") 'Semana Programada
    ItemLst.SubItems(1) = "-" ' Status
    ItemLst.SubItems(2) = "-" ' Total
    ItemLst.SubItems(3) = "-" ' FCE
    ItemLst.SubItems(4) = "-" ' Desenhos
    ItemLst.SubItems(5) = "-" ' OS
    ItemLst.SubItems(6) = "-" ' Revisão OS
    
    ItemLst.SubItems(9) = vTotal3101SC01 ' Revisão OS
    ItemLst.SubItems(10) = vTotal3101SC02 ' Revisão OS
    ItemLst.SubItems(11) = vTotal3101SC03 ' Revisão OS
    ItemLst.SubItems(12) = vTotal3101SC04 ' Revisão OS
    ItemLst.SubItems(13) = vTotal3101SC05 ' Revisão OS
    ItemLst.SubItems(14) = vTotal3101SC06 ' Revisão OS
    ItemLst.SubItems(15) = vTotal3101SC07 ' Revisão OS
    ItemLst.SubItems(16) = vTotal3101SC08 ' Revisão OS
    ItemLst.SubItems(17) = vTotal3101SC09 ' Revisão OS
    ItemLst.SubItems(18) = vTotal3101SC10 ' Revisão OS
    ItemLst.SubItems(19) = vTotal3101SC12 ' Revisão OS
    ItemLst.SubItems(20) = vTotal3102SC01 ' Revisão OS
    ItemLst.SubItems(21) = vTotal3102SC02 ' Revisão OS
    ItemLst.SubItems(22) = vTotal3103SC01 ' Revisão OS
    ItemLst.SubItems(23) = vTotal3103SC02 ' Revisão OS
    ItemLst.SubItems(24) = vTotal3104SC01 ' Revisão OS
    ItemLst.SubItems(25) = vTotal3104SC02 ' Revisão OS
    ItemLst.SubItems(26) = vTotal3105SC01 ' Revisão OS
    ItemLst.SubItems(27) = vTotal3105SC02 ' Revisão OS
    ItemLst.SubItems(28) = vTotal3106SC01 ' Revisão OS
    ItemLst.SubItems(29) = vTotal4101SC03 ' Revisão OS
    For X = 1 To 29
        ItemLst.ForeColor = &H8000&
        ItemLst.ListSubItems(X).ForeColor = &H8000&
        ItemLst.ListSubItems(X).Bold = True
    Next
End Sub

Private Sub InsereSemana()
    Dim X As Integer, Y As Integer, vGuardaPosAtual As Integer
    Y = ListView1.ListItems.Count
    
    'GUARDA POSIÇÃO ATUAL
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then vGuardaPosAtual = X
    Next
    'INSERE SEMANA PROGRAMADA NO QUE FOI CHECADO NO LISTVIEW
    For X = 2 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            If ListView1.ListItems.Item(X) <> "-" Then
                
                If IsDate(DTPicker2.Value) Then
                    Exit Sub
                End If
                
                If Val(ListView1.ListItems.Item(X)) <= DatePart("ww", CDate(vDataDoBanco), vbMonday, vbFirstFourDays) Then 'Val(Text2.Text) Then
                    mobjMsg.Abrir "A semana do item selecionado não pode ser modificada", Ok, critico, "Atenção"
                Else
                    ListView1.ListItems.Item(X) = Text2.Text
                    ListView1.SelectedItem.ListSubItems.Item(8) = ListView1.SelectedItem.ListSubItems.Item(8) & ";" & DTPicker1.Value
                    
                    If Val(ListView1.ListItems.Item(X)) = DatePart("ww", CDate(vDataDoBanco), vbMonday, vbFirstFourDays) Then
                        ListView1.SelectedItem.ListSubItems.Item(1) = "Extra PPS"
                    Else
                        ListView1.SelectedItem.ListSubItems.Item(1) = "PPS"
                    End If
                End If
            Else
                ListView1.ListItems.Item(X) = Text2.Text
                ListView1.SelectedItem.ListSubItems.Item(8) = ListView1.SelectedItem.ListSubItems.Item(8) & ";" & DTPicker1.Value
                If Val(ListView1.ListItems.Item(X)) = DatePart("ww", CDate(vDataDoBanco), vbMonday, vbFirstFourDays) Then
                    ListView1.SelectedItem.ListSubItems.Item(1) = "Extra PPS"
                Else
                    ListView1.SelectedItem.ListSubItems.Item(1) = "PPS"
                End If
            End If
        End If
    Next
    
    'RESTAURA POSIÇÃO ATUAL
    ListView1.ListItems.Item(vGuardaPosAtual).Selected = True
End Sub

Private Function GravaProgramacao()
'On Error GoTo Err
    GravaProgramacao = True
    Dim rsGeraProg As New ADODB.Recordset
    Dim SqlGeraProg As String
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer
    cnBanco.BeginTrans
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        
        If ListView1.ListItems.Item(X) <> "-" And Len(ListView1.SelectedItem.ListSubItems.Item(8)) > 10 Then
            converteSemana Val(ListView1.ListItems.Item(X)), DTPicker1, ""
        End If
        
        If Mid$(ListView1.SelectedItem.ListSubItems.Item(8), 1, 4) <> "null" And Len(ListView1.SelectedItem.ListSubItems.Item(8)) > 10 Then
'        If ListView1.SelectedItem.ListSubItems.Item(8) <> "-" And ListView1.SelectedItem.ListSubItems.Item(8) <> "" Then
            SqlGeraProg = "update tbMPItens set dataprevista = '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "',datatermino = '" & Format(ListView1.SelectedItem.ListSubItems.Item(31), "YYYY-MM-DD") & "' where status <> 3 and idos = '" & Val(ListView1.SelectedItem.ListSubItems.Item(5)) & "' and revisaoos = '" & Val(ListView1.SelectedItem.ListSubItems.Item(6)) & "'" & _
            "and dataprevista = '" & Format(Mid$(ListView1.SelectedItem.ListSubItems.Item(8), 1, 10), "YYYY-MM-DD") & "'"
            rsGeraProg.Open SqlGeraProg, cnBanco
        ElseIf Mid$(ListView1.SelectedItem.ListSubItems.Item(8), 1, 4) = "null" And Len(ListView1.SelectedItem.ListSubItems.Item(8)) > 10 Then
            SqlGeraProg = "update tbMPItens set dataprevista = '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "',datatermino = '" & Format(ListView1.SelectedItem.ListSubItems.Item(31), "YYYY-MM-DD") & "' where status <> 3 and idos = '" & Val(ListView1.SelectedItem.ListSubItems.Item(5)) & "' and revisaoos = '" & Val(ListView1.SelectedItem.ListSubItems.Item(6)) & "'" & _
            "and dataprevista is " & Mid$(ListView1.SelectedItem.ListSubItems.Item(8), 1, 4) & ""
            rsGeraProg.Open SqlGeraProg, cnBanco
        End If
        'End If
    Next
    'rsGeraProg.Update
    cnBanco.CommitTrans
'Err:
'    cnBanco.RollbackTrans
'    GravaProgramacao = False
End Function

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub ListView1_Click()
    Dim X As Integer, Y As Integer, vGuardaPosAtual As Integer
    Y = ListView1.ListItems.Count
    
    'GUARDA POSIÇÃO ATUAL
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then vGuardaPosAtual = X
    Next
    
    If IsDate(ListView1.SelectedItem.ListSubItems.Item(31)) Then
        For X = 2 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                If ListView1.SelectedItem.ListSubItems.Item(30) <> "-" Then
                    Text5.Text = ListView1.SelectedItem.ListSubItems.Item(30)
                    DTPicker2.Value = ListView1.SelectedItem.ListSubItems.Item(31)
                    Frame5.Caption = Mid$(Frame5.Caption, 1, 22) & " " & ListView1.SelectedItem.ListSubItems.Item(5) & "/" & ListView1.SelectedItem.ListSubItems.Item(6)
                Else
                    Text5.Text = ""
                    DTPicker2.Value = ""
                    Frame5.Caption = Mid$(Frame5.Caption, 1, 22) & " "
                End If
            End If
        Next
    Else
        Text5.Text = ""
        DTPicker2.Value = ""
        Frame5.Caption = Mid$(Frame5.Caption, 1, 22) & " "
    End If
    DTPicker2.Value = ""
    ListView1.ListItems.Item(vGuardaPosAtual).Selected = True
    'Text5.Text = ""
End Sub

Private Sub optButton_Click(Index As Integer)
    'listview_cabecalho
    If Index = 0 Then
        SkinLabel7 = "PPS Geral"
        CompoeLV 0
    ElseIf Index = 1 Then
        If Val(Text1.Text) > 53 Then
            Text1.Text = ""
            Text1.SetFocus
            optButton(1).Value = False
            Exit Sub
        End If
        If Text1.Text = "" Then
            mobjMsg.Abrir "Informe a semana que deseja filtrar", Ok, critico, "Atenção"
            optButton(1).Value = False
            Text1.SetFocus
        Else
            SkinLabel7 = "PPS da semana: " & Text1.Text
            CompoeLV 1
        End If
    ElseIf Index = 2 Then
        SkinLabel7 = "OS's não programadas"
        CompoeLV 2
    ElseIf Index = 3 Then
        SkinLabel7 = "PPS + Atraso"
        CompoeLV 3
    ElseIf Index = 4 Then
        SkinLabel7 = "Extra PPS"
        CompoeLV 4
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        optButton(1).Value = True
        If Val(Text1.Text) > 53 Then
            Text1.Text = ""
            Text1.SetFocus
            optButton(1).Value = False
            Exit Sub
        End If
        If Text1.Text = "" Then
            mobjMsg.Abrir "Informe a semana que deseja filtrar", Ok, critico, "Atenção"
            optButton(1).Value = False
            Text1.SetFocus
        Else
            CompoeLV 1
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text = "" Or Text2.Text = "-" Then
        Exit Sub
    End If
    If Val(Text2.Text) > 53 Then
        Text2.Text = ""
        DTPicker1.Value = ""
        Text2.SetFocus
        Exit Sub
    End If
    If Val(Text2.Text) < DatePart("ww", (vDataDoBanco), vbMonday, vbFirstFourDays) Then
        mobjMsg.Abrir "A semana a ser programada é MENOR que a semana atual. Deseja programar a semana para o PRÓXIMO ano?", YesNo, pergunta, "ZEUS"
        Dim vDataProg As String
        If Tp = 1 Then
            Command1.Enabled = True
            converteSemana Val(Text2.Text), DTPicker1, ""
            vDataProg = DTPicker1.Value
            DTPicker1.Value = Format(DatePart("d", vDataProg, vbMonday, vbFirstFourDays) & "/" & DatePart("m", vDataProg, vbMonday, vbFirstFourDays) & "/" & DatePart("yyyy", vDataProg) + 1, "dd/mm/yyyy", vbMonday, vbFirstFourDays)
        Else
            DTPicker1.Value = ""
            Text2.Text = ""
        End If
    Else
        Command1.Enabled = True
        converteSemana Val(Text2.Text), DTPicker1, ""
    End If
    'Text2.SetFocus
    'SomaMinutos
    'AchaSemanaFim DTPicker1.Value
End Sub

Private Sub SomaMinutos()
    'On Error Resume Next
    Dim X As Integer, Y As Integer, vGuardaPosAtual As Integer, j As Integer
    Y = ListView1.ListItems.Count
    
    'GUARDA POSIÇÃO ATUAL
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then vGuardaPosAtual = X
    Next
    'INSERE SEMANA PROGRAMADA NO QUE FOI CHECADO NO LISTVIEW
    For X = 2 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            vSomaMinutosOS = 0
            For j = 9 To 29
                If ListView1.SelectedItem.ListSubItems.Item(j) <> "-" Then
                    vSomaMinutosOS = vSomaMinutosOS + (Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(j), 1, 4)) * 60)
                    vSomaMinutosOS = vSomaMinutosOS + (Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(j), 4, 2)))
                End If
            Next
            AchaSemanaFim DTPicker1.Value
        End If
    Next
    'RESTAURA POSIÇÃO ATUAL
    ListView1.ListItems.Item(vGuardaPosAtual).Selected = True
End Sub

Private Sub AchaSemanaFim(vDataProgramada As String)
    Dim vMinutosTrabalhados As Integer
    Dim vDiasTrabalhadosOS As Integer
    Dim X As Integer
    Dim vDataFim As Date
    vMinutosTrabalhados = 528
    vDiasTrabalhadosOS = (vSomaMinutosOS / vMinutosTrabalhados) / Val(Text4.Text)
    vDataFim = DTPicker1.Value
    For X = 1 To vDiasTrabalhadosOS
        vDataFim = vDataFim + 1
        If DatePart("W", vDataFim, vbMonday, vbFirstFourDays) = 1 Or DatePart("W", vDataFim, vbMonday, vbFirstFourDays) = 7 Then
            vDataFim = vDataFim + 1
        End If
    Next
    If DTPicker2.Value = "" Or IsNull(DTPicker2.Value) Then
        ListView1.SelectedItem.ListSubItems.Item(30) = DatePart("ww", vDataFim, vbMonday, vbFirstFourDays)
        ListView1.SelectedItem.ListSubItems.Item(31) = vDataFim
    Else
        ListView1.SelectedItem.ListSubItems.Item(30) = Text5.Text
        ListView1.SelectedItem.ListSubItems.Item(31) = DTPicker2.Value
    End If
'    Msgbox "Data Termino:" & vDataFim & " - Semana Termino: " & DatePart("ww", vDataFim, vbMonday, vbFirstFourDays)
End Sub
