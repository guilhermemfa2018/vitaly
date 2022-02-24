VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonitorar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoramento da Produção"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19515
   Icon            =   "frmMonitorar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   19515
   Begin VB.CommandButton Command1 
      Caption         =   "Encerrar apropriação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   30
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":0CCA
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":19A4
            Key             =   "FC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":267E
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame12 
      Caption         =   "Setor Selecionado"
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmMonitorar.frx":3358
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações da Apropriação "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   4455
      Begin VB.Frame Frame9 
         Caption         =   "Dados da Tarefa "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   4215
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":33B0
            TabIndex        =   26
            Top             =   960
            Width           =   3975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "frmMonitorar.frx":340A
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   1800
            OleObjectBlob   =   "frmMonitorar.frx":3464
            TabIndex        =   24
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmMonitorar.frx":34BE
            TabIndex        =   23
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":3518
            TabIndex        =   20
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":3572
            TabIndex        =   19
            Top             =   720
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "frmMonitorar.frx":35E6
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   1800
            OleObjectBlob   =   "frmMonitorar.frx":3650
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmMonitorar.frx":36B0
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":3718
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Apropriado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   3840
         Width           =   2055
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":3776
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Orçado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   2055
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":37D0
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Sub-centro"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":382A
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Satatus "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   2295
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":3884
            TabIndex        =   21
            Top             =   960
            Width           =   2055
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            ScaleHeight     =   585
            ScaleWidth      =   2025
            TabIndex        =   14
            Top             =   240
            Width           =   2055
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage4 
               Height          =   480
               Left            =   0
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":38DE
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage3 
               Height          =   480
               Left            =   480
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":45BC
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
               Height          =   480
               Left            =   960
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":529A
               Props           =   5
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Colaborador "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         Begin VB.PictureBox Picture1 
            Height          =   1815
            Left            =   120
            ScaleHeight     =   1755
            ScaleWidth      =   1515
            TabIndex        =   12
            Top             =   240
            Width           =   1575
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   1815
               Left            =   -120
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   3201
               Image           =   "frmMonitorar.frx":5F78
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colaboradores (Status)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   14775
      Begin MSComctlLib.ListView ListView3 
         Height          =   8175
         Left            =   10200
         TabIndex        =   29
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   14420
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   8775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   15478
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fábrica (Setores)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      Begin MSComctlLib.ListView ListView2 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Label Label53 
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   8880
      Width           =   4215
   End
End
Attribute VB_Name = "frmMonitorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vStatus As String
Private vSubCentro As String
Private vChapaEncerra As String
Private vPosition As Integer

Private Sub Command1_Click()
    EncerraAprop
End Sub

Private Sub Form_Load()
    listview_cabecalho
    CompoeLV
    HabBotao 1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_Resize()
    DimensionaForm
'    Dim XSize As Integer
'    Dim YSize As Integer
    
'    On Error Resume Next
'    If Form.WindowState <> 0 Then Exit Sub
    
'    Me.Top = 0
'    Me.Left = 0
'    Me.Height = Me.Height * YSize
'    Me.Width = Me.Width * XSize
    
'    For i = 0 To Me.Controls.Count - 1
'        Me.Controls(i).Left = Me.Controls(i).Left * XSize
'        Me.Controls(i).Top = Me.Controls(i).Top * YSize
'        Me.Controls(i).Height = Me.Controls(i).Height * YSize
'        Me.Controls(i).Width = Me.Controls(i).Width * XSize
'    Next i
End Sub

Private Sub ListView1_Click()
    CompoeControles
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    CompoeControles
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    CompoeControles
End Sub

Private Sub ListView2_Click()
    If ListView2.ListItems.Item(1).Selected = True Then
        SkinLabel20 = "PREPARAÇÃO"
        vSubCentro = "'3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3106.SC-01'"
    End If
    If ListView2.ListItems.Item(2).Selected = True Then
        SkinLabel20 = "MONTAGEM"
        vSubCentro = "'3000.3103.SC-01','3000.3103.SC-02'"
    End If
    If ListView2.ListItems.Item(3).Selected = True Then
        SkinLabel20 = "SOLDA"
        vSubCentro = "'3000.3104.SC-01','3000.3104.SC-02'"
    End If
    If ListView2.ListItems.Item(4).Selected = True Then
        SkinLabel20 = "ACABAMENTO"
        vSubCentro = "'3000.3105.SC-01','3000.3105.SC-02'"
    End If
    CompoeLV1
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Chapa", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Centro Custo", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "CC", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "OS", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "OS Rev.", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Operação", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Grupo", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Orçamento", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "C.Barra", ListView1.Width / 5
    ListView1.View = lvwList 'Modo de Exibição do seu Listview

    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "C.Barra", ListView3.Width / 3
    ListView3.ColumnHeaders.Add , , "Entrada", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "Saida", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "Parada", ListView3.Width / 4
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview

End Sub

Private Sub CompoeLV1()
    Dim rsStatus As New ADODB.Recordset
    Dim sqlStatus As String
'    sqlStatus = "select b.chapa,b.NOME,a.codigobarra,CONVERT (VARCHAR, a.dataent,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada,f.NOME,f.CODREDUZIDO,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status," & _
'                "c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
'                "from CORPORERM.dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join CORPORERM.dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join CORPORERM.dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra " & _
'                "where b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO <> 'D' or b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO = 'D' AND GETDATE ( )<b.DTDESLIGAMENTO+1 Order by f.CODREDUZIDO,b.NOME"
    
    
    sqlStatus = "Select b.chapa,b.NOME,a.codigobarra,CONVERT (VARCHAR, a.dataent,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada,f.NOME,f.CODREDUZIDO,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status,c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
                "from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra " & _
                "where b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO <> 'D' or b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO = 'D' AND GETDATE ( )<b.DTDESLIGAMENTO+1 " & _
                "union " & _
                "select b.chapa COLLATE SQL_Latin1_General_CP1_CI_AI,b.nome COLLATE SQL_Latin1_General_CP1_CI_AI,a.codigobarra COLLATE SQL_Latin1_General_CP1_CI_AI,CONVERT (VARCHAR, a.dataent ,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada COLLATE SQL_Latin1_General_CP1_CI_AI,b.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI,b.idcc COLLATE SQL_Latin1_General_CP1_CI_AI,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' " & _
                "when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status,c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
                "from tbTerceirizados as b left join tbOsMov as a on a.chapa = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra where b.ativo = 'S' and b.idcc in (" & vSubCentro & ") Order by f.CODREDUZIDO,b.NOME"
    
    
    rsStatus.Open sqlStatus, cnBanco, adOpenKeyset, adLockReadOnly
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView1.ListItems.Clear
    While Not rsStatus.EOF
        If rsStatus.Fields(10) = "A" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 1)
        ElseIf rsStatus.Fields(10) = "FC" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 2)
        ElseIf rsStatus.Fields(10) = "P" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 3)
        End If
        ItemLst.SubItems(1) = "" & rsStatus.Fields(0)
        ItemLst.SubItems(2) = "" & Mid$(rsStatus.Fields(8), 19, 30)
        ItemLst.SubItems(3) = "" & rsStatus.Fields(10)
        ItemLst.SubItems(4) = "" & Mid$(rsStatus.Fields(8), 1, 15)
        ItemLst.SubItems(5) = "" & rsStatus.Fields(11)
        ItemLst.SubItems(6) = "" & rsStatus.Fields(12)
        ItemLst.SubItems(7) = "" & rsStatus.Fields(13)
        ItemLst.SubItems(8) = "" & rsStatus.Fields(14)
        ItemLst.SubItems(9) = "" & rsStatus.Fields(15)
        ItemLst.SubItems(10) = "" & rsStatus.Fields(2)
        rsStatus.MoveNext
        X = X + 1
    Wend
    rsStatus.Close
    Set rsStatus = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub CompoeLV()
    Dim ItemLst As ListItem
    ListView2.ColumnHeaders.Add , , "", ListView2.Width / 1.1
    Set ItemLst = ListView2.ListItems.Add(, , "Preparação")
    Set ItemLst = ListView2.ListItems.Add(, , "Montagem")
    Set ItemLst = ListView2.ListItems.Add(, , "Solda")
    Set ItemLst = ListView2.ListItems.Add(, , "Acabamento")
End Sub

Private Sub CompoeControles()
    Dim mStream As ADODB.Stream
    
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            vPosition = X
            Exit For
        End If
    Next
    If ListView1.SelectedItem.ListSubItems.Item(3) = "A" Then  'Verde
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiNoGrayScale
        SkinLabel4.Caption = "APROPRIANDO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "FC" Then 'Laranja
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiNoGrayScale
        aicAlphaImage4.grayScale = aiCCIR709
        SkinLabel4.Caption = "OCIOSO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "P" Then 'Vermelho
        aicAlphaImage2.grayScale = aiNoGrayScale
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiCCIR709
        SkinLabel4.Caption = "PARADO"
    End If
    SkinLabel5.Caption = ListView1.SelectedItem.ListSubItems.Item(4)
    SkinLabel8.Caption = ListView1.SelectedItem.ListSubItems.Item(5)
    SkinLabel9.Caption = ListView1.SelectedItem.ListSubItems.Item(6)
    SkinLabel10.Caption = ListView1.SelectedItem.ListSubItems.Item(7)
    SkinLabel11.Caption = ListView1.SelectedItem.ListSubItems.Item(8)
    SkinLabel12.Caption = ListView1.SelectedItem.ListSubItems.Item(9)
    SkinLabel15.Caption = ListView1.SelectedItem.ListSubItems.Item(10)
    
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "A" Then
        SkinLabel5.Caption = "-"
        SkinLabel8.Caption = "-"
        SkinLabel9.Caption = "-"
        SkinLabel10.Caption = "-"
        SkinLabel11.Caption = "-"
        SkinLabel12.Caption = "-"
        SkinLabel13.Caption = "-"
        SkinLabel15.Caption = ListView1.SelectedItem.ListSubItems.Item(10)
        If SkinLabel15.Caption <> "" Then
            Dim rsAchaParada As New ADODB.Recordset
            Dim sqlAchaParada As String
            sqlAchaParada = "select a.nmparada from tbParadas as a where a.codigo = '" & SkinLabel15.Caption & "'"
            rsAchaParada.Open sqlAchaParada, cnBanco, adOpenKeyset, adLockReadOnly
            If rsAchaParada.RecordCount > 0 Then
                SkinLabel11 = rsAchaParada.Fields(0)
            End If
            rsAchaParada.Close
            Set rsAchaParada = Nothing
        End If
    End If
    
    'HABILITA BOTÃO PARA FINALIZAR APROPRIAÇÃO
    HabBotao X
    
    'PEGA IMAGEM GRAVADO NO BANCO SQL E EXIBE EM UM COMPONENTE DE IMAGEM
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        sqlCompoe = "select c.IDIMAGEM,a.chapa,a.nome,b.IMAGEM from " & vBancoTotvs & ".dbo.PFUNC as a left join " & vBancoTotvs & ".dbo.PPESSOA as c on a.CODPESSOA = c.CODIGO left join " & vBancoTotvs & ".dbo.GIMAGEM as b on c.IDIMAGEM = b.ID " & _
                    "where a.CHAPA = '" & ListView1.SelectedItem.ListSubItems.Item(1) & "'  order by a.nome"
    Else
        sqlCompoe = "select a.foto,a.chapa,a.nome,a.foto from tbTerceirizados as a where a.CHAPA = '" & ListView1.SelectedItem.ListSubItems.Item(1) & "' order by a.nome"
    End If
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        Set mStream = New ADODB.Stream
        mStream.Type = adTypeBinary
        mStream.Open
        mStream.Write rsCompoe.Fields(3).Value
        mStream.SaveToFile App.Path & "\Temp.jpg", adSaveCreateOverWrite
        aicAlphaImage1.ClearImage
        aicAlphaImage1.LoadImage_FromFile (App.Path & "\temp.jpg")
        Kill App.Path & "\Temp.jpg"
    Else
        label53 = rsCompoe.Fields(3) 'Local onde esta armazenado a foto do coloborador
        aicAlphaImage1.LoadImage_FromFile (label53.Caption)
    End If
    
    rsCompoe.Close
    Set rsCompoe = Nothing
    calculaTempoApropriado
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        compoeAprop Mid$(ListView1.ListItems.Item(X), 1, 5)
    Else
        compoeAprop Mid$(ListView1.ListItems.Item(X), 1, 11)
    End If
End Sub

Private Sub HabBotao(vPosicao As Integer)
On Error Resume Next
    Dim rsFimAprop As New ADODB.Recordset
    Dim sqlFimAprop As String
    sqlFimAprop = "Select a.multiplic from tbusuarios as a where a.nome= '" & NomUsu & "'"
    rsFimAprop.Open sqlFimAprop, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFimAprop.Fields(0) = "S" Then
        Command1.Visible = True
        If aicAlphaImage4.grayScale = aiNoGrayScale Or aicAlphaImage3.grayScale = aiNoGrayScale Then
            Command1.Enabled = True
        Else
            Command1.Enabled = False
        End If
    End If
    vChapaEncerra = Mid$(ListView1.ListItems.Item(vPosicao), 1, 5)
    
    rsFimAprop.Close
    Set rsFimAprop = Nothing
End Sub

Private Sub compoeAprop(vChapa As String)
    Dim rsAprop As New ADODB.Recordset
    Dim sqlAprop As String
    Dim ItemLst As ListItem
    
    sqlAprop = "select a.codigobarra,CONVERT (VARCHAR, a.horaent, 108) as entrada,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada from tbOsMov as a where a.dataent = CONVERT (date, GETUTCDATE()) and a.chapa = '" & vChapa & "' order by a.chapa,a.horaent"
    rsAprop.Open sqlAprop, cnBanco, adOpenKeyset, adLockReadOnly
    ListView3.ListItems.Clear
    While Not rsAprop.EOF
        Set ItemLst = ListView3.ListItems.Add(, , rsAprop.Fields(0))
        ItemLst.SubItems(1) = "" & rsAprop.Fields(1)
        ItemLst.SubItems(2) = "" & rsAprop.Fields(2)
        If rsAprop.Fields(2) <> "" And rsAprop.Fields(3) = "" Then
            ItemLst.SubItems(3) = "Baixa indevida"
            ItemLst.ListSubItems(3).ForeColor = &HC0&
        Else
            ItemLst.SubItems(3) = "" & rsAprop.Fields(3)
        End If
        rsAprop.MoveNext
    Wend
    rsAprop.Close
    Set rsAprop = Nothing
End Sub

Private Sub calculaTempoApropriado()
    Dim rsHAprop As New ADODB.Recordset
    Dim sqlHAprop As String
    Dim vHorasApropriadas As String
    
    sqlHAprop = "select CONVERT (VARCHAR, a.horasai-a.horaent, 108) as horaent from tbOsMov  as a where a.codigobarra = '" & ListView1.SelectedItem.ListSubItems.Item(10) & "'"
    rsHAprop.Open sqlHAprop, cnBanco, adOpenKeyset, adLockReadOnly
    vHorasApropriadas = "00:00"
    Do While Not rsHAprop.EOF
        If Not IsNull(rsHAprop.Fields(0)) Then somaTempoPPSAtraso rsHAprop.Fields(0), vHorasApropriadas
        rsHAprop.MoveNext
    Loop
    rsHAprop.Close
    Set rsHAprop = Nothing
    If SkinLabel12 = "-" Or SkinLabel12 = "" Then
        SkinLabel13 = "-"
    Else
        SkinLabel13 = vHorasApropriadas
    End If
End Sub

Private Sub EncerraAprop()
    Dim rsFimAprop As New ADODB.Recordset
    Dim sqlFimAprop As String
    Dim vID As String
    sqlFimAprop = "select id from tbOsMov where chapa = '" & vChapaEncerra & "' and datasai is null"
    rsFimAprop.Open sqlFimAprop, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFimAprop.RecordCount > 0 Then
        vID = rsFimAprop.Fields(0)
        rsFimAprop.Close
        Set rsFimAprop = Nothing
        If Time > "17:00:00" Then ' Maior que 17:00 horas
        'If Time < "17:00:00" Then ' Menor que 17:00 horas
            sqlFimAprop = "update tbOsMov set horasai = '17:00:00', datasai = '" & Format(Date, "YYYY-MM-DD") & "', idparada = '9018' where id = '" & vID & "'"
            rsFimAprop.Open sqlFimAprop, cnBanco
            CompoeLV1
            compoeAprop vChapaEncerra
            ListView1.ListItems.Item(vPosition).Selected = True
            CompoeControles
        ElseIf Time > "11:00:00" And Time < "12:00:00" Then
            sqlFimAprop = "update tbOsMov set horasai = '11:00:00', datasai = '" & Format(Date, "YYYY-MM-DD") & "', idparada = '9018' where id = '" & vID & "'"
            rsFimAprop.Open sqlFimAprop, cnBanco
            CompoeLV1
            compoeAprop vChapaEncerra
            ListView1.ListItems.Item(vPosition).Selected = True
            CompoeControles
        Else
            mobjMsg.Abrir "Apropriação fora do período de encerramento", , critico
        End If
    End If
End Sub

Private Function somaTempoPPSAtraso(vTempo, vOndeAcumula As String)
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
