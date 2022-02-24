VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConvocacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convocação de treinamento"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "frmConvocacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin SGCH.chameleonButton cmdConvocacao 
      Height          =   615
      Index           =   3
      Left            =   8640
      TabIndex        =   28
      Tag             =   "Enviar convocação por email"
      ToolTipText     =   "Enviar convocação por email"
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConvocacao.frx":0CCA
      PICN            =   "frmConvocacao.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdConvocacao 
      Height          =   615
      Index           =   2
      Left            =   8040
      TabIndex        =   16
      Tag             =   "Imprimir convocação"
      ToolTipText     =   "Imprimir convocação"
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConvocacao.frx":19C0
      PICN            =   "frmConvocacao.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtConvocacao 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   27
      Top             =   6360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Programações"
      TabPicture(0)   =   "frmConvocacao.frx":26B6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Colaboradores"
      TabPicture(1)   =   "frmConvocacao.frx":26D2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Configurações"
      TabPicture(2)   =   "frmConvocacao.frx":26EE
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "mskConvocacao(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtConvocacao(1)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame6"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DTPicker1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtConvocacao(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtConvocacao 
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   8
         Tag             =   "Local do treinamento"
         ToolTipText     =   "Local do treinamento"
         Top             =   840
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   53608449
         CurrentDate     =   40896
      End
      Begin VB.Frame Frame6 
         Caption         =   "Horário: Início/Término"
         Height          =   735
         Left            =   6960
         TabIndex        =   24
         Top             =   1200
         Width           =   2175
         Begin MSMask.MaskEdBox mskConvocacao 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "Hora de início do curso/treinamento"
            ToolTipText     =   "Hora de início do curso/treinamento"
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskConvocacao 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Tag             =   "Hora de término do curso/treinamento"
            ToolTipText     =   "Hora de término do curso/treinamento"
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informe o tipo de convocação "
         Height          =   1455
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton optConvocacao 
            Caption         =   "Convocação Individual"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   2295
         End
         Begin VB.OptionButton optConvocacao 
            Caption         =   "Convocação Geral"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Selecione as programações as quais deseja emitir a convocação de treinamento "
         Height          =   5415
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   9015
         Begin SGCH.chameleonButton chameleonButton1 
            Height          =   615
            Left            =   600
            TabIndex        =   2
            Tag             =   "Listar colaboradores das programações selecionadas"
            ToolTipText     =   "Listar colaboradores das programações selecionadas"
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmConvocacao.frx":270A
            PICN            =   "frmConvocacao.frx":2726
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   4215
            Left            =   120
            TabIndex        =   0
            Top             =   1080
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   7435
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CheckBox Check1 
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione os colaboradores que deverão ser convocados "
         Height          =   5415
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   9015
         Begin VB.CheckBox Check4 
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4695
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   8281
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Texto "
         Height          =   3855
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   9015
         Begin VB.TextBox txtConvocacao 
            Height          =   3495
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   13
            Tag             =   "Texto"
            ToolTipText     =   "Texto"
            Top             =   240
            Width           =   8775
         End
      End
      Begin VB.TextBox txtConvocacao 
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Tag             =   "Responsável pela convocação"
         ToolTipText     =   "Responsável pela convocação"
         Top             =   1560
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mskConvocacao 
         Height          =   285
         Index           =   2
         Left            =   6960
         TabIndex        =   9
         Tag             =   "Carga horária do curso/treinamento"
         ToolTipText     =   "Carga horária do curso/treinamento"
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Carga horária:"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Local do Treinamento:"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Data:"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Responsável pela convocação:"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1320
         Width           =   2415
      End
   End
   Begin SGCH.chameleonButton cmdConvocacao 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConvocacao.frx":3400
      PICN            =   "frmConvocacao.frx":341C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdConvocacao 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Tag             =   "Salvar dados das configurações"
      ToolTipText     =   "Salvar dados das configurações"
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConvocacao.frx":40F6
      PICN            =   "frmConvocacao.frx":4112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmConvocacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private G As Integer
Private idConvocacao As Integer
Private vEmailAprovador As String
Private vEmailCol As String
Private vConvocado As String
Private vRegistro As String
Private vSetor As String

Private Sub chameleonButton1_Click()
    incCandConv
    SSTab1.Tab = 1
End Sub

Private Sub Check1_Click()
    MarcaDesmarca ListView2
End Sub

Private Sub Check4_Click()
    MarcaDesmarca ListView1
End Sub

Private Sub cmdConvocacao_Click(Index As Integer)
    Select Case Index
    Case 0
        If MsgBox("Deseja salvar os dados das configurações da convocação?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            gravaDados
        End If
    Case 1
        If MsgBox("Deseja sair da tela de convocação de treinamento?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Unload Me
            Set frmConvocacao = Nothing
        End If
    Case 2
        If MsgBox("Deseja salvar os dados das configurações da convocação?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            gravaDados
        End If
        If optConvocacao(0).Value = True Then
            FCRConvocacao.Show 1
        ElseIf optConvocacao(1).Value = True Then
            FCRConvInd.Show 1
        End If
    Case 3
'****************** EM DESENVOLVIMENTO ********************
        If dadosEmail = False Then Exit Sub
        Dim X As Integer, Y As Integer
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Checked = True Then
                ListView1.ListItems.Item(X).Selected = True
                vConvocado = ListView1.SelectedItem.ListSubItems.Item(1)
                vRegistro = ListView1.SelectedItem.ListSubItems.Item(2)
                vSetor = ListView1.SelectedItem.ListSubItems.Item(3)
                vEmailCol = ListView1.SelectedItem.ListSubItems.Item(4)
                If vEmailCol <> "" Then
                    If vSMTP <> "" Then
                        If enviaEmail = False Then Exit Sub
                    Else
                    End If
                End If
            End If
        Next
        MsgBox "Email enviado com sucesso!", vbInformation, "SGCH"
'******************
    End Select
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    listview_cabecalho
    CompoeProgs
    criaTabTemp
    'restauraDados
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "CPF", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Nome do colaborador", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "Registro", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Setor", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "Email", ListView1.Width / 3
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Programação", ListView1.Width / 7
    ListView2.ColumnHeaders.Add , , "Tipo", ListView1.Width / 5
    ListView2.ColumnHeaders.Add , , "Data Início", ListView1.Width / 8
    ListView2.ColumnHeaders.Add , , "Data Fim", ListView1.Width / 8
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    
End Sub

Private Sub CompoeProgs()
    Dim rsCompoeProgs As New ADODB.Recordset
    Dim sqlCompoeProgs As String
    
    Dim rsConfigProgs As New ADODB.Recordset
    Dim sqlConfigProgs As String
    
    sqlCompoeProgs = "select a.codprogramacao,b.tipo,c.datainicio,c.datafim from tbpendentescur as a inner join tbtreinamentos as b on a.codcoligada = '" & vCodColigada & "' and a.codtreinamento = b.codtreinamento inner join tbprogramacao as c on a.codprogramacao = c.codprogramacao " & _
                     "where a.status = 'Agendado' or a.status = 'Reagendado' group by a.codprogramacao,b.tipo,c.datainicio,c.datafim order by a.codprogramacao"
    rsCompoeProgs.Open sqlCompoeProgs, cnBanco, adOpenKeyset, adLockReadOnly
    While Not rsCompoeProgs.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsCompoeProgs.Fields(0), "000000")) 'codprogramacao
        ItemLst.SubItems(1) = "" & rsCompoeProgs.Fields(1) 'tipo
        ItemLst.SubItems(2) = "" & Format(rsCompoeProgs.Fields(2), "dd/mm/yyyy") 'Data de inicio da programação
        ItemLst.SubItems(3) = "" & Format(rsCompoeProgs.Fields(3), "dd/mm/yyyy") 'Data de fim da programação
        rsCompoeProgs.MoveNext
    Wend
    rsCompoeProgs.Close
    Set rsCompoeProgs = Nothing
    
    sqlConfigProgs = "select a.tipoconvocacao,a.tipotreinamento,a.responsavel,a.texto,a.dataconvocacao,a.horarioini,a.horariofim,a.cargahoraria,a.local from tbConfConvocacao as a where a.codcoligada = '" & vCodColigada & "'"
    rsConfigProgs.Open sqlConfigProgs, cnBanco, adOpenKeyset, adLockReadOnly
    If rsConfigProgs.RecordCount > 0 Then
        If rsConfigProgs.Fields(0) = 0 Then optConvocacao(0).Value = True
        If rsConfigProgs.Fields(0) = 1 Then optConvocacao(1).Value = True
        txtConvocacao(3) = rsConfigProgs.Fields(1)
        txtConvocacao(1) = rsConfigProgs.Fields(2)
        txtConvocacao(2) = rsConfigProgs.Fields(3)
        txtConvocacao(0) = rsConfigProgs.Fields(8)
        DTPicker1.Value = rsConfigProgs.Fields(4)
    
        mskConvocacao(0).PromptInclude = False
        mskConvocacao(1).PromptInclude = False
        mskConvocacao(2).PromptInclude = False
    
        mskConvocacao(0) = rsConfigProgs.Fields(5)
        mskConvocacao(1) = rsConfigProgs.Fields(6)
        mskConvocacao(2) = rsConfigProgs.Fields(7)
    
        mskConvocacao(0).PromptInclude = True
        mskConvocacao(1).PromptInclude = True
        mskConvocacao(2).PromptInclude = True
    End If
    rsConfigProgs.Close
    Set rsConfigProgs = Nothing
End Sub

Private Sub MarcaDesmarca(LV As ListView)
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems(X).Selected = True
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        Else
            LV.ListItems.Item(X).Checked = True
        End If
    Next
End Sub

Private Sub incCandConv() 'Incluir Filtrado no Processo Seletivo
    Dim Y As Integer, X As Integer, P As Integer
    Y = ListView2.ListItems.Count
    Dim rsCompoeLVColabs As New ADODB.Recordset
    Dim sqlCompoeLVColabs As String
    ListView1.ListItems.Clear
    txtConvocacao(3).Text = ""
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Selected = True
            sqlCompoeLVColabs = "select a.cpf,a.nomecolaborador,a.codcolaborador,d.nomesetor,a.email from tbcolaboradores as a inner join tbpendentescur as b on a.codcoligada = '" & vCodColigada & "' and a.ativo = 'S' and a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbsetores as d on c.codsetor = d.codsetor where codprogramacao = '" & Val(ListView2.ListItems.Item(X)) & "' order by a.nomecolaborador"
            rsCompoeLVColabs.Open sqlCompoeLVColabs, cnBanco, adOpenKeyset, adLockReadOnly
            If InStr(txtConvocacao(3), ListView2.SelectedItem.ListSubItems.Item(1)) = 0 Then
                If txtConvocacao(3) = "" Then
                    txtConvocacao(3) = ListView2.SelectedItem.ListSubItems.Item(1)
                Else
                    txtConvocacao(3) = txtConvocacao(3) & "/" & ListView2.SelectedItem.ListSubItems.Item(1)
                End If
            End If
            
            Dim ItemLst As ListItem
            Dim K As Integer, L As Integer, LV3Edit As String
            LV3Edit = ""
            For P = 1 To rsCompoeLVColabs.RecordCount
                Set ItemLst = ListView1.ListItems.Add(, , rsCompoeLVColabs.Fields(0))
                ItemLst.SubItems(1) = rsCompoeLVColabs.Fields(1)
                ItemLst.SubItems(2) = rsCompoeLVColabs.Fields(2)
                ItemLst.SubItems(3) = rsCompoeLVColabs.Fields(3)
                ItemLst.SubItems(4) = rsCompoeLVColabs.Fields(4)
                rsCompoeLVColabs.MoveNext
            Next
            rsCompoeLVColabs.Close
        End If
    Next
    Set rsCompoeLVColabs = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 1
    Me.ListView1.SortOrder = lvwDescending
End Sub

Private Sub gravaDados()
    If ValidaCampos = False Then Exit Sub
    
    Dim X As Integer
    Dim rsGravaConv As New ADODB.Recordset
    Dim sqlGravaConv As String
    
    sqlGravaConv = "Delete from tbConfConvocacao where codcoligada = '" & vCodColigada & "'"
    rsGravaConv.Open sqlGravaConv, cnBanco
    
    sqlGravaConv = "Select * from tbConfConvocacao where codcoligada = '" & vCodColigada & "'"
    rsGravaConv.Open sqlGravaConv, cnBanco, adOpenKeyset, adLockOptimistic
    
    rsGravaConv.AddNew
    If optConvocacao(0).Value = True Then rsGravaConv.Fields(1) = 0
    If optConvocacao(1).Value = True Then rsGravaConv.Fields(1) = 1
    rsGravaConv.Fields(2) = txtConvocacao(3).Text
    rsGravaConv.Fields(3) = txtConvocacao(1).Text
    rsGravaConv.Fields(4) = txtConvocacao(2).Text
    rsGravaConv.Fields(5) = DTPicker1
    rsGravaConv.Fields(6) = mskConvocacao(0)
    rsGravaConv.Fields(7) = mskConvocacao(1)
    rsGravaConv.Fields(8) = mskConvocacao(2)
    rsGravaConv.Fields(9) = txtConvocacao(0).Text
    rsGravaConv.Fields(10) = vCodColigada 'Codigo da coligada
    rsGravaConv.Update
    rsGravaConv.Close
    Set rsGravaConv = Nothing
    
    sqlGravaConv = "Select * from tbConfConvocacao where codcoligada = '" & vCodColigada & "'"
    rsGravaConv.Open sqlGravaConv, cnBanco, adOpenKeyset, adLockReadOnly
    idConvocacao = rsGravaConv.Fields(0)
    rsGravaConv.Close
    Set rsGravaConv = Nothing
    
    For G = 1 To ListView1.ListItems.Count
        GravaColaboradores
    Next
    MsgBox "Configurações salvas com sucesso!", vbInformation, "SGCH"
End Sub

Private Function ValidaCampos()
    ValidaCampos = False
    For X = 0 To 2
        If txtConvocacao(X).Text = "" Then
            MsgBox "Favor informar o campo " & Me.txtConvocacao(X).Tag, vbInformation, "Atenção"
            Me.txtConvocacao(X).SetFocus
            Exit Function
        End If
    Next
    If mskConvocacao(0) = "__:__" Then
        MsgBox "Não foi informado o horário de início do treinamento", vbInformation, "Atenção"
        Me.mskConvocacao(0).SetFocus
        Exit Function
    End If
    If mskConvocacao(1) = "__:__" Then
        MsgBox "Não foi informado o horário de término do treinamento", vbInformation, "Atenção"
        Me.mskConvocacao(1).SetFocus
        Exit Function
    End If
    If mskConvocacao(2) = "__:__" Then
        MsgBox "Não foi informado a carga horária", vbInformation, "Atenção"
        Me.mskConvocacao(2).SetFocus
        Exit Function
    End If
    
    ValidaCampos = True
End Function


Private Sub criaTabTemp()
On Error Resume Next
    'Criando uma tabela temporária global
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    SqlTabTemp = "CREATE TABLE ##tbColabsConvocados(ID int NOT NULL,CPF VARCHAR(20) NOT NULL,nomecolaborador VARCHAR(100) NOT NULL,registro VARCHAR(20) NOT NULL, setor VARCHAR(100) NOT NULL)"
    rsTabTemp.Open SqlTabTemp, cnBanco
End Sub
    
Private Sub GravaColaboradores()
On Error Resume Next
    ListView1.ListItems.Item(G).Selected = True
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    
    If ListView1.ListItems.Item(G).Checked = True Then
        sqlGravaColaboradores = "INSERT INTO ##tbColabsConvocados(ID,cpf,nomecolaborador,registro,setor) VALUES('" & idConvocacao & "','" & ListView1.ListItems.Item(G) & "','" & ListView1.SelectedItem.ListSubItems.Item(1) & "','" & ListView1.SelectedItem.ListSubItems.Item(2) & "','" & ListView1.SelectedItem.ListSubItems.Item(3) & "')"
        rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
    End If
End Sub

'As duas Subs abaixo faz com que ordene o listview pela coluna que vc clicar
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView2, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView2
    If .SortKey <> Column.Index - 1 Then
        .SortKey = Column.Index - 1
        .SortOrder = lvwAscending
    Else
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End If
    .Sorted = -1
    End With
End Sub

Private Function dadosEmail()
    dadosEmail = False
    Dim rsEnviaEmail As New ADODB.Recordset
    Dim SqlEnviaEmail As String
    SqlEnviaEmail = "Select email from tbUsuarios where codcoligada = '" & vCodColigada & "' and nome = '" & NomUsu & "'"
    rsEnviaEmail.Open SqlEnviaEmail, cnBanco, adOpenKeyset, adLockOptimistic
    vEmailAprovador = rsEnviaEmail.Fields(0)
    If vEmailAprovador = "" Then
        MsgBox "Email do usuário LOGADO não está cadastrado"
        Exit Function
    End If
    rsEnviaEmail.Close
    Set rsEnviaEmail = Nothing
    dadosEmail = True
End Function

Private Function enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
On Error GoTo errMail
    enviaEmail = False

    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
        .To = vEmailCol  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .From = vEmailAprovador  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "SGCH - CONVOCAÇÃO DE TREINAMENTO"
        
        .HTMLBody = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'><HTML xmlns='http://www.w3.org/TR/REC-html40' xmlns:o = 'urn:schemas-microsoft-com:office:office' xmlns:w = 'urn:schemas-microsoft-com:office:word'><HEAD><TITLE>Empresa Modelo – Industria</TITLE><META http-equiv=Content-Type content='text/html; charset=windows-1252'><META content=Word.Document name=ProgId><META content='MSHTML 6.00.2900.5512' name=GENERATOR><META content='Microsoft Word 11' name=Originator><LINK " & _
        "href='convoc_arquivos/filelist.xml' rel=File-List><!--[if gte mso 9]><xml> <o:DocumentProperties>  <o:Author>GMFA Inform&#225;tica</o:Author>  <o:Template>Normal</o:Template>  <o:LastAuthor>GMFA Inform&#225;tica</o:LastAuthor>  <o:Revision>2</o:Revision>  <o:TotalTime>9</o:TotalTime>  <o:Created>2011-12-23T12:03:00Z</o:Created><o:LastSaved>2011-12-23T12:03:00Z</o:LastSaved>  <o:Pages>1</o:Pages>  <o:Words>78</o:Words>  <o:Characters>425</o:Characters>  <o:Company>UBEC</o:Company>  <o:Lines>3</o:Lines>  <o:Paragraphs>1</o:Paragraphs>  <o:CharactersWithSpaces>502</o:CharactersWithSpaces>  <o:Version>11.9999</o:Version> </o:DocumentProperties>" & _
        "</xml><![endif]--><!--[if gte mso 9]><xml> <w:WordDocument>  <w:HyphenationZone>21</w:HyphenationZone>  <w:PunctuationKerning/>  <w:ValidateAgainstSchemas/>  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText><w:Compatibility>   <w:BreakWrappedTables/>   <w:SnapToGridInCell/>   <w:WrapTextWithPunct/>   <w:UseAsianBreakRules/>   <w:DontGrowAutofit/>  </w:Compatibility>  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel> </w:WordDocument></xml><![endif]--><!--[if gte mso 9]><xml> <w:LatentStyles DefLockedState='false' LatentStyleCount='156'> </w:LatentStyles></xml><![endif]--><STYLE><!-- /* Font Definitions */ @font-face   {font-family:Calibri;" & _
        "panose-1:2 15 5 2 2 2 4 3 2 4;mso-font-charset:0;mso-generic-font-family:swiss;mso-font-pitch:variable;mso-font-signature:-1610611985 1073750139 0 0 159 0;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal  {mso-style-parent:'';   margin:0cm; margin-bottom:.0001pt;  mso-pagination:widow-orphan;    font-size:12.0pt;   font-family:'Times New Roman';  mso-fareast-font-family:'Times New Roman';}p.MsoHeader, li.MsoHeader, div.MsoHeader{margin:0cm;margin-bottom:.0001pt;mso-pagination:widow-orphan;   tab-stops:center 212.6pt right 425.2pt; font-size:12.0pt;   font-family:'Times New Roman';  mso-fareast-font-family:'Times New Roman';}p.MsoFooter, li.MsoFooter, div.MsoFooter {margin:0cm;    margin-bottom:.0001pt;  mso-pagination:widow-orphan;    tab-stops:center 212.6pt right 425.2pt; font-size:12.0pt;   font-family:'Times New Roman';  mso-fareast-font-family:'Times New Roman';}" & _
        "/* Page Definitions */ @page   {mso-footnote-separator:url('convoc_arquivos/header.htm') fs;   mso-footnote-continuation-separator:url('convoc_arquivos/header.htm') fcs;  mso-endnote-separator:url('convoc_arquivos/header.htm') es; mso-endnote-continuation-separator:url('convoc_arquivos/header.htm') ecs;}@page Section1    {size:595.3pt 841.9pt;  margin:70.9pt 1.0cm 70.9pt 1.0cm;   mso-header-margin:35.45pt;  mso-footer-margin:35.45pt;  mso-paper-source:0;}div.Section1    {page:Section1;}--></STYLE><!--[if gte mso 10]><style> /* Style Definitions */ table.MsoNormalTable {mso-style-name:'Tabela normal';    mso-tstyle-rowband-size:0;  mso-tstyle-colband-size:0;  mso-style-noshow:yes;   mso-style-parent:'';    mso-padding-alt:0cm 5.4pt 0cm 5.4pt;    mso-para-margin:0cm;    mso-para-margin-bottom:.0001pt; mso-pagination:widow-orphan;    font-size:10.0pt;   font-family:'Times New Roman';  mso-ansi-language:#0400;    mso-fareast-language:#0400; mso-bidi-language:#0400;}" & _
        "table.MsoTableGrid {mso-style-name:'Tabela com grade'; mso-tstyle-rowband-size:0;  mso-tstyle-colband-size:0;  border:solid windowtext 1.0pt;  mso-border-alt:solid windowtext .5pt;   mso-padding-alt:0cm 5.4pt 0cm 5.4pt;    mso-border-insideh:.5pt solid windowtext;   mso-border-insidev:.5pt solid windowtext;   mso-para-margin:0cm;    mso-para-margin-bottom:.0001pt; mso-pagination:widow-orphan;    font-size:10.0pt;   font-family:'Times New Roman';  mso-ansi-language:#0400;mso-fareast-language:#0400; mso-bidi-language:#0400;}</style><![endif]--></HEAD><BODY lang=PT-BR style='tab-interval: 35.4pt'><DIV class=Section1><TABLE class=MsoTableGrid style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; MARGIN-LEFT: 14.4pt; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none; BORDER-COLLAPSE: collapse; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 480; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt; mso-border-insideh: .5pt solid windowtext; mso-border-insidev: .5pt solid windowtext' " & _
        "cellSpacing=0 cellPadding=0 border=1><TBODY><TR style='HEIGHT: 71.1pt; mso-yfti-irow: 0; mso-yfti-firstrow: yes'>    <TD     style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: windowtext 1pt solid; WIDTH: 513pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; HEIGHT: 71.1pt; mso-border-alt: solid windowtext .5pt'     vAlign=top width=684 colSpan=5>      <P class=MsoNormal       style='MARGIN-TOP: 6pt; LINE-HEIGHT: 150%; TEXT-ALIGN: center' " & _
        "align=center><B style='mso-bidi-font-weight: normal'><SPAN style='FONT-FAMILY: Calibri'>" & NomeEmpresa & "<o:p></o:p></SPAN></B></P>      <P class=MsoNormal style='LINE-HEIGHT: 150%; TEXT-ALIGN: center'       align=center><B style='mso-bidi-font-weight: normal'><SPAN       style='FONT-SIZE: 16pt; LINE-HEIGHT: 150%; FONT-FAMILY: Calibri'>CONVOCAÇÃO       DE TREINAMENTO<o:p></o:p></SPAN></B></P>      <P class=MsoNormal       style='LINE-HEIGHT: 150%; tab-stops: 225.75pt center 251.1pt'><SPAN       style='FONT-FAMILY: Calibri'><SPAN " & _
        "style='mso-tab-count: 2'></SPAN><CENTER>" & txtConvocacao(3) & "</CENTER><o:p></o:p></SPAN></P></TD></TR>  <TR style='mso-yfti-irow: 1'>    <TD     style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: windowtext 1pt solid; WIDTH: 216pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt'     vAlign=top width=288>      <P class=MsoNormal       style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Participante:<o:p></o:p></SPAN></P>" & _
        "<P class=MsoNormal style='mso-pagination: none; tab-stops: 18.0pt 36.0pt 54.0pt 72.0pt 90.0pt 108.0pt 126.0pt 144.0pt 162.0pt 180.0pt 198.0pt 216.0pt 234.0pt; mso-layout-grid-align: none'><B       style='mso-bidi-font-weight: normal'><SPAN       style='COLOR: black; FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'> " & vConvocado & " </SPAN></B><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'><o:p></o:p></SPAN></P></TD>    <TD     style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: medium none; WIDTH: 54.15pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "vAlign=top width=72><P class=MsoNormal style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Registro:<o:p></o:p></SPAN></P>      <P class=MsoNormal       style='mso-pagination: none; tab-stops: 18.0pt 36.0pt 54.0pt 72.0pt 90.0pt; mso-layout-grid-align: none'><B       style='mso-bidi-font-weight: normal'><SPAN       style='COLOR: black; FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'> " & vRegistro & " </SPAN></B><B       style='mso-bidi-font-weight: normal'><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'><o:p></o:p></SPAN></B></P></TD>    <TD " & _
        "style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: medium none; WIDTH: 242.85pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt'     vAlign=top width=324 colSpan=3>      <P class=MsoNormal       style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Setor:<o:p></o:p></SPAN></P>      <P class=MsoNormal       style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><B " & _
        "style='mso-bidi-font-weight: normal'><SPAN style='COLOR: black; FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & vSetor & "</SPAN></B><SPAN  style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'><o:p></o:p></SPAN></P></TD></TR>  <TR style='mso-yfti-irow: 2'>    <TD   style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: windowtext 1pt solid; WIDTH: 216pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt' " & _
        "vAlign=top width=288> <P class=MsoNormal style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN       style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Local:<o:p></o:p></SPAN></P>      <P class=MsoNormal       style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><B       style='mso-bidi-font-weight: normal'><SPAN    style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & txtConvocacao(0) & "<o:p></o:p></SPAN></B></P></TD>    <TD " & _
        "style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: medium none; WIDTH: 81pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt'     vAlign=top width=108 colSpan=2>      <P class=MsoNormal    style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN   style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Data:<o:p></o:p></SPAN></P>" & _
        "<P class=MsoNormal    style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><B    style='mso-bidi-font-weight: normal'><SPAN  style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & DTPicker1.Value & "<o:p></o:p></SPAN></B></P></TD>  <TD  style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: medium none; WIDTH: 108pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "vAlign=top width=144> <P class=MsoNormal style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN   style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Carga   horária:<o:p></o:p></SPAN></P>   <P class=MsoNormal   style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><B  style='mso-bidi-font-weight: normal'><SPAN   style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & mskConvocacao(2) & "<o:p></o:p></SPAN></B></P></TD> <TD  style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: medium none; WIDTH: 108pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "vAlign=top width=144>  <P class=MsoNormal style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><SPAN style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>Horário:<o:p></o:p></SPAN></P> <P class=MsoNormal style='MARGIN: 3pt 0cm; tab-stops: 225.75pt center 251.1pt'><B style='mso-bidi-font-weight: normal'><SPAN style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & mskConvocacao(0) & " às " & mskConvocacao(1) & "h" & "<o:p></o:p></SPAN></B></P></TD></TR>  <TR style='mso-yfti-irow: 3; mso-yfti-lastrow: yes'> <TD style='BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0cm; BORDER-LEFT: windowtext 1pt solid; WIDTH: 513pt; PADDING-TOP: 0cm; BORDER-BOTTOM: windowtext 1pt solid; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: solid windowtext .5pt' " & _
        "vAlign=top width=684 colSpan=5> <P class=MsoNormal  style='MARGIN: 3pt 0cm; TEXT-ALIGN: justify; mso-pagination: none; tab-stops: 18.0pt 36.0pt 54.0pt 72.0pt 90.0pt 108.0pt 126.0pt 144.0pt 162.0pt 180.0pt 198.0pt 216.0pt 234.0pt 252.0pt 270.0pt 288.0pt 306.0pt 324.0pt 342.0pt 360.0pt 378.0pt 396.0pt 414.0pt 432.0pt 450.0pt 468.0pt 486.0pt 504.0pt 522.0pt; mso-layout-grid-align: none'><SPAN style='COLOR: black; FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'>" & txtConvocacao(2) & "</SPAN><SPAN " & _
        "style='FONT-FAMILY: Calibri; mso-bidi-font-family: Arial'><o:p></o:p></SPAN></P></TD></TR><![if !supportMisalignedColumns]>  <TR height=0>    <TD style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none' width=288></TD><TD style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none'  width=72></TD><TD style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none' width=36></TD><TD style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none' " & _
        "width=144></TD> <TD style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none' width=144></TD></TR><![endif]></TBODY></TABLE><P class=MsoNormal><o:p>&nbsp;</o:p></P></DIV></BODY></HTML>"
        .Send
    End With
    enviaEmail = True
    Exit Function
errMail:
    MsgBox "Email não enviado para o usuário solicitante do PDO." & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SGCH"
    enviaEmail = False
    Exit Function
End Function


