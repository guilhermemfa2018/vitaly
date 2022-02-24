VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos"
   ClientHeight    =   8640
   ClientLeft      =   3270
   ClientTop       =   1275
   ClientWidth     =   6585
   Icon            =   "frmGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmGrupos.frx":0CCA
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmGrupos.frx":0D3C
      TabIndex        =   15
      Top             =   240
      Width           =   615
   End
   Begin ZEUS.chameleonButton chameleonButton11 
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7920
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
      MICON           =   "frmGrupos.frx":0DA8
      PICN            =   "frmGrupos.frx":0DC4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton chameleonButton12 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7920
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
      MICON           =   "frmGrupos.frx":1A9E
      PICN            =   "frmGrupos.frx":1ABA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   7920
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox mskCadastro 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   " Pemissões "
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   4440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGrupos.frx":2794
               Key             =   "sim"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGrupos.frx":31A6
               Key             =   "nao"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Permissões de tela "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   5280
         Width           =   6015
         Begin VB.Frame Frame4 
            Caption         =   "Nome menu "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3000
            TabIndex        =   17
            Top             =   240
            Visible         =   0   'False
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmGrupos.frx":3BB8
               TabIndex        =   18
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Filtrar"
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   14
            Top             =   840
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Imprimir"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   13
            Top             =   480
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Editar"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Incluir"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Salvar"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   10
            Top             =   480
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Excluir"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            Top             =   840
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4980
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   8784
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsGrupo As New ADODB.Recordset
Private SqlGrupo As String

Private rsSalvar As New ADODB.Recordset
Private Status As String

Private Sub chameleonButton11_Click()
    Unload Me
    Set frmGrupos = Nothing
End Sub

Private Sub chameleonButton12_Click()
    mobjMsg.Abrir "Deseja salvar os dados do Grupo?", YesNo, pergunta, "ZEUS"
    If Tp = 1 Then
        Bot_salvar
        gravaLog "Código grupo: " & mskCadastro(0), "Nome grupo: " & txtCadastro(0), ""
        Unload Me
    End If
End Sub

Private Sub chkGravar_Click(Index As Integer)
    'If chkGravar(7).Value = 1 Then
    '    chkGravar(9).Enabled = True
    '    chkGravar(10).Enabled = True
    'Else
    '    chkGravar(9).Value = 0
    '    chkGravar(10).Value = 0
    '    chkGravar(9).Enabled = False
    '    chkGravar(10).Enabled = False
    'End If
End Sub

Private Sub Command1_Click()
'    IncluiTreeview
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    If Status = "novo" Then
        LimpaControles
        montaEstrutTreeview
    ElseIf Status = "editar" Then
        ResultPesq
        montaEstrutTreeview
        CompoeTreeview
    End If
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlGrupo = "select * from tbgrupo where codcoligada = '" & vCodcoligada & "' and codigo ='" & Val(varGlobal) & "'"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGrupo.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Grupo não encontrado", Ok, critico, "Atenção"
    End If
    rsGrupo.Close
    Set rsGrupo = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    For X = 0 To mskCadastro.Count - 1
        mskCadastro(X) = ""
    Next
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X) = ""
    Next
    mskCadastro(0).Text = Format(GeraCodigo, "000000") & ""
        
End Sub

Private Sub CompoeControles()
    mskCadastro(0).Enabled = False
    mskCadastro(0).Text = Format(rsGrupo.Fields(0), "000000") & ""
    txtCadastro(0).Text = rsGrupo.Fields(1)
    If rsGrupo.Fields(2) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Sub CompoeTreeview()
On Error GoTo Err
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim i As Integer
    SqlTreeview = "Select * from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and idmenu <> 0 and idgrupo = '" & Val(Me.mskCadastro(0)) & "' and tipo <> 'CAT' order by id"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    With TreeView1
        For i = 1 To .Nodes.Count
            .Nodes(i).Expanded = True
            If rsTreeview.EOF Then Exit For
            If rsTreeview.Fields(5) = "S" Then
                TreeView1.Nodes(i).Checked = True
            Else
                TreeView1.Nodes(i).Checked = False
            End If
            rsTreeview.MoveNext
        Next
    End With
    rsTreeview.Close
    Set rsTreeview = Nothing

    SqlTreeview = "Select * from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and idmenu = 0 and idgrupo = '" & Val(Me.mskCadastro(0)) & "' order by id"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    rsTreeview.Close
    Set rsTreeview = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    AbrirGrupo
    SqlGera = "Select top 1 * from tbGrupo where codcoligada = '" & vCodcoligada & "' order by codigo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGrupo.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    mskCadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharGrupo
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Private Sub AbrirGrupo()
On Error GoTo Err
    SqlGrupo = "Select * from tbGrupo where codcoligada = '" & vCodcoligada & "' Order by codigo"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockOptimistic
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub FecharGrupo()
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub

Private Sub Bot_salvar()
On Error GoTo Err
    Dim SqlSalvar As String
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer, i As Integer
10  cnBanco.BeginTrans
    SqlSalvar = "Select * from tbGrupo where codcoligada = '" & vCodcoligada & "' and tbGrupo.codigo = '" & Val(Me.mskCadastro(0)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    mskCadastro(0).PromptInclude = False
    If txtCadastro(0).Text <> "" Then
        mskCadastro(0).PromptInclude = False
        If mskCadastro(0).Text <> "" Then
            If rsSalvar.RecordCount = 0 Then
                rsSalvar.AddNew
                rsSalvar.Fields(0) = mskCadastro(0).ClipText
                rsSalvar.Fields(1) = txtCadastro(0).Text
                If Check1.Value = 0 Then
                    rsSalvar.Fields(2) = "N"
                Else
                    rsSalvar.Fields(2) = "S"
                End If
                rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
                LimpaControles
            Else
                rsSalvar.Fields(1) = txtCadastro(0)
                If Check1.Value = 0 Then
                    rsSalvar.Fields(2) = "N"
                Else
                    rsSalvar.Fields(2) = "S"
                End If
                rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
            End If
            rsSalvar.Update
        Else
            mobjMsg.Abrir "Favor Preencher o campo código!", Ok, critico, "Atenção"
        End If
    Else
        mobjMsg.Abrir "Favor Preencher o campo Descrição!", Ok, critico, "Atenção"
    End If
    rsSalvar.Close
    Set rsSalvar = Nothing
    Dim xInc As String, xEdi As String, xSal As String, xExc As String, xImp As String, xFil As String, xAva As String, xAdi As String, xDem As String, xAdiRes As String, xAdiResp As String
    
    If chkGravar(0).Value = 0 Then xInc = "N" Else xInc = "S"
    If chkGravar(1).Value = 0 Then xEdi = "N" Else xEdi = "S"
    If chkGravar(2).Value = 0 Then xSal = "N" Else xSal = "S"
    If chkGravar(3).Value = 0 Then xExc = "N" Else xExc = "S"
    If chkGravar(4).Value = 0 Then xImp = "N" Else xImp = "S"
    If chkGravar(5).Value = 0 Then xFil = "N" Else xFil = "S"
    
    SqlSalvar = "Delete from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and tbConfGrupo.idgrupo = '" & Val(Me.mskCadastro(0)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
'===========================
    Dim rsMenuExpert As New ADODB.Recordset
    Dim sqlMenuExpert As String
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    sqlMenuExpert = "Select * from tbMenuConf order by idsub"
    rsMenuExpert.Open sqlMenuExpert, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsMenuExpert.RecordCount > 0 Then
        While Not rsMenuExpert.EOF
            SqlMenu = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",'" & rsMenuExpert.Fields(0) & "','" & rsMenuExpert.Fields(1) & "','" & rsMenuExpert.Fields(2) & "','" & rsMenuExpert.Fields(3) & "','S','" & rsMenuExpert.Fields(5) & "','" & rsMenuExpert.Fields(6) & "')"
            rsMenu.Open SqlMenu, cnBanco
            rsMenuExpert.MoveNext
        Wend
        rsMenuExpert.Close
        Set rsMenuExpert = Nothing
    Else
        SqlSalvar = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'01','TAB','Cadastros','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'01','CAT','Primários','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'02','CAT','Secundários','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0101','BUT','Ramo de atividades','S','" & vCodcoligada & "',1);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0102','BUT','Clientes','S','" & vCodcoligada & "',2);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0103','BUT','Transportadoras','S','" & vCodcoligada & "',3);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0104','BUT','Tipo material','S','" & vCodcoligada & "',4);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0205','BUT','Materiais','S','" & vCodcoligada & "',5);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0206','BUT','Itens verificação','S','" & vCodcoligada & "',6);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0207','BUT','Projetos','S','" & vCodcoligada & "',7);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'0208','BUT','Processos','S','" & vCodcoligada & "',8);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'02','TAB','Orçamentos','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'11','CAT','Vendas','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'1111','BUT','Serviços','S','" & vCodcoligada & "',9);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'03','TAB','Planejamento','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'21','CAT','Planejamento e Controle de Produção','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'2121','BUT','FCE','S','" & vCodcoligada & "',10);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'2122','BUT','LM','S','" & vCodcoligada & "',11);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'2123','BUT','LD','S','" & vCodcoligada & "',12);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'2124','BUT','OS','S','" & vCodcoligada & "',13);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'2125','BUT','Controle de Desenhos','S','" & vCodcoligada & "',28);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'04','TAB','Produção','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'31','CAT','Acompanhamento de Produção','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'3131','BUT','OS Acompamenhamento','S','" & vCodcoligada & "',13);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'3132','BUT','Evolução','S','" & vCodcoligada & "',14);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'05','TAB','Inspeção/Expedição','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'41','CAT','Emissão de Relatórios','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'4141','BUT','Emitir Relatório','S','" & vCodcoligada & "',15);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'4142','BUT','Imprimir relatório','S','" & vCodcoligada & "',16);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'06','TAB','Configurações','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'51','CAT','Parametrizações','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'52','CAT','Aparência','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5151','BUT','Sistema','S','" & vCodcoligada & "',17);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5152','BUT','Grupos','S','" & vCodcoligada & "',18);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5153','BUT','Usuários','S','" & vCodcoligada & "',19);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5254','BUT','Menu','S','" & vCodcoligada & "',20);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5255','BUT','Skin','S','" & vCodcoligada & "',21);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5256','BUT','Fundo','S','" & vCodcoligada & "',22);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'07','TAB','Sobre','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'61','CAT','Sobre','S','" & vCodcoligada & "',0);" & _
                    "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'6161','BUT','Sobre ZEUS','S','" & vCodcoligada & "',23);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'6162','BUT','Ajuda do ZEUS','S','" & vCodcoligada & "',24);"

        rsSalvar.Open SqlSalvar, cnBanco
    End If

    SqlSalvar = "Select * from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and tbConfGrupo.idmenu <> 0 and tbConfGrupo.idgrupo = '" & Val(Me.mskCadastro(0)) & "' and tbConfGrupo.tipo <> 'CAT' order by id"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    On Error Resume Next
    With TreeView1
        For i = 1 To .Nodes.Count
            If TreeView1.Nodes(i).Checked = True Then
                rsSalvar.Fields(5) = "S" 'grava S se checkbox estiver marcado
            Else
                rsSalvar.Fields(5) = "N" 'grava N se checkbox estiver marcado
            End If
            rsSalvar.MoveNext
        Next
    End With
    cnBanco.CommitTrans
    rsSalvar.Close
    Set rsSalvar = Nothing
    AtualizaListview
    mobjMsg.Abrir "Dados gravados com sucesso!", Ok, informacao, "ZEUS"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = vListViewPrincipal.ListItems.Count
    For X = 1 To Y
        If vListViewPrincipal.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = vListViewPrincipal.ListItems.Add(, , Format(mskCadastro(0), "000000"))
        ItemLst.SubItems(1) = txtCadastro(0).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "OK"
        End If
    Else
        vListViewPrincipal.SelectedItem.ListSubItems.Item(1) = txtCadastro(0).Text
        If Check1.Value = 0 Then
            vListViewPrincipal.SelectedItem.ListSubItems.Item(2) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            vListViewPrincipal.SelectedItem.ListSubItems.Item(2) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(2).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub montaEstrutTreeview()
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim X As Integer
    
    SqlTreeview = "select a.*,b.incluir,b.editar,b.excluir,b.salvar,b.imprimir,b.filtrar from tbMenu as a left join tbConfGrupo as b on a.idmenu = b.idmenu and a.nome = b.nome and b.idgrupo = '" & mskCadastro(0) & "' where a.tipo <> 'CAT'"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsTreeview.RecordCount > 0 Then X = rsTreeview.Fields(0)
    On Error Resume Next
    Do While Not rsTreeview.EOF
        TreeView1.Nodes.Add , , "no" & X, rsTreeview.Fields(3)
        If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
        Do While rsTreeview.Fields(0) = X And Not rsTreeview.EOF
            TreeView1.Nodes.Add "no" & X, tvwChild, , rsTreeview.Fields(3) & " (" & rsTreeview.Fields(6) & "/" & rsTreeview.Fields(7) & "/" & rsTreeview.Fields(8) & "/" & rsTreeview.Fields(9) & "/" & rsTreeview.Fields(10) & "/" & rsTreeview.Fields(11) & ")"
            If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
        Loop
        If Not rsTreeview.EOF Then X = rsTreeview.Fields(0)
    Loop
    rsTreeview.Close
    Set rsTreeview = Nothing
End Sub

Private Sub TreeView1_Click()
    AlteraTreeview
    IncluiTreeview
End Sub

Private Sub AlteraTreeview()
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 0 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 50)
            ElseIf Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 1 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 50)
            ElseIf Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 2 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + InStr(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 100), "\") + 1, 50)
            End If
        End If
    Next
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As MSComctlLib.Node
    Set aux = Node.Child
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        If Not aux.Child Is Nothing Then
            TreeView1_NodeCheck aux
        End If
        Set aux = aux.Next
    Loop
    Set aux = Node.Parent
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        Set aux = aux.Parent
    Loop
End Sub

Private Sub configControles()
    If vSal = "N" Then
        chameleonButton12.UseGreyscale = True
        chameleonButton12.DragMode = 1
        chameleonButton12.SpecialEffect = cbEngraved
    End If
End Sub

Private Sub IncluiTreeview()
    On Error GoTo Err
    Dim rsAchaSelecao As New ADODB.Recordset
    Dim SqlAchaSelecao As String
    Dim vTipo As String, vIDSub As String
    Dim vTam As Integer, X As Integer
    
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    
10  cnBanco.BeginTrans
    
    Dim RECEBE As String
    Dim Contador As Integer
    Contador = 0
    For X = 1 To Len(SkinLabel13)
        If Mid(SkinLabel13, X, 1) = "(" Then
            Exit For
        Else
            RECEBE = RECEBE & Mid(SkinLabel13, X, 1)
        End If
    Next
    SkinLabel13 = RECEBE
    
    SqlAchaSelecao = "Select * from tbConfGrupo as a where a.idgrupo = '" & Val(Me.mskCadastro(0)) & "' and a.nome = '" & SkinLabel13 & "' and a.codcoligada = '" & Val(vCodcoligada) & "'"
    rsAchaSelecao.Open SqlAchaSelecao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaSelecao.RecordCount > 0 Then
        vTipo = rsAchaSelecao.Fields(3)
        vIDSub = rsAchaSelecao.Fields(2)
    Else
        cnBanco.RollbackTrans
        Exit Sub
    End If
    
    If vTipo = "TAB" Then
        vTam = 2 'Mid$(vIDSub, 1, 2)
    ElseIf vTipo = "CAT" Then
        vTam = 5 'Mid$(vIDSub, 1, 5)
    ElseIf vTipo = "BUT" Then
        vTam = 7 'Mid$(vIDSub, 1, 7)
    End If
    
    If chkGravar(0).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set incluir = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set incluir = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    If chkGravar(1).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set editar = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set editar = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    If chkGravar(2).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set salvar = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set salvar = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    If chkGravar(3).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set excluir = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set excluir = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    If chkGravar(4).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set imprimir = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set imprimir = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    If chkGravar(5).Value = 1 Then
        SqlMenu = "Update tbConfGrupo set filtrar = 'S' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    Else
        SqlMenu = "Update tbConfGrupo set filtrar = 'N' where idgrupo = '" & Val(Me.mskCadastro(0)) & "' and codcoligada = '" & Val(vCodcoligada) & "' and substring(idsub,1," & vTam & ") = '" & vIDSub & "'"
    End If
    rsMenu.Open SqlMenu, cnBanco
    cnBanco.CommitTrans
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

