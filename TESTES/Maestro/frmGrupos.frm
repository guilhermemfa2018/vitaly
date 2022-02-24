VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos"
   ClientHeight    =   8640
   ClientLeft      =   3270
   ClientTop       =   1275
   ClientWidth     =   7665
   Icon            =   "frmGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmGrupos.frx":0CCA
      TabIndex        =   22
      Top             =   720
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmGrupos.frx":0D3C
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
   Begin MAESTRO.chameleonButton chameleonButton11 
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
   Begin MAESTRO.chameleonButton chameleonButton12 
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
      Left            =   6480
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
      Width           =   7455
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
         Width           =   7215
         Begin VB.CheckBox chkGravar 
            Height          =   255
            Index           =   7
            Left            =   4200
            TabIndex        =   13
            Top             =   120
            Width           =   255
         End
         Begin VB.Frame Frame4 
            Caption         =   "      Admitir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   4080
            TabIndex        =   18
            Top             =   120
            Width           =   3015
            Begin VB.CheckBox chkGravar 
               Caption         =   "Reprovado"
               Enabled         =   0   'False
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   20
               Top             =   720
               Width           =   1335
            End
            Begin VB.CheckBox chkGravar 
               Caption         =   "Aprovado com restição"
               Enabled         =   0   'False
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   2295
            End
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Avaliar"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Filtrar"
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   16
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Imprimir"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Demitir"
            Height          =   255
            Index           =   8
            Left            =   3120
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Editar"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Incluir"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Salvar"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkGravar 
            Caption         =   "Excluir"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4980
         Left            =   150
         TabIndex        =   5
         Top             =   255
         Width           =   7155
         _ExtentX        =   12621
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
      Width           =   6255
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
    mobjMsg.Abrir "Deseja salvar os dados do Grupo?", YesNo, pergunta, "SGC"
    If Tp = 1 Then
        Bot_salvar
        gravaLog "Código grupo: " & mskCadastro(0), "Nome grupo: " & txtCadastro(0), ""
        Unload Me
    End If
End Sub

Private Sub chkGravar_Click(Index As Integer)
    If chkGravar(7).Value = 1 Then
        chkGravar(9).Enabled = True
        chkGravar(10).Enabled = True
    Else
        chkGravar(9).Value = 0
        chkGravar(10).Value = 0
        chkGravar(9).Enabled = False
        chkGravar(10).Enabled = False
    End If
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
    SqlGrupo = "select * from tbgrupo where codcoligada = '" & vCodcoligada & "' and codigo ='" & Val(varGlobal) & "'"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGrupo.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Grupo não encontrado", Ok, critico, "Atenção"
    End If
    rsGrupo.Close
    Set rsGrupo = Nothing
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
    If rsTreeview.Fields(5) = "S" Then chkGravar(0).Value = 1 Else chkGravar(0).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(1).Value = 1 Else chkGravar(1).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(2).Value = 1 Else chkGravar(2).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(3).Value = 1 Else chkGravar(3).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(4).Value = 1 Else chkGravar(4).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(5).Value = 1 Else chkGravar(5).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(6).Value = 1 Else chkGravar(6).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(7).Value = 1 Else chkGravar(7).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(8).Value = 1 Else chkGravar(8).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(9).Value = 1 Else chkGravar(9).Value = 0
    rsTreeview.MoveNext
    If rsTreeview.Fields(5) = "S" Then chkGravar(10).Value = 1 Else chkGravar(10).Value = 0
    
    rsTreeview.Close
    Set rsTreeview = Nothing
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
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
End Function

Private Sub AbrirGrupo()
    SqlGrupo = "Select * from tbGrupo where codcoligada = '" & vCodcoligada & "' Order by codigo"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharGrupo()
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub

Private Sub Bot_salvar()
'On Error GoTo TrataErro
    Dim SqlSalvar As String
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer, i As Integer
    cnBanco.BeginTrans
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
    
    If chkGravar(0).Value = 0 Then vInc = "N" Else vInc = "S"
    If chkGravar(1).Value = 0 Then vEdi = "N" Else vEdi = "S"
    If chkGravar(2).Value = 0 Then vSal = "N" Else vSal = "S"
    If chkGravar(3).Value = 0 Then vExc = "N" Else vExc = "S"
    If chkGravar(4).Value = 0 Then vImp = "N" Else vImp = "S"
    If chkGravar(5).Value = 0 Then vFil = "N" Else vFil = "S"
    If chkGravar(6).Value = 0 Then vAva = "N" Else vAva = "S"
    If chkGravar(7).Value = 0 Then vAdi = "N" Else vAdi = "S"
    If chkGravar(8).Value = 0 Then vDem = "N" Else vDem = "S"
    If chkGravar(9).Value = 0 Then vAdiRes = "N" Else vAdiRes = "S"
    If chkGravar(10).Value = 0 Then vAdiRep = "N" Else vAdiRep = "S"
    
    SqlSalvar = "Delete from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and tbConfGrupo.idgrupo = '" & Val(Me.mskCadastro(0)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    
       SqlSalvar = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'01','TAB','Cadastros','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'81','CAT','Primários','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'82','CAT','Secundários','S','" & vCodcoligada & "',0);" & _
                  "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8181','BUT','Departamentos','S','" & vCodcoligada & "',1);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8182','BUT','Setores','S','" & vCodcoligada & "',2);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8184','BUT','Cargos','S','" & vCodcoligada & "',3);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8185','BUT','Habilidades','S','" & vCodcoligada & "',4);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8186','BUT','Escolaridade','S','" & vCodcoligada & "',5);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8187','BUT','Avaliações','S','" & vCodcoligada & "',6);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8288','BUT','Cursos/Treinamentos','S','" & vCodcoligada & "',7);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8289','BUT','Matriz de Capacitação','S','" & vCodcoligada & "',8);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8290','BUT','Candidatos','S','" & vCodcoligada & "',9);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",1,'8291','BUT','Colaboradores','S','" & vCodcoligada & "',10);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'02','TAB','Recrutamento','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'01','CAT','Seleção de pessoal','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'0101','BUT','Requisição de pessoal','S','" & vCodcoligada & "',11);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",2,'0102','BUT','Processo seletivo','S','" & vCodcoligada & "',12);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'03','TAB','Decisão','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'11','CAT','Decisões gerenciais','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",3,'1111','BUT','PDO','S','" & vCodcoligada & "',13);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'04','TAB','Capacitação','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'21','CAT','Capacitação de pessoal','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'2121','BUT','Programação','S','" & vCodcoligada & "',14);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'2122','BUT','Restrições','S','" & vCodcoligada & "',15);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'2123','BUT','INTD','S','" & vCodcoligada & "',16);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",4,'2124','BUT','ADP','S','" & vCodcoligada & "',17);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'05','TAB','Relatórios','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'31','CAT','Relatórios','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'3131','BUT','Gráficos de Competências','S','" & vCodcoligada & "',26);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'3132','BUT','Programação anual de treinamentos','S','" & vCodcoligada & "',27);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'3133','BUT','Relação de cargos por treinamento','S','" & vCodcoligada & "',28);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",5,'3134','BUT','Rel-04','S','" & vCodcoligada & "',26);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'06','TAB','Configurações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'51','CAT','Parametrizações','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'52','CAT','Aparência','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5151','BUT','Sistema','S','" & vCodcoligada & "',18);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5152','BUT','Grupos','S','" & vCodcoligada & "',19);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5153','BUT','Usuários','S','" & vCodcoligada & "',20);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5254','BUT','Menu','S','" & vCodcoligada & "',21);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5255','BUT','Skin','S','" & vCodcoligada & "',22);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",6,'5256','BUT','Fundo','S','" & vCodcoligada & "',23);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'07','TAB','Sobre','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'61','CAT','Sobre','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'6161','BUT','Sobre SGC','S','" & vCodcoligada & "',24);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & Val(Me.mskCadastro(0)) & ",7,'6162','BUT','Ajuda do SGC','S','" & vCodcoligada & "',25);"
    rsSalvar.Open SqlSalvar, cnBanco
    
    SqlSalvar = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKINC','" & vInc & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKEDI','" & vEdi & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKSAL','" & vSal & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKEXC','" & vExc & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKIMP','" & vImp & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKFIL','" & vFil & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKAVA','" & vAva & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKADI','" & vAdi & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKDEM','" & vDem & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKADIRES','" & vAdiRes & "'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(" & Val(Me.mskCadastro(0)) & ",0,0,'CHK','CHKADIREP','" & vAdiRep & "'," & vCodcoligada & ");"
    rsSalvar.Open SqlSalvar, cnBanco

    SqlSalvar = "Select * from tbConfGrupo where codcoligada = '" & vCodcoligada & "' and tbConfGrupo.idmenu <> 0 and tbConfGrupo.idgrupo = '" & Val(Me.mskCadastro(0)) & "' and tbConfGrupo.tipo <> 'CAT' order by id"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
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
    mobjMsg.Abrir "Dados gravados com sucesso!", Ok, informacao, "SGC"
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(mskCadastro(0), "000000"))
        ItemLst.SubItems(1) = txtCadastro(0).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(0).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ReportIcon = "OK"
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
    SqlTreeview = "Select * from tbMenu where tipo <> 'CAT' and codcoligada = '" & vCodcoligada & "'"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    
    X = rsTreeview.Fields(0)
    On Error Resume Next
    Do While Not rsTreeview.EOF
        TreeView1.Nodes.Add , , "no" & X, rsTreeview.Fields(3)
        If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
        Do While rsTreeview.Fields(0) = X And Not rsTreeview.EOF
            TreeView1.Nodes.Add "no" & X, tvwChild, , rsTreeview.Fields(3)
            If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
        Loop
        If Not rsTreeview.EOF Then X = rsTreeview.Fields(0)
    Loop
    rsTreeview.Close
    Set rsTreeview = Nothing
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
