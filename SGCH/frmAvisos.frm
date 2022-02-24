VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAvisos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alertas"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmAvisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nome da Consulta"
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   480
         Top             =   1560
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
               Picture         =   "frmAvisos.frx":0CCA
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAvisos.frx":16DC
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4895
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4895
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   255
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abaixo alguns alertas a serem considerados:"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   3195
      Left            =   0
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   5636
      Image           =   "frmAvisos.frx":20EE
      Props           =   5
   End
End
Attribute VB_Name = "frmAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGenerico As New ADODB.Recordset
Private SqlGenerico As String

Private Sub Form_Load()
    If apontaLV = 0 Or apontaLV = 1 Or apontaLV = 16 Then
        Frame1.Visible = True
        ListView2.Visible = True
        frmAvisos.Caption = "MATRIZ"
        If Campo4 = 1 Then
            Frame1.Caption = "Experiência exigida pela MATRIZ"
            lv_cabGenerico "Código", "Nome do cargo", "Tempo de experiência", ""
            If apontaLV = 0 Or apontaLV = 1 Then SqlGenerico = "select a.codcargo,b.nomecargo,a.tmpoexp,'' from tbMatrizExp as a inner join tbcargos as b on a.codcoligada = '" & vCodcoligada & "' and a.codcargo = b.codcargo where codmatriz = " & Val(Mid(chamaForm.txtCadMatriz(4), 1, 6))
            If apontaLV = 16 Then SqlGenerico = "select a.codcargo,b.nomecargo,a.tmpoexp from tbMatrizExp as a inner join tbcargos as b on a.codcoligada = '" & vCodcoligada & "' and a.codcargo = b.codcargo where codmatriz = " & Val(chamaForm.txtINTD(5))
            carregaGenerico
        ElseIf Campo4 = 2 Then
            Frame1.Caption = "Cursos/treinamentos exigidos pela MATRIZ"
            lv_cabGenerico "Código", "Nome do treinamento", "Nível", "Status"
            'If apontaLV = 0 Or apontaLV = 1 Then SqlGenerico = "select a.codtreinamento,b.nometreinamento,a.codnivel from tbMatrizCur as a inner join tbtreinamentos as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento where codmatriz = " & Val(Mid(chamaForm.txtCadMatriz(4), 1, 6))
            If apontaLV = 0 Or apontaLV = 1 Then
                Dim vCPF As String
                chamaForm.mskCadMatriz.PromptInclude = False
                vCPF = chamaForm.mskCadMatriz
                chamaForm.mskCadMatriz.PromptInclude = True
                SqlGenerico = "select a.codtreinamento,b.nometreinamento,a.codnivel,CASE WHEN c.codtreinamento is null THEN 'N' ELSE 'S' END Status from tbMatrizCur as a inner join tbtreinamentos as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento left join tbColaboradoresCur as c on a.codtreinamento = c.codtreinamento and c.cpf = '" & vCPF & "' where codmatriz = '" & Val(Mid(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
            End If
            
            If apontaLV = 16 Then SqlGenerico = "select a.codtreinamento,b.nometreinamento,a.codnivel from tbMatrizCur as a inner join tbtreinamentos as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento where codmatriz = " & Val(chamaForm.txtINTD(5))
            carregaGenerico
        ElseIf Campo4 = 3 Then
            Frame1.Caption = "Graduação exigida pela MATRIZ"
            lv_cabGenerico "Código", "Formação escolar", "Pontuação", ""
            If apontaLV = 0 Or apontaLV = 1 Then SqlGenerico = "select a.codescolaridade,b.nomeescolaridade,a.pontuacao,'' from tbMatrizEsc as a inner join tbEscolaridade as b  on a.codcoligada = '" & vCodcoligada & "' and a.codescolaridade = b.codescolaridade where codmatriz = " & Val(Mid(chamaForm.txtCadMatriz(4), 1, 6))
            If apontaLV = 16 Then SqlGenerico = "select a.codescolaridade,b.nomeescolaridade,a.pontuacao from tbMatrizEsc as a inner join tbEscolaridade as b  on a.codcoligada = '" & vCodcoligada & "' and a.codescolaridade = b.codescolaridade where codmatriz = " & Val(chamaForm.txtINTD(5))
            carregaGenerico
        End If
    ElseIf apontaLV = 1 Then
        Frame1.Visible = True
        ListView2.Visible = True
    Else
        listview_cabecalho
        carrregaAvisos
    End If
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    'EXPERIÊNCIAS
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "", ListView1.Width / 1.2
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub lv_cabGenerico(coluna1 As String, coluna2 As String, coluna3 As String, coluna4 As String)
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , coluna1, ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , coluna2, ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , coluna3, ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , coluna4, ListView2.Width / 6
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub carregaGenerico()
    Dim ItemLst As ListItem
    rsGenerico.Open SqlGenerico, cnBanco, adOpenKeyset, adLockReadOnly
    While Not rsGenerico.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsGenerico.Fields(0), "000000"))
        ItemLst.SubItems(1) = rsGenerico.Fields(1)
        ItemLst.SubItems(2) = rsGenerico.Fields(2)
        If apontaLV = 0 Or apontaLV = 1 Then
            ItemLst.SubItems(3) = rsGenerico.Fields(3)
            If ItemLst.SubItems(3) = "S" Then
                ItemLst.SubItems(3) = ""
                ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
            ElseIf ItemLst.SubItems(3) = "N" Then
                ItemLst.SubItems(3) = ""
                ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
            End If
        End If
        rsGenerico.MoveNext
    Wend
    rsGenerico.Close
End Sub

Private Sub carrregaAvisos()
    Dim rsColaboradores As New ADODB.Recordset
    Dim SqlColaboradores As String
    Dim ItemLst As ListItem
   
    'COLABORADORES ABAIXO DA MEDIA
    SqlColaboradores = "Select Count (*) from tbcolaboradores as a where a.ativo = 'S' and a.mediageral < '" & vAprovadoRest & "'and a.codcoligada = '" & vCodcoligada & "'"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "Colaboradores abaixo da média"
    End If
    rsColaboradores.Close
    
    'COLABORADORES INATIVOS
    SqlColaboradores = "Select Count (*) from tbcolaboradores as a where a.codcoligada = '" & vCodcoligada & "' and a.ativo is null and a.homologacaonum is null or a.codcoligada = '" & vCodcoligada & "' and a.ativo='N' and a.homologacaonum is null"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "Colaboradores inativos"
    End If
    rsColaboradores.Close
    
    'PROCESSO SELETIVO EXPIRADO
    SqlColaboradores = "Select count(*) from tbprocessos as a where getdate() > a.datafim and status = 'Aberto' and a.codcoligada = '" & vCodcoligada & "'"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "Processo(s) Seletivo(s) expirado(s)"
    End If
    rsColaboradores.Close
    
    'INTD EXPIRADA
    SqlColaboradores = "select count(*) from tbintd as a inner join tbcolaboradores as b on a.codcolaborador = b.id and b.ativo = 'S' where getdate() > a.datafim and status = 'Aberto' and a.codcoligada = '" & vCodcoligada & "'"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "INTD(s) expirada(s)"
    End If
    rsColaboradores.Close
    
    'PROGRAMAÇÕES PENDENTES
    SqlColaboradores = "Select count(*) from tbpendentescur as a inner join tbcolaboradores as b on a.cpf = b.cpf and a.codcoligada = '" & vCodcoligada & "' inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento where a.status = 'Pendente' and a.ativo = 'S' and b.ativo = 'S'"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "Programações Pendentes"
    End If
    rsColaboradores.Close
    
    'PROGRAMAÇÕES AGENDADAS EXPIRADAS
    SqlColaboradores = "select count (*) from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz " & _
    "inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on " & _
    "a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao inner join tbUsuMultiplic as h " & _
    "on a.codtreinamento = h.codtreinamento where b.ativo = 'S' and a.ativo = 'S' and a.status='Agendado' and h.codusuario = '" & CodUsu & "' and getdate() > g.avaldata or b.ativo = 'S' and a.ativo = 'S' and " & _
    "a.status='Agendado' and g.avaldata is null and h.codusuario = '" & CodUsu & "' and getdate() > g.avaldata or b.ativo = 'S' and a.ativo = 'S' and a.status='Reagendado' and h.codusuario = '" & CodUsu & "' and getdate() > g.avaldata or b.ativo = 'S' and " & _
    "a.ativo = 'S' and a.status='Reagendado' and g.avaldata is null and h.codusuario = '" & CodUsu & "' and getdate() > g.avaldata"

    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "Programações Agendadas EXPIRADAS"
    End If
    rsColaboradores.Close
   
    'ADP EXPIRADA
    SqlColaboradores = "Select count(*) from tbListaADP as a where getdate() > a.datavencimento"
    rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockOptimistic
    If rsColaboradores.Fields(0) > 0 Then
        Set ItemLst = ListView1.ListItems.Add(, , rsColaboradores.Fields(0))
        ItemLst.SubItems(1) = "ADP(s) - Avaliação de Desempenho Profissional não avaliada(s)"
    End If
    rsColaboradores.Close
    Set rsColaboradores = Nothing
    If ListView1.ListItems.Count = 0 Then Unload Me
End Sub
