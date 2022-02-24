VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAvisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alertas"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   Icon            =   "frmAvisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Visualizei"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inconsistências"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9855
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4895
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
   End
   Begin IMRM.chameleonButton cmdFiltro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3240
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAvisos.frx":20EE
      PICN            =   "frmAvisos.frx":210A
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
Attribute VB_Name = "frmAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGenerico As New ADODB.Recordset
Private SqlGenerico As String

Private Sub cmdFiltro_Click(Index As Integer)
    'Principal.Caption = Principal.Caption & " - Estoque: " & Combo1.Text
    'vLocalEstoque = Mid$(Combo1.Text, 1, 4)
    If Check1.Value = 1 Then
        marcaVisualizado
    End If
    Unload Me
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        Principal.Caption = Principal.Caption & " - Estoque: " & Combo1.Text
        vLocalEstoque = Mid$(Combo1.Text, 1, 4)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Frame1.Visible = True
    ListView2.Visible = True
    frmAvisos.Caption = "AVISOS"
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    
    lv_cabGenerico "ID Medição", "Mov Totvs", "Status Medição", "Observação"
    If carrregaAvisos = False Then Unload Me
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub lv_cabGenerico(coluna1 As String, coluna2 As String, coluna3 As String, coluna4 As String)
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , coluna1, ListView2.Width / 4
    ListView2.ColumnHeaders.Add , , coluna2, ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , coluna3, ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , coluna4, ListView2.Width / 2.5
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Function carregaGenerico()
    Dim ItemLst As ListItem
    rsGenerico.Open SqlGenerico, cnBanco, adOpenKeyset, adLockReadOnly
    carregaGenerico = False
    While Not rsGenerico.EOF
        Set ItemLst = ListView2.ListItems.Add(, , rsGenerico.Fields(0))
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
        carregaGenerico = True
        rsGenerico.MoveNext
    Wend
    rsGenerico.Close
    
    Set rsGenerico = Nothing
End Function

Private Function carrregaAvisos()
    On Error Resume Next
    Dim ItemLst As ListItem
    Dim rsMedPJ As New ADODB.Recordset
    Dim SqlMedPJ As String
   
    Dim rsGravaAviso As New ADODB.Recordset
    Dim SqlGravaAviso As String
   
   
    carrregaAvisos = False
   
    'MEDICOES DE PJ - INCONSISTENCIA
    SqlMedPJ = "select C.codigo,c.idmovintegracao, case when b.idstatus = 2 then 'Reprovado' else '-' end as status from " & vBancoSAP & ".dbo.ID_APROP_MEDICAO as a left join " & vBancoSAP & ".dbo.ID_APROP_APROVACAO as b on a.IDAPROVACAO = b.ID LEFT JOIN tbMedicoesPJ AS C ON A.ID = C.codigo left join tbavisos as d on convert(varchar(20),a.id) = d.idmedicao where IDSTATUS = 2 AND C.idmovintegracao IS NOT NULL and d.status is null"
    rsMedPJ.Open SqlMedPJ, cnBanco, adOpenKeyset, adLockReadOnly
    
    While Not rsMedPJ.EOF
        Set ItemLst = ListView2.ListItems.Add(, , rsMedPJ.Fields(0))
        ItemLst.SubItems(1) = rsMedPJ.Fields(1)
        ItemLst.SubItems(2) = rsMedPJ.Fields(2)
        ItemLst.SubItems(3) = "Medição foi reprovada. Porém já foi lançada no RM"
        
        SqlGravaAviso = "Insert into tbAvisos(IDMEDICAO) Values(" & rsMedPJ.Fields(0) & ")"
        rsGravaAviso.Open SqlGravaAviso, cnBanco
        
        rsMedPJ.MoveNext
        carrregaAvisos = True
    Wend
    rsMedPJ.Close
     
    'MEDICOES DE TERCEIROS - INCONSISTENCIA
    SqlMedPJ = "select C.codigo,c.idmovintegracao,case when b.idstatus = 2 then 'Reprovado' else '-' end as status from " & vBancoSAP & ".dbo.ID_APROP_MEDICAOTERCEIRO as a left join " & vBancoSAP & ".dbo.ID_APROP_APROVACAO as b on A.ID = b.IDMEDICAOTERCEIRO LEFT JOIN tbMedicoesTerceiro AS C ON A.CODIGO = C.codigo COLLATE SQL_Latin1_General_CP1_CI_AS left join tbavisos as d on a.CODIGO COLLATE SQL_Latin1_General_CP1_CI_AS = convert(varchar(20),d.idmedicao) where IDSTATUS = 2 AND C.idmovintegracao IS NOT NULL and d.status is null"
    rsMedPJ.Open SqlMedPJ, cnBanco, adOpenKeyset, adLockReadOnly
    
    While Not rsMedPJ.EOF
        Set ItemLst = ListView2.ListItems.Add(, , rsMedPJ.Fields(0))
        ItemLst.SubItems(1) = rsMedPJ.Fields(1)
        ItemLst.SubItems(2) = rsMedPJ.Fields(2)
        ItemLst.SubItems(3) = "Medição foi reprovada. Porém já foi lançada no RM"
        
        SqlGravaAviso = "Insert into tbAvisos(IDMEDICAO) Values('" & rsMedPJ.Fields(0) & "')"
        rsGravaAviso.Open SqlGravaAviso, cnBanco
        
        rsMedPJ.MoveNext
        carrregaAvisos = True
    Wend
    rsMedPJ.Close

End Function

Private Sub marcaVisualizado()
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer, Y As Integer
    
    Dim rsVisualizado As New ADODB.Recordset
    Dim SqlVisualizado As String
    
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        If ListView2.ListItems.Item(X).Selected = True And ListView2.ListItems.Item(X).Checked = True Then
            SqlVisualizado = "update tbAvisos set status = 1 where idmedicao = '" & ListView2.ListItems.Item(X) & "'"
            rsVisualizado.Open SqlVisualizado, cnBanco
            Set rsVisualizado = Nothing
        End If
    Next
End Sub

Private Sub ListView1_Click()
    FCRValidadeCertif.Show 1
End Sub

