VERSION 5.00
Begin VB.Form frmPesquisar 
   BorderStyle     =   0  'None
   Caption         =   "Exemplo de Consulta usando o ListView"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   14895
   Icon            =   "frmpesquisar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Informações "
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.PictureBox picBg 
         Height          =   495
         Left            =   13680
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   936.333
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCliForJ As New ADODB.Recordset
Private rsCliForF As New ADODB.Recordset
Private rsCliFor As New ADODB.Recordset

Private Sub cmdconsulta_Click(Index As Integer)
    Dim Y As Integer, X As Integer
    Select Case Index
    Case 0
        ListView1.ListItems(1).Selected = True
        ListView1.ListItems(1).EnsureVisible
        ListView1.SetFocus
    Case 1
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X > 1 Then
            ListView1.ListItems(X - 1).Selected = True
            ListView1.ListItems(X - 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 2
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X < Y Then
            ListView1.ListItems(X + 1).Selected = True
            ListView1.ListItems(X + 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 3
        Y = ListView1.ListItems.Count
        ListView1.ListItems(Y).Selected = True
        ListView1.ListItems(Y).EnsureVisible
        ListView1.SetFocus
    Case 4
        Pesquisa = "novo"
        frmClientes.Show 1
        ListView1.ListItems.Clear
        Form_Load
    Case 5
        Pesquisa = "editar"
        AlteraListview
        frmClientes.Show 1
        ListView1.ListItems.Clear
        Form_Load
    Case 6
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        varGlobal = ListView1.ListItems.Item(X)
        Tipo = ListView1.SelectedItem.ListSubItems.Item(7)
        Pesquisa = "excluir"
        ExcluirListview
        ListView1.ListItems.Clear
        Form_Load
    Case 7
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Frame1.Caption = "Cadastro de Hóspedes"
    frmPesquisar.Top = 1440
    frmPesquisar.Left = 110
    AbrirTabelas
    listview_cabecalho 'Chama a Sub que monta o cabeçalho das colunas do Listview
    Combo1.Text = "Nome" 'Inicializa o combo com a palavra "Codigo"
    Compoe_Listview 'Chama a Sub q lista os dados no Listview
    IniciaBarra
    FecharTabelas
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Codigo", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Endereço", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "CEP", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Bairro", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Cidade", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "UF", ListView1.Width / 20
    ListView1.ColumnHeaders.Add , , "F/J", ListView1.Width / 20
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub Compoe_Listview()
    ' Declaração de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim sql As String ' Variavel q recebe a query de conexão com a tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    
    sql = "select * from tbclifor"
    rsListview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear 'Limpa o listview
    
    'O loop abaixo se posiciona no primeiro registro da tabela Orders
    'preenche as colunas do Listview com os campos corespondentes na tabela
    'vai para o próximo registro e realiza o procedimento novamente ate chegar ao último registro
    While Not rsListview.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsListview(0), "000000"))
        ItemLst.SubItems(1) = "" & rsListview.Fields(11)
        ItemLst.SubItems(2) = "" & rsListview.Fields(1)
        ItemLst.SubItems(3) = "" & rsListview.Fields(2)
        ItemLst.SubItems(4) = "" & rsListview.Fields(3)
        ItemLst.SubItems(5) = "" & rsListview.Fields(4)
        ItemLst.SubItems(6) = "" & rsListview.Fields(5)
        If rsListview.Fields(10) = 1 Then
            ItemLst.SubItems(7) = "J"
        Else
            ItemLst.SubItems(7) = "F"
        End If
        rsListview.MoveNext
    Wend
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
    'Fecha a conexao com a tabela Orders e limpa a memória
    rsListview.Close
    Set rsListview = Nothing
End Sub

'As duas Subs abaixo faz com que ordene o listview pela coluna que vc clicar
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView1
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

Private Sub ListView1_DblClick()
    Pesquisa = "editar"
    AlteraListview
    frmClientes.Show 1
    ListView1.ListItems.Clear
    Form_Load
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        Pesquisar ' Sub que realiza a Pesquisa no Listview mediante ao que foi digitado no TexBox Text1 e ao q foi selecionado no ComboBox Combo1
    End If
End Sub

Private Sub Pesquisar()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count 'Conta as linhas preenchidas do Listview
    If Y > 0 Then 'Entra nessa condição se o Listview não estiver vazio
        'Nesse caso o "X" vai trabalhar como contador e
        'também será utilizado para percorrer as linhas do listview
        'começando de 1 até o numero de linha preenchidas no Listview
        
        '----------------------------
        picBg.Width = ListView1.Width
        picBg.Height = ListView1.ListItems(1).Height * (ListView1.ListItems.Count)
        picBg.ScaleHeight = ListView1.ListItems.Count
        picBg.ScaleWidth = 1
        picBg.DrawWidth = 1
        picBg.Cls
        '----------------------------
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True 'Seleciona a linha de acordo com o valor de "X"
            'Os procedimentos abaixo serão realizados de acordo com o q for selecionado no ComboBox Combo1
            If Combo1.Text = "Codigo" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Codigo" em todo Listview
                If UCase(ListView1.ListItems.Item(X)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
         
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Nome" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Nome" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(1)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Endereço" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Estabelecimento" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(2)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "CEP" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Endereço" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(3)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Bairro" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Cidade" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(4)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Cidade" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "UF" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(5)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "UF" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "País" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(6)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    ListView1.Picture = picBg.Image
                    'ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "" Then
                'Se não for selecionado nada no ComboBox Combo1
                MsgBox "Nenhum filtro de pesquisa selecionado"
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub AbrirTabelas()
    Sqlp = "Select * from tbcliFor Order by codclifor"
    rsCliFor.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
    
    Sqlpj = "Select * from tbcliFor, tbjuridica where tbjuridica.codclifor = tbclifor.codclifor order by tbclifor.codclifor"
    rsCliForJ.Open Sqlpj, cnBanco, adOpenKeyset, adLockOptimistic
    
    Sqlpf = "Select * from tbcliFor, tbfisica where tbfisica.codclifor = tbclifor.codclifor order by tbclifor.codclifor"
    rsCliForF.Open Sqlpf, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharTabelas()
    rsCliFor.Close
    Set rsCliFor = Nothing
            
    rsCliForJ.Close
    Set rsCliForJ = Nothing
    
    rsCliForF.Close
    Set rsCliForF = Nothing
End Sub

Private Sub ExcluirListview()
On Error GoTo TrataErro
    Dim SqlGCF As String
    Dim SqlGpj As String
    Dim SqlGpf As String
    
    Dim rsGCF As New ADODB.Recordset
    Dim rsGpf As New ADODB.Recordset
    Dim rsGpj As New ADODB.Recordset
'    Cnbanco.BeginTrans ' Inicia a transação
    If MsgBox("Confirma Exclusão", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        SqlGCF = "Delete from tbclifor where tbclifor.codclifor= " & Val(varGlobal) 'Me.txtcadastro(0))
        rsGCF.Open SqlGCF, cnBanco
        If Tipo = "J" Then
            SqlGpj = "Delete from tbjuridica where tbjuridica.codclifor= " & Val(varGlobal) 'Me.txtcadastro(0))
            rsGpj.Open SqlGpj, cnBanco
        ElseIf Tipo = "F" Then
            SqlGpf = "Delete from tbfisica where tbfisica.codclifor= " & Val(varGlobal) 'Me.txtcadastro(9))
            rsGpf.Open SqlGpf, cnBanco
        End If
'       Cnbanco.CommitTrans
        MsgBox "Registro excluido com sucesso", vbInformation, "Ok!"
    End If
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub AlteraListview()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    varGlobal = ListView1.ListItems.Item(X)
End Sub

Private Sub IniciaBarra()
    '-------------------------
    'Incializa o estilo do PictureBox
    '------------------------
    picBg.BackColor = ListView1.BackColor
    picBg.ScaleMode = vbTwips
    picBg.BorderStyle = vbBSNone
    picBg.AutoRedraw = True
    picBg.Visible = False
End Sub
