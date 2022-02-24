VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrintRels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintRels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Semana/Ano"
      Height          =   855
      Left            =   960
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   4200
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   1440
      OleObjectBlob   =   "frmPrintRels.frx":3469A
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "FCE"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informe o período "
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   114688001
         CurrentDate     =   41660
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   114688001
         CurrentDate     =   41660
      End
   End
   Begin ZEUS.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   2160
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
      MICON           =   "frmPrintRels.frx":34702
      PICN            =   "frmPrintRels.frx":3471E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2160
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
      MICON           =   "frmPrintRels.frx":353F8
      PICN            =   "frmPrintRels.frx":35414
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o relatório que deseja visualizar "
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmPrintRels.frx":360EE
         Left            =   120
         List            =   "frmPrintRels.frx":36101
         TabIndex        =   0
         Text            =   "Ponto X Apropriação"
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPrintRels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CaminhoArquivo As String
Private NomeArquivo As String
Private pathArq As String
Private Plan As Object 'Aplicação Excel
Private vStatusOperacao As Integer
Private rsApropriacao As New ADODB.Recordset
Private SqlApropriacao As String
Private vProgress As Integer
Private vGuardaLegenda As String

Private Sub cmdImprimir_Click(Index As Integer)
'On Error Resume Next
    Select Case Index
    Case 0
        If apontaLV = 9 Then
            If Combo1.ListIndex = 0 Then
                FCRConfronto.Show 1
            ElseIf Combo1.ListIndex = 1 Then
                SalvaXLS 1 'Plano de Carga
            ElseIf Combo1.ListIndex = 2 Then
                vDataFilter1 = DTPicker1.Value
                vDataFilter2 = DTPicker2.Value
                
                
                
                FCRApropriacao.Show 1
            ElseIf Combo1.ListIndex = 3 Then
                'Evolução de Fabricação
                If Text1.Text <> "" Then
                    SalvaXLS 2
                Else
                    mobjMsg.Abrir "Favor informar o nº da FCE", Ok, critico, "Atenção"
                End If
            ElseIf Combo1.ListIndex = 4 Then
                'preparaParada
                vGuardaLegenda = Principal.StatusBar1.Panels(3).Text
                preparaHA
            End If
        ElseIf apontaLV = 12 Then
            If Combo1.ListIndex = 0 Then
                'Evolução de Fabricação
                If Text1.Text <> "" Then
                    SalvaXLS 2
                Else
                    mobjMsg.Abrir "Favor informar o nº da FCE", Ok, critico, "Atenção"
                End If
            End If
        ElseIf apontaLV = 19 Then
            FCRLibparaInsp.Show 1
        ElseIf apontaLV = 20 Then
            If Not IsNull(DTPicker1.Value) And IsNull(DTPicker2.Value) Then
                mobjMsg.Abrir "Favor informar a 2ª data do período", Ok, critico, "Atenção"
                Exit Sub
            End If
            If Combo1.ListIndex = 0 Then
                FCRFatFCE.Show 1
            ElseIf Combo1.ListIndex = 1 Then
                FCRFatFCESint.Show 1
            ElseIf Combo1.ListIndex = 2 Then
                vDataFilter1 = DTPicker1.Value
                vDataFilter2 = DTPicker2.Value
                FCRFatPeriodo.Show 1
            End If
        End If
    Case 1
        Unload Me
        Set frmPrintRels = Nothing
    End Select
End Sub

Private Sub Combo1_Click()
    If apontaLV = 9 Then
        If Combo1.ListIndex = 1 Then
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
        ElseIf Combo1.ListIndex = 2 Then
            DTPicker1.Value = ""
            DTPicker2.Value = ""
            If Combo1.ListIndex = 1 Then
                Frame2.Visible = True
                Frame3.Visible = False
                Frame4.Visible = False
            ElseIf Combo1.ListIndex = 2 Then
                Frame2.Visible = True
                Frame3.Visible = False
                Frame4.Visible = False
            End If
        ElseIf Combo1.ListIndex = 3 Then
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
        ElseIf Combo1.ListIndex = 4 Then
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = True
        Else
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
        End If
    ElseIf apontaLV = 12 Then
        If Combo1.ListIndex = 0 Then
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
        End If
    ElseIf apontaLV = 20 Then
        DTPicker1.Value = ""
        DTPicker2.Value = ""
        If Combo1.ListIndex = 1 Then
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
        ElseIf Combo1.ListIndex = 2 Then
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
        End If
    End If
End Sub

Private Sub Combo1_LostFocus()
    If apontaLV = 9 Then
        If Combo1.ListIndex = 1 Then
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
        ElseIf Combo1.ListIndex = 2 Then
            DTPicker1.Value = ""
            DTPicker2.Value = ""
            If Combo1.ListIndex = 1 Then
                Frame2.Visible = True
                Frame3.Visible = False
                Frame4.Visible = False
            ElseIf Combo1.ListIndex = 2 Then
                Frame2.Visible = True
                Frame3.Visible = False
                Frame4.Visible = False
            End If
        ElseIf Combo1.ListIndex = 3 Then
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
        ElseIf Combo1.ListIndex = 4 Then
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = True
        Else
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    If apontaLV = 9 Then
        Combo1.Clear
        Combo1.AddItem "Ponto X Apropriação"
        Combo1.AddItem "Plano de Carga"
        Combo1.AddItem "Apropriação"
        Combo1.AddItem "Evolução de Fabricação"
        Combo1.AddItem "ROP"
    ElseIf apontaLV = 12 Then
        Combo1.Clear
        Combo1.AddItem "Evolução de Fabricação"
    ElseIf apontaLV = 19 Then
        Combo1.Clear
        Combo1.AddItem "Itens Liberados para Inspeção"
    ElseIf apontaLV = 20 Then
        Combo1.Clear
        Combo1.AddItem "Faturamento Gerencial (Analítico)"
        Combo1.AddItem "Faturamento Gerencial (Sintético)"
        Combo1.AddItem "Faturamento Por Período"
    End If
    Combo1.ListIndex = 0
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub SalvaXLS(Indice As Integer)
On Error GoTo testa_erro
    'If Text5.Text = "" Then
    '    Msgbox "Os dados do orçamento devem ser informados"
    '    Exit Sub
    'End If
    CaminhoArquivo = ""
    NomeArquivo = ""
    CaminhoArquivo = pathArq 'Mid$(frmConfiguracao.txtCaminho, 1, Len(frmConfiguracao.txtCaminho) - Len("contratoNOVO.mdb"))
    If Indice = 1 Then
        NomeArquivo = "Plano de Carga.xls"
    ElseIf Indice = 2 Then
        NomeArquivo = "Fabricacao.xls"
    End If
    
    cdg.Filter = "Planilha do Excel (*.xls)|*.xls"
    cdg.flags = cdlOFNHideReadOnly
    cdg.InitDir = CaminhoArquivo
    cdg.FileName = NomeArquivo
    pathArq = cdg.FileName
    cdg.ShowSave
    If Trim(pathArq) <> "" Then
        If Indice = 1 Then
            ExportaExcelCarga 'Plano de Carga
        ElseIf Indice = 2 Then
            ExportaExcelEvolucao
        End If
    End If
    Exit Sub
testa_erro:
    If Err.Number = 32755 Then
        mobjMsg.Abrir "Procedimento cancelado", Ok, critico, "Atenção"
    End If
End Sub

Private Sub ExportaExcelCarga()
'On Error Resume Next
    Dim SommaCC As Double
    Dim vTCNC1 As String, vTCNC2 As String, vTGuil As String, vTTPuns As String, vTRosq As String, vTFRadial As String, vTFPrisma As String, vTFMag As String, vTSerraFita As String, vTCorte As String, vTDesemp As String, vTPrensa As String, vTMonC As String, vTMonN As String, vTSolC As String, vTSolN As String, vTAcabC As String, vTAcabN As String, vTCal As String, vTTrac As String, vTQua As String
    
    Dim j As Integer, K As Integer, L As Integer
    
    'Dados das OSs que estão dentro do intervalo de tempo informado
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    Dim vOS As Integer
    Dim vRevisao As Integer, vsemana As Integer
    
    SkinLabel1.Visible = True
    mobjMsg.Abrir "Esse procedimento pode demorar alguns minutos.", Ok, critico, "Atenção"
    
''    SqlOS = "select B.idprogramacao,B.idos,B.idoperacao,B.idcc,'Horas' = dbo.FN_CONVMIN((cast(replace(B.tempocalc,'.','') as money)/100)),DATEPART(WK,B.dataprevista) as Semana,d.desenho,f.fce,f.projeto " & _
''            "from tbmpitens as B INNER JOIN tbMP AS E ON B.idprogramacao = E.idprogramacao INNER JOIN tbProjetos AS F ON E.codprojeto = F.codprojeto left join tbitemlm as c on SUBSTRING(b.desenhos,1,2) = c.codlm and " & _
''            "replace(SUBSTRING(b.desenhos,3,4),';','') = c.codseq and F.fce = C.fce left join tbDesenhos as d on c.codigodes = d.iddesenho " & _
''            "where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' order by B.dataprevista,B.idos,B.idcc"

'    SqlOS = "select a.idprogramacao,b.idos,b.idoperacao,B.idcc,'Horas' = dbo.FN_CONVMIN((cast(replace(B.tempocalc,'.','') as money)/100)),DATEPART(WK,B.dataprevista) as Semana,a.desenho,f.fce,f.projeto,b.revisaoos " & _
'            "from tbmp as a inner join tbmpitens as b on a.idprogramacao = b.idprogramacao INNER JOIN tbProjetos AS F ON a.codprojeto = F.codprojeto where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' order by B.dataprevista,B.idos,b.revisaoos,b.idoperacao,B.idcc"

'LISTA TODAS AS OSs E SUAS OPERAÇÕES NO PERIODO INFORMADO
    SqlOS = "Set datefirst 1 select a.idprogramacao,b.idos,max(b.idoperacao),B.idcc,'Horas' = dbo.FN_CONVMIN(sum(((cast(replace(b.tempocalc,'.','') as money))/100))),DATEPART(WK,B.dataprevista) as Semana,a.desenho,f.fce,f.projeto,b.revisaoos,max(g.percentualBaixado),b.status " & _
            "from tbmp as a inner join tbmpitens as b on a.idprogramacao = b.idprogramacao INNER JOIN tbProjetos AS F ON a.codprojeto = F.codprojeto left join tbMPBaixaParcial as g on b.idos = g.idos and b.revisaoos = g.revisao and b.idoperacao = g.idoperacao where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' group by " & _
            "a.idprogramacao,b.idos,B.idcc,B.dataprevista,a.desenho,f.fce,f.projeto,b.revisaoos,b.status order by B.dataprevista,B.idos,b.revisaoos,B.idcc"


'ABAIXO: QUERY PARA REALIZACAO DE TESTES EM OSs PROBLEMATICAS 23/12/2019 (EH A MESMA QUERY ACIMA SO QUE COM OSs ESPECIFICAS)
'    SqlOS = "Set datefirst 1 select a.idprogramacao,b.idos,max(b.idoperacao),B.idcc,'Horas' = dbo.FN_CONVMIN(sum(((cast(replace(b.tempocalc,'.','') as money))/100))),DATEPART(WK,B.dataprevista) as Semana,a.desenho,f.fce,f.projeto,b.revisaoos,max(g.percentualBaixado),b.status " & _
'            "from tbmp as a inner join tbmpitens as b on a.idprogramacao = b.idprogramacao INNER JOIN tbProjetos AS F ON a.codprojeto = F.codprojeto left join tbMPBaixaParcial as g on b.idos = g.idos and b.revisaoos = g.revisao and b.idoperacao = g.idoperacao " & _
'            "where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' AND " & _
'            "B.idos in(2916,2936,3602,3785,3826,3845,3671,3777,3779,3947)  " & _
'            "group by " & _
'            "a.idprogramacao,b.idos,B.idcc,B.dataprevista,a.desenho,f.fce,f.projeto,b.revisaoos,b.status order by B.dataprevista,B.idos,b.revisaoos,B.idcc"






    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsOS.Open SqlOS, cnBanco, adOpenKeyset, adLockReadOnly
    
    
    'Dim Plan As Object 'Aplicação Excel
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")

    'PLANILHA DE LISTA DE MATERIAIS
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\PLANO DE CARGA.xls"
    Plan.Visible = True
    Plan.UserControl = False

    Plan.ScreenUpdating = False
    Plan.EnableEvents = False
    Plan.Calculation = xlManual
    
'----------------------------------
    Dim vAcumulaData1() As String
    Dim vVinteQuatroHoras() As String
    Dim vAcumulaData2(3) As Integer
    Dim F As Integer
    Dim vText As Date
    Dim vText2 As String
    vText = "23:59"
'----------------------------------
    
    j = 7
    X = 1
    'Dim valor1 As Double, valor2 As Double, valor3 As Double, valor4 As Double, valor5 As Double, QtdTotCJ As Double
    
    With Plan
            .Range("F3").Value = DTPicker1.Value
            .Range("H3").Value = DTPicker2.Value
    End With
    
    While Not rsOS.EOF
        vOS = rsOS.Fields(1)
        vRevisao = rsOS.Fields(9)
        vsemana = rsOS.Fields(5)
        With Plan
            .Range("A" & j).Value = rsOS.Fields(7) & " - " & rsOS.Fields(8) ' FCE/Projeto
            .Range("B" & j).Value = rsOS.Fields(6) ' Desenho
            .Range("C" & j).Value = rsOS.Fields(5) ' nº da Semana
            .Range("D" & j).Value = rsOS.Fields(1) ' nº da OS - Ordem de Serviço
            .Range("E" & j).Value = rsOS.Fields(9) ' nº da REVISÃO da OS - Ordem de Serviço
            
            .Range("F" & 5).Value = "3101.SC-01 (CNC1)" 'Cabeçalho
            .Range("I" & 5).Value = "3101.SC-02 (CNC2)" 'Cabeçalho
            .Range("L" & 5).Value = "3101.SC-03 (GUILH)" 'Cabeçalho
            .Range("O" & 5).Value = "3101.SC-04 (PUNS)" 'Cabeçalho
            .Range("R" & 5).Value = "3101.SC-05 (ROSQ)" 'Cabeçalho
            .Range("U" & 5).Value = "3101.SC-06 (FR)" 'Cabeçalho
            .Range("X" & 5).Value = "3101.SC-07 (FPRIS)" 'Cabeçalho
            .Range("AA" & 5).Value = "3101.SC-08 (FBM)" 'Cabeçalho
            .Range("AD" & 5).Value = "3101.SC-09 (SRF)" 'Cabeçalho
            .Range("AG" & 5).Value = "3101.SC-10 (C/R)" 'Cabeçalho
            .Range("AJ" & 5).Value = "3101.SC-12 (DC)" 'Cabeçalho
            .Range("AM" & 5).Value = "3102.SC-01 (PRE)" 'Cabeçalho
            .Range("AP" & 5).Value = "3102.SC-02 (CAL)" 'Cabeçalho
            .Range("AS" & 5).Value = "3106.SC-01 (TRAÇ)" 'Cabeçalho
            .Range("AV" & 5).Value = "3103.SC-01 (MON C)" 'Cabeçalho
            .Range("AY" & 5).Value = "3103.SC-02 (MON N)" 'Cabeçalho
            .Range("BB" & 5).Value = "3104.SC-01 (SOL C)" 'Cabeçalho
            .Range("BE" & 5).Value = "3104.SC-02 (SOL N)" 'Cabeçalho
            .Range("BH" & 5).Value = "3105.SC-01 (ACA C)" 'Cabeçalho
            .Range("BK" & 5).Value = "3105.SC-02 (ACA N)" 'Cabeçalho
            .Range("BN" & 5).Value = "7103.SC-02 (QUA)" 'Cabeçalho
        End With
        
        Do While vOS = rsOS.Fields(1) And vRevisao = rsOS.Fields(9) And vsemana = rsOS.Fields(5)
            
            'If vOS = "2748" Then Msgbox "Aqui"
            
            With Plan
                vStatusOperacao = 0 'A cada vez que muda o registro zera o status para não correr o risco de pegar residuo do status da operação anterior
'INICIO         'CNC1
                If rsOS.Fields(3) = "3000.3101.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("F" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("E" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCNC1)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("F" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm:ss"), vTCNC1)
                            Wend
                            If vRevisao = 0 Then .Range("F" & j).Value = somaTempoAcumulado(CDate(vText2), vTCNC1)
                        Else
                            If vRevisao = 0 Then .Range("F" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCNC1)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("H" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("G" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("G" & j).Value < .Range("F" & j).Value And rsOS.Fields(11) = 3 Then .Range("F" & j).Value = .Range("G" & j).Value
                End If
'INICIO         'CNC2
                If rsOS.Fields(3) = "3000.3101.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("I" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("H" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCNC2)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("I" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTCNC2)
                            Wend
                            If vRevisao = 0 Then .Range("I" & j).Value = somaTempoAcumulado(CDate(vText2), vTCNC2)
                        Else
                            If vRevisao = 0 Then .Range("I" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCNC2)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("K" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("J" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("J" & j).Value < .Range("I" & j).Value And rsOS.Fields(11) = 3 Then .Range("I" & j).Value = .Range("J" & j).Value
                End If
'INICIO         'Guilhotina
                If rsOS.Fields(3) = "3000.3101.SC-03" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("L" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("K" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTGuil)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("L" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTGuil)
                            Wend
                            If vRevisao = 0 Then .Range("L" & j).Value = somaTempoAcumulado(CDate(vText2), vTGuil)
                        Else
                            If vRevisao = 0 Then .Range("L" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTGuil)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("N" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("M" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("M" & j).Value < .Range("L" & j).Value And rsOS.Fields(11) = 3 Then .Range("L" & j).Value = .Range("M" & j).Value
                End If
'INICIO         'Tesoura Punsionadeira
                If rsOS.Fields(3) = "3000.3101.SC-04" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("O" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("N" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTTPuns)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("O" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTTPuns)
                            Wend
                            If vRevisao = 0 Then .Range("O" & j).Value = somaTempoAcumulado(CDate(vText2), vTTPuns)
                        Else
                            If vRevisao = 0 Then .Range("O" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTTPuns)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("Q" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("P" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("P" & j).Value < .Range("O" & j).Value And rsOS.Fields(11) = 3 Then .Range("O" & j).Value = .Range("P" & j).Value
                End If
'INICIO         'Rosqueadeira
                If rsOS.Fields(3) = "3000.3101.SC-05" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("R" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("Q" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTRosq)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("R" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTRosq)
                            Wend
                            If vRevisao = 0 Then .Range("R" & j).Value = somaTempoAcumulado(CDate(vText2), vTRosq)
                        Else
                            If vRevisao = 0 Then .Range("R" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTRosq)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("T" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("S" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("S" & j).Value < .Range("R" & j).Value And rsOS.Fields(11) = 3 Then .Range("R" & j).Value = .Range("S" & j).Value
                End If
'INICIO         'Furadeira Radial
                If rsOS.Fields(3) = "3000.3101.SC-06" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("U" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("T" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFRadial)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("U" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTFRadial)
                            Wend
                            If vRevisao = 0 Then .Range("U" & j).Value = somaTempoAcumulado(CDate(vText2), vTFRadial)
                        Else
                            If vRevisao = 0 Then .Range("U" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFRadial)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("W" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("V" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("V" & j).Value < .Range("U" & j).Value And rsOS.Fields(11) = 3 Then .Range("U" & j).Value = .Range("V" & j).Value
                End If
'INICIO         'Furadeira Prismática
                If rsOS.Fields(3) = "3000.3101.SC-07" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("X" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("W" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFPrisma)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("X" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTFPrisma)
                            Wend
                            If vRevisao = 0 Then .Range("X" & j).Value = somaTempoAcumulado(CDate(vText2), vTFPrisma)
                        Else
                            If vRevisao = 0 Then .Range("X" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFPrisma)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("Z" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("Z" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("Z" & j).Value < .Range("X" & j).Value And rsOS.Fields(11) = 3 Then .Range("X" & j).Value = .Range("Z" & j).Value
                End If
'INICIO         'Furadeira Base Magnética
                If rsOS.Fields(3) = "3000.3101.SC-08" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AA" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("Z" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFMag)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AA" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTFMag)
                            Wend
                            If vRevisao = 0 Then .Range("AA" & j).Value = somaTempoAcumulado(CDate(vText2), vTFMag)
                        Else
                            If vRevisao = 0 Then .Range("AA" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTFMag)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AC" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AB" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AB" & j).Value < .Range("AA" & j).Value And rsOS.Fields(11) = 3 Then .Range("AA" & j).Value = .Range("AB" & j).Value
                End If
'INICIO         'Serra Fita Franho
                If rsOS.Fields(3) = "3000.3101.SC-09" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AD" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AC" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSerraFita)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AD" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTSerraFita)
                            Wend
                            If vRevisao = 0 Then .Range("AD" & j).Value = somaTempoAcumulado(CDate(vText2), vTSerraFita)
                        Else
                            If vRevisao = 0 Then .Range("AD" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSerraFita)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AF" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AE" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AE" & j).Value < .Range("AD" & j).Value And rsOS.Fields(11) = 3 Then .Range("AD" & j).Value = .Range("AE" & j).Value
                End If
'INICIO         'Corte/Recorte
                If rsOS.Fields(3) = "3000.3101.SC-10" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AG" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AF" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCorte)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AG" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTCorte)
                            Wend
                            If vRevisao = 0 Then .Range("AG" & j).Value = somaTempoAcumulado(CDate(vText2), vTCorte)
                        Else
                            If vRevisao = 0 Then .Range("AG" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCorte)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AI" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AH" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AH" & j).Value < .Range("AG" & j).Value And rsOS.Fields(11) = 3 Then .Range("AG" & j).Value = .Range("AH" & j).Value
                End If
'INICIO         'Desempeno a Calor
                If rsOS.Fields(3) = "3000.3101.SC-12" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AJ" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AI" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTDesemp)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AJ" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTDesemp)
                            Wend
                            If vRevisao = 0 Then .Range("AJ" & j).Value = somaTempoAcumulado(CDate(vText2), vTDesemp)
                        Else
                            If vRevisao = 0 Then .Range("AJ" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTDesemp)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AL" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AK" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AK" & j).Value < .Range("AJ" & j).Value And rsOS.Fields(11) = 3 Then .Range("AJ" & j).Value = .Range("AK" & j).Value
                End If
'INICIO         'Prensa
                If rsOS.Fields(3) = "3000.3102.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AM" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AL" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTPrensa)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AM" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTPrensa)
                            Wend
                            If vRevisao = 0 Then .Range("AM" & j).Value = somaTempoAcumulado(CDate(vText2), vTPrensa)
                        Else
                            If vRevisao = 0 Then .Range("AM" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTPrensa)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AO" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AN" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AN" & j).Value < .Range("AM" & j).Value And rsOS.Fields(11) = 3 Then .Range("AM" & j).Value = .Range("AN" & j).Value
                End If
'INICIO         'Calandra
                If rsOS.Fields(3) = "3000.3102.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AP" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AO" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCal)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AP" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTCal)
                            Wend
                            If vRevisao = 0 Then .Range("AP" & j).Value = somaTempoAcumulado(CDate(vText2), vTCal)
                        Else
                            If vRevisao = 0 Then .Range("AP" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTCal)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AR" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AQ" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AQ" & j).Value < .Range("AP" & j).Value And rsOS.Fields(11) = 3 Then .Range("AP" & j).Value = .Range("AQ" & j).Value
                End If
'INICIO         'Traçagem
                If rsOS.Fields(3) = "3000.3106.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AS" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AR" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTTrac)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AS" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTTrac)
                            Wend
                            If vRevisao = 0 Then .Range("AS" & j).Value = somaTempoAcumulado(CDate(vText2), vTTrac)
                        Else
                            If vRevisao = 0 Then .Range("AS" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTTrac)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AU" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AT" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AT" & j).Value < .Range("AS" & j).Value And rsOS.Fields(11) = 3 Then .Range("AS" & j).Value = .Range("AT" & j).Value
                End If
'INICIO         'Montagem Caldeiraria
                If rsOS.Fields(3) = "3000.3103.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AV" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AU" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTMonC)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AV" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTMonC)
                            Wend
                            If vRevisao = 0 Then .Range("AV" & j).Value = somaTempoAcumulado(CDate(vText2), vTMonC)
                        Else
                            If vRevisao = 0 Then .Range("AV" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTMonC)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("AX" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AW" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AW" & j).Value < .Range("AV" & j).Value And rsOS.Fields(11) = 3 Then .Range("AV" & j).Value = .Range("AW" & j).Value
                End If
'INICIO         'Montagem Naval
                If rsOS.Fields(3) = "3000.3103.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("AY" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("AX" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTMonN)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("AY" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTMonN)
                            Wend
                            If vRevisao = 0 Then .Range("AY" & j).Value = somaTempoAcumulado(CDate(vText2), vTMonN)
                        Else
                            If vRevisao = 0 Then .Range("AY" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTMonN)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BA" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))

                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("AZ" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("AZ" & j).Value < .Range("AY" & j).Value And rsOS.Fields(11) = 3 Then .Range("AY" & j).Value = .Range("AZ" & j).Value
                End If
'INICIO         'Solda Caldeiraria
                If rsOS.Fields(3) = "3000.3104.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("BB" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("BA" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSolC)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("BB" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTSolC)
                            Wend
                            If vRevisao = 0 Then .Range("BB" & j).Value = somaTempoAcumulado(CDate(vText2), vTSolC)
                        Else
                            If vRevisao = 0 Then .Range("BB" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSolC)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BD" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("BC" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("BC" & j).Value < .Range("BB" & j).Value And rsOS.Fields(11) = 3 Then .Range("BB" & j).Value = .Range("BC" & j).Value
                End If
'INICIO         'Solda Naval
                If rsOS.Fields(3) = "3000.3104.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("BE" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("BD" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSolN)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("BE" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTSolN)
                            Wend
                            If vRevisao = 0 Then .Range("BE" & j).Value = somaTempoAcumulado(CDate(vText2), vTSolN)
                        Else
                            If vRevisao = 0 Then .Range("BE" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTSolN)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BG" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("BF" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("BF" & j).Value < .Range("BE" & j).Value And rsOS.Fields(11) = 3 Then .Range("BE" & j).Value = .Range("BF" & j).Value
                End If
'INICIO         'Acabamento Caldeiraria
                If rsOS.Fields(3) = "3000.3105.SC-01" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("BH" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("BG" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTAcabC)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("BH" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTAcabC)
                            Wend
                            If vRevisao = 0 Then .Range("BH" & j).Value = somaTempoAcumulado(CDate(vText2), vTAcabC)
                        Else
                            If vRevisao = 0 Then .Range("BH" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTAcabC)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BJ" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("BI" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("BI" & j).Value < .Range("BH" & j).Value And rsOS.Fields(11) = 3 Then .Range("BH" & j).Value = .Range("BI" & j).Value
                End If
'INICIO         'Acabamento Naval
                If rsOS.Fields(3) = "3000.3105.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("BK" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("BJ" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTAcabN)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("BK" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTAcabN)
                            Wend
                            If vRevisao = 0 Then .Range("BK" & j).Value = somaTempoAcumulado(CDate(vText2), vTAcabN)
                        Else
                            If vRevisao = 0 Then .Range("BK" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTAcabN)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BM" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("BL" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("BL" & j).Value < .Range("BK" & j).Value And rsOS.Fields(11) = 3 Then .Range("BK" & j).Value = .Range("BL" & j).Value
                End If
        
'INICIO         'Controle de Qualidade
                If rsOS.Fields(3) = "7000.7103.SC-02" Then
                    If rsOS.Fields(4) = " " Then
                        .Range("BN" & j).Value = Format("0000:00", "hh:mm") ' 3101.SC-01
                    Else
                        '.Range("BJ" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTAcabN)
'-------------------------------------------------------------------------
'TESTE
                        vAcumulaData1 = Split(rsOS.Fields(4), ":")
                        vVinteQuatroHoras = Split(vText, ":")
                        
                        If vAcumulaData1(0) >= 24 Then
                            While vAcumulaData1(0) >= 24
                                For F = 0 To 1
                                    If Val(vAcumulaData1(F)) > Val(vVinteQuatroHoras(F)) Then
                                        vAcumulaData2(F) = CInt(vAcumulaData1(F)) - CInt(vVinteQuatroHoras(F))
                                    End If
                                Next F
                                vText2 = vAcumulaData2(0) & ":" & Format(vAcumulaData2(1), "00")
                                vAcumulaData1(0) = vAcumulaData2(0)
                                vAcumulaData1(1) = vAcumulaData2(1)
                                If vRevisao = 0 Then .Range("BN" & j).Value = somaTempoAcumulado(Format(vText, "hh:mm"), vTQua)
                            Wend
                            If vRevisao = 0 Then .Range("BN" & j).Value = somaTempoAcumulado(CDate(vText2), vTQua)
                        Else
                            If vRevisao = 0 Then .Range("BN" & j).Value = somaTempoAcumulado(rsOS.Fields(4), vTQua)
                        End If
'TESTE
'-------------------------------------------------------------------------
                    End If
                    .Range("BP" & j).Value = somaTempoCC(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    'Calcula tempo realizado
                    If vRevisao = 0 Then .Range("BO" & j).Value = somaTempoReal(rsOS.Fields(1), rsOS.Fields(3), rsOS.Fields(9), rsOS.Fields(5))
                    If .Range("BO" & j).Value < .Range("BN" & j).Value And rsOS.Fields(11) = 3 Then .Range("BN" & j).Value = .Range("BO" & j).Value
                End If
            End With
            rsOS.MoveNext
            If rsOS.EOF Then Exit Do
        Loop
        j = j + 1
        vTCNC1 = ""
        vTCNC2 = ""
        vTGuil = ""
        vTTPuns = ""
        vTRosq = ""
        vTFRadial = ""
        vTFPrisma = ""
        vTFMag = ""
        vTSerraFita = ""
        vTCorte = ""
        vTDesemp = ""
        vTPrensa = ""
        vTMonC = ""
        vTMonN = ""
        vTSolC = ""
        vTSolN = ""
        vTAcabC = ""
        vTAcabN = ""
        vTCal = ""
        vTTrac = ""
        vTQua = ""
    Wend
    
    Plan.Range("A1").Select
    
    Plan.Columns("E:BP").EntireColumn.AutoFit 'Ajusta as colunas
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.ActiveWorkbook.SaveAs cdg.FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    
    Plan.EnableEvents = True
    Plan.ScreenUpdating = True
    With Plan
        .Calculation = xlAutomatic
        .MaxChange = 0.001
    End With
    ActiveWorkbook.PrecisionAsDisplayed = False
    
    
    
    'Plan.Close
    Set Plan = Nothing
    SkinLabel1.Visible = False
    mobjMsg.Abrir "Dados exportados com sucesso", Ok, informacao, "Atenção"
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, O MSOffice não esta instalado nesse computador!", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Function achaBaixaOS(vOSbaixada As Integer, vCCBaixado As String)
    Dim rsachaBaixaOS As New ADODB.Recordset
    Dim SqlachaBaixaOS As String
    SqlachaBaixaOS = "select a.idoperacao,a.idcc,b.percentualbaixado,'',a.idos from tbMPItens as a left join tbMPBaixaParcial as b on a.idos = b.idos and a.revisaoos = b.revisao and a.idoperacao = b.idoperacao where a.idos = '" & vOSbaixada & "' and a.idcc = '" & vCCBaixado & "'"
    rsachaBaixaOS.Open SqlachaBaixaOS, cnBanco, adOpenKeyset, adLockReadOnly
    If rsachaBaixaOS.RecordCount > 0 Then achaBaixaOS = rsachaBaixaOS.Fields(2)
    rsachaBaixaOS.Close
End Function

Private Function somaTempoAcumulado(vTempo As Date, vOndeAcumula As String)
    Dim seg As Long, min As Long, hora As Long
    Dim tempo As Long
    Dim matriz2

    matriz2 = Split(vTempo, ":")
    tempo = tempo + (CLng(matriz2(0)) * 3600)
    tempo = tempo + (CLng(matriz2(1)) * 60)
    tempo = tempo + CLng(matriz2(2))
    'hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    'tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    'min = Int(tempo / 60) ' aki calculamos os minutos
    'seg = tempo - (min * 60) 'aki subtraimos do tempo a qtde de segundos referentes aos minutos inteiros sobrandos os segundos
    
    If vOndeAcumula <> "" Then
        matriz2 = Split(vOndeAcumula, ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
        tempo = tempo + CLng(matriz2(2))
    End If
    
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    seg = tempo - (min * 60) 'aki subtraimos do tempo a qtde de segundos referentes aos minutos inteiros sobrandos os segundos
    
    vOndeAcumula = Format(hora, "0000") & ":" & Format(min, "00") & ":" & Format(seg, "00")
    somaTempoAcumulado = vOndeAcumula
End Function

Private Function somaTempoPPSAtraso(vTempo, vOndeAcumula As String)
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

Private Function somaTempoCC(vOS As Integer, vCC As String, vRevisao As Integer, vsemana As Integer)
    Dim tempo As Long
    Dim seg As Long, min As Long, hora As Long
    Dim matriz
    Dim matriz2
    Dim rsSomaCC As New ADODB.Recordset
    Dim SqlSomaCC As String
    
    SqlSomaCC = "set DATEFIRST 1 select b.idprogramacao,b.idos,b.idcc,a.codigobarra,a.chapa,a.dataent,CONVERT (VARCHAR, a.horaent, 108) as Hora_Ent,CONVERT (VARCHAR, a.horasai, 108) as Hora_Sai,CONVERT (VARCHAR, (a.horasai - horaent), 108) as Hora_Aprop,b.status " & _
                "from tbOsMov as a inner join tbmpitens as b on a.codigobarra = b.codigobarra where a.datasai is not null and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and  b.idcc = '" & vCC & "' and b.idos = '" & vOS & "' and b.revisaoos = '" & vRevisao & "' and DATEPART(WK,B.dataprevista) = '" & vsemana & "' order by B.dataprevista,b.idos,b.revisaoos,b.idcc"
    rsSomaCC.Open SqlSomaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSomaCC.RecordCount > 0 Then
        If rsSomaCC.Fields(9) > 3 Then
            vStatusOperacao = 3
        Else
            vStatusOperacao = rsSomaCC.Fields(9)
        End If
    End If
    
    tempo = 0
    While Not rsSomaCC.EOF
        matriz2 = Split(rsSomaCC.Fields(8), ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
        tempo = tempo + CLng(matriz2(2))
        rsSomaCC.MoveNext
    Wend
    rsSomaCC.Close
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    seg = tempo - (min * 60) 'aki subtraimos do tempo a qtde de segundos referentes aos minutos inteiros sobrandos os segundos
    
    somaTempoCC = Format(hora, "0000") & ":" & Format(min, "00") & ":" & Format(seg, "00")
    'lblTotal.Caption = Format(hora, "00") & ":" & Format(min, "00") & ":" & Format(seg, "00")
End Function

'SOMA TEMPO REALIZADO
'A ROTINA CONSIDERA SE HÁ UMA OU MAIS OPERAÇÕES NO MESMO CENTRO DE CUSTO
'SE O STATUS ESTA FECHADO OU NÃO (3 OU 2)
Private Function somaTempoReal(vOS As Integer, vCC As String, vRevisao As Integer, vsemana As Integer)
On Error Resume Next
    Dim tempo As Long
    Dim seg As Long, min As Long, hora As Long
    'Dim matriz
    Dim matriz2
    Dim rsTempoReal As New ADODB.Recordset
    Dim SqlTempoReal As String
    Dim vConverte As Double
    
'    SqlTempoReal = "select B.idos,B.idoperacao,B.idcc,'Horas' = dbo.FN_CONVMIN((cast(replace(B.tempocalc,'.','') as money)/100)),B.status,c.percentualBaixado from tbmpitens as B left join tbMPBaixaParcial as C " & _
'                "on b.idos = c.idos and b.idoperacao = c.idoperacao where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and b.idos = '" & vOS & "' and b.idcc ='" & vCC & "' and substring(b.idcc,1,1) <> '7' order by B.idos,B.idcc,B.idoperacao"
'    rsTempoReal.Open SqlTempoReal, cnBanco, adOpenKeyset, adLockReadOnly
    
'    If vOS = 724 And vCC = "3000.3101.SC-03" Then
'        Msgbox "aki"
'    End If
    
    SqlTempoReal = "SET DATEFIRST 1 select B.idos,B.idoperacao,B.idcc,'Horas' = dbo.FN_CONVMIN((cast(replace(B.tempocalc,'.','') as money)/100)),B.status,c.percentualBaixado from tbmpitens as B left join tbMPBaixaParcial as C " & _
                "on b.idos = c.idos and b.revisaoos = c.revisao and b.idoperacao = c.idoperacao where B.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' AND '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and b.idos = '" & vOS & "' and b.revisaoos = '" & vRevisao & "' and DATEPART(WK,B.dataprevista) = '" & vsemana & "' and b.idcc ='" & vCC & "' order by B.dataprevista,B.idos,b.revisaoos,B.idcc"
    rsTempoReal.Open SqlTempoReal, cnBanco, adOpenKeyset, adLockReadOnly
    
    tempo = 0
    While Not rsTempoReal.EOF
        
        If rsTempoReal.Fields(4) = 2 And Not IsNull(rsTempoReal.Fields(5)) Then
             vConverte = Replace(rsTempoReal.Fields(3), ":", ",") * rsTempoReal.Fields(5) / 100
             vConverte = Replace(Round(vConverte), ",", ":")
             matriz2 = Split(vConverte, ":")
             tempo = tempo + (CLng(matriz2(0)) * 3600)
             tempo = tempo + (CLng(matriz2(1)) * 60)
             tempo = tempo + CLng(matriz2(2))
        ElseIf rsTempoReal.Fields(4) = 3 Then
             matriz2 = Split(rsTempoReal.Fields(3), ":")
             tempo = tempo + (CLng(matriz2(0)) * 3600)
             tempo = tempo + (CLng(matriz2(1)) * 60)
             tempo = tempo + CLng(matriz2(2))
        End If
        
''        matriz2 = Split(rsSomaCC.Fields(8), ":")
'        tempo = tempo + (CLng(matriz2(0)) * 3600)
'        tempo = tempo + (CLng(matriz2(1)) * 60)
'        tempo = tempo + CLng(matriz2(2))
        rsTempoReal.MoveNext
    Wend
    rsTempoReal.Close
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    seg = tempo - (min * 60) 'aki subtraimos do tempo a qtde de segundos referentes aos minutos inteiros sobrandos os segundos
    
    somaTempoReal = Format(hora, "0000") & ":" & Format(min, "00") & ":" & Format(seg, "00")
    'lblTotal.Caption = Format(hora, "00") & ":" & Format(min, "00") & ":" & Format(seg, "00")
End Function

'REALIZAR SOMA DOS DOIS TEMPOS (NAO DEU CERTO)
'Private Function somaDoisTempos(vTempo1 As String, vTempo2 As String)
'On Error Resume Next
'    Dim seg As Long, min As Long, hora As Long
'    Dim tempo As Long
'    Dim matriz
'    Dim matriz2
'
'    matriz = Split(vTempo1, ":")
'    tempo = tempo + (CLng(matriz(0)) * 3600)
'    tempo = tempo + (CLng(matriz(1)) * 60)
'    tempo = tempo + CLng(matriz(2))
'
'    matriz2 = Split(vTempo2, ":")
'    tempo = tempo + (CLng(matriz2(0)) * 3600)
'    tempo = tempo + (CLng(matriz2(1)) * 60)
'    tempo = tempo + CLng(matriz2(2))
'
'    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
'    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
'    min = Int(tempo / 60) ' aki calculamos os minutos
'    seg = tempo - (min * 60) 'aki subtraimos do tempo a qtde de segundos referentes aos minutos inteiros sobrandos os segundos
'
'    somaDoisTempos = Format(hora, "0000") & ":" & Format(min, "00") & ":" & Format(seg, "00")
'    'lblTotal.Caption = Format(hora, "00") & ":" & Format(min, "00") & ":" & Format(seg, "00")
'End Function

'Private Sub preencheVermelho(posi As Integer)
'    Plan.Range("A" & posi & ":Z" & posi).Select
'    With Plan.Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorAccent2
'        .TintAndShade = 0.799981688894314
'        .PatternTintAndShade = 0
'    End With
'End Sub

'Private Sub preencheBranco(posi As Integer)
'    Plan.Range("A" & posi & ":Z" & posi).Select
'    With Plan.Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorDark1
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'End Sub

'Private Sub contornoDVermelho(vLin As Integer, vCol As String)
'    Plan.Range(vCol & vLin).Select
'    With Plan.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .ThemeColor = 6
'        .TintAndShade = -0.249946592608417
'        .Weight = xlThin
'    End With
'End Sub

'Private Sub contornoEVermelho(vLin As Integer, vCol As String)
'    Plan.Range(vCol & vLin).Select
'    With Plan.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .ThemeColor = 6
'        .TintAndShade = -0.249946592608417
'        .Weight = xlThin
'    End With
'End Sub

'Private Sub contornoBVermelho(vLin As Integer)
'    Plan.Range("A" & vLin & ":Z" & vLin).Select
'    With Plan.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .ThemeColor = 6
'        .TintAndShade = -0.249946592608417
'        .Weight = xlThin
'    End With
'End Sub


'Private Sub tracejarVermelho(vLin As Integer)
'    Plan.Range("A" & vLin & ":Z" & vLin).Select
'    With Plan.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlDot
'        .ThemeColor = 6
'        .TintAndShade = -0.249946592608417
'        .Weight = xlThin
'    End With
'End Sub

Private Sub ExportaExcelEvolucao()
'On Error Resume Next
    'Dim vTCNC1 As String, vTCNC2 As String, vTGuil As String, vTTPuns As String, vTRosq As String, vTFRadial As String, vTFPrisma As String, vTFMag As String, vTSerraFita As String, vTCorte As String, vTDesemp As String, vTPrensa As String, vTMonC As String, vTMonN As String, vTSolC As String, vTSolN As String, vTAcabC As String, vTAcabN As String, vTCal As String, vTTrac As String
    
    Dim j As Integer, K As Integer, L As Integer, X As Integer
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")

    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\Fabricacao.xlsx"
    Plan.Visible = True
    Plan.UserControl = False

    'Dados das OSs que estão dentro do intervalo de tempo informado
    Dim rsCab As New ADODB.Recordset
    Dim SqlCab As String
    Dim rsEvo As New ADODB.Recordset
    Dim SqlEvo As String
    Dim vOS As Integer
    Dim vLin As Integer, vCol As Integer, vContaCol As Integer
    
    
'codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-01','4000.4101.SC-02','4000.4101.SC-03','7000.7103.SC-02')
    
    
'    SqlCab = "select case when a.codreduzido = '3000.3106.SC-01' then '3000.3101.SC-00' else a.codreduzido end as codreduzido,SUBSTRING(c.NOME,19,50) as nome_CC " & _
'    "from tbFormula as a inner join tbApropriacao as b on a.codreduzido = b.codreduzido inner join corporerm.dbo.GCCUSTO as c on b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS = c.CODREDUZIDO " & _
'    "where SUBSTRING(a.codreduzido,1,9) <> '7000.7108' group by a.codreduzido,c.NOME order by codreduzido"
    
    SqlCab = "select case when a.codreduzido = '3000.3106.SC-01' then '3000.3101.SC-00' else a.codreduzido end as codreduzido,SUBSTRING(c.NOME,19,50) as nome_CC " & _
    "from tbFormula as a inner join tbApropriacao as b on a.codreduzido = b.codreduzido inner join  " & vBancoTotvs & ".dbo.GCCUSTO as c on b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS = c.CODREDUZIDO " & _
    "where a.codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','7000.7103.SC-02') group by a.codreduzido,c.NOME order by codreduzido"
    rsCab.Open SqlCab, cnBanco, adOpenKeyset, adLockReadOnly
    
    vLin = 3
    vCol = 9
    vContaCol = 9
    While Not rsCab.EOF
        With Plan
            If Mid$(rsCab.Fields(0), 1, 4) <> "4000" Then

                If rsCab.Fields(0) = "3000.3103.SC-01" Then vCol = vCol + 1
                
                .Cells(vLin, vCol) = rsCab.Fields(0)
                If .Cells(vLin, vCol) = "3000.3101.SC-00" Then .Cells(vLin, vCol) = "3000.3106.SC-01"
                .Cells(vLin + 1, vCol) = rsCab.Fields(1)
                vCol = vCol + 1
            End If
            
            rsCab.MoveNext
            vContaCol = vContaCol + 1
            
            'SE O CENTRO DE CUSTO FOR = 3000.3101.SC-01 O SISTEMA VAI PULAR MAIS UMA COLUNA
            vContaCol = vContaCol + 1
            If rsCab.EOF Then
                .Cells(vLin, vCol) = "3"
                .Cells(vLin + 1, vCol) = "Rel. Fabricação"
                
                .Cells(vLin, vCol + 1) = "10"
                .Cells(vLin + 1, vCol + 1) = "Rel. Pintura"
                
                .Cells(vLin, vCol + 2) = "11"
                .Cells(vLin + 1, vCol + 2) = "Rel. Expedição"
            End If
        End With
    Wend
    
    
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    
    Dim rsGrava As New ADODB.Recordset
    Dim sqlGrava As String
   
'    sqlGrava = "CREATE TABLE ##PesoPosicoes (fce integer,codlm integer,codseq integer,pesototal float)" & _
'               "INSERT INTO ##PesoPosicoes (fce,codlm, codseq, pesototal) " & _
'               "Select a.fce,a.codlm,MAX(a.codseq) as codseq,sum((a.quantcj*a.quantunit*a.pesounit)) as PesoTotal from tbItemLM as a " & _
'               "inner join tbPosicoes as b on a.codigopos = b.codigopos inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbProjetos as d on c.codprojeto = d.codprojeto " & _
'               "where a.fce = '" & Val(Text1.Text) & "' group by a.fce,d.projeto,c.desenho,b.posicao,b.descposicao,a.codlm order by a.fce,d.projeto,c.desenho,b.posicao"
'    rsGrava.Open sqlGrava, cnBanco
'
'    SqlEvo = "Select a.fce,d.projeto,MAX(a.codlm) codlm,c.desenho,Max(c.revisao) as revisao,b.descposicao as descricao,b.posicao as posicao,MAX(a.quantcj) as quantidade,MAX(f.pesototal) as PesTotal,e.idoperacao,e.idcc,MAX(e.status) as status " & _
'             "from tbItemLM as a inner join tbPosicoes as b on a.codigopos = b.codigopos inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbProjetos as d on c.codprojeto = d.codprojeto inner join tbositens as e on a.fce = e.fce and " & _
'             "a.codlm = e.codlm and a.codseq = e.codseq inner join ##PesoPosicoes as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq where a.fce = '" & Val(Text1.Text) & "' group by a.fce,d.projeto,c.desenho,b.posicao,b.descposicao,e.idoperacao,e.idcc " & _
'             "order by a.fce,d.projeto,c.desenho,b.posicao,e.idoperacao"
'    rsEvo.Open SqlEvo, cnBanco, adOpenKeyset, adLockReadOnly

'    SqlEvo = "Select a.fce,d.projeto,MAX(a.codlm) as codlm,c.desenho,Max(c.revisao) as revisao,b.descposicao as descricao,b.posicao as posicao,MAX(a.quantcj) as quantidade,MAX(b.pesoposicao) AS PesoPosicao,e.idoperacao,e.idcc,MAX(e.status) as status,MAX(a.codseq) as codseq " & _
'    "from tbItemLM as a inner join tbPosicoes as b on a.codigopos = b.codigopos inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbProjetos as d on c.codprojeto = d.codprojeto inner join tbositens as e on a.fce = e.fce and " & _
'    "a.codlm = e.codlm and a.codseq = e.codseq where a.fce = '" & Val(Text1.Text) & "' and e.idcc in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','7000.7103.SC-02') " & _
'    "group by a.fce,d.projeto,c.desenho,b.posicao,b.descposicao,e.idoperacao,e.idcc order by a.fce,d.projeto,c.desenho,b.posicao,e.idoperacao"
    
    
    SqlEvo = "Select a.fce,d.projeto,MAX(a.codlm) as codlm,c.desenho,Max(c.revisao) as revisao,b.descposicao as descricao,b.posicao as posicao,MAX(a.quantcj) as quantidade,MAX(b.pesoposicao) AS PesoPosicao,e.idoperacao,e.idcc,MAX(e.status) as status,MAX(a.codseq) as codseq " & _
             "from tbItemLM as a inner join tbPosicoes as b on a.codigopos = b.codigopos inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbProjetos as d on c.codprojeto = d.codprojeto LEFT join tbositens as e on a.fce = e.fce and a.codlm = e.codlm and a.codseq = e.codseq " & _
             "and e.idcc in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','7000.7103.SC-02','') " & _
             "where a.fce = '" & Val(Text1.Text) & "' group by a.fce,d.projeto,c.desenho,b.posicao,b.descposicao,e.idoperacao,e.idcc order by a.fce,d.projeto,c.desenho,b.posicao,e.idoperacao"

    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsEvo.Open SqlEvo, cnBanco, adOpenKeyset, adLockReadOnly

    j = 5
    vLin = 3
    vCol = 9
    Dim vPesoRel As Double
    With Plan
    While Not rsEvo.EOF
'        If rsEvo.Fields(11) = 3 Then
            .Cells(j, 1) = rsEvo.Fields(0) 'FCE
            .Cells(j, 2) = rsEvo.Fields(3) 'Desenho
            .Cells(j, 3) = rsEvo.Fields(1) 'Projeto
            .Cells(j, 4) = rsEvo.Fields(4) 'Rev
            .Cells(j, 5) = rsEvo.Fields(6) 'Posição
            .Cells(j, 6) = rsEvo.Fields(5) 'Descrição
            .Cells(j, 7) = rsEvo.Fields(7) 'Quantidade
            .Cells(j, 8) = rsEvo.Fields(8) 'Peso Total
             
            If rsEvo.Fields(11) >= 2 Then
                For X = 9 To vContaCol
                    vPesoRel = 0
                    If Cells(vLin, vCol) = rsEvo.Fields(10) Then
                            If vCol < 32 Then
                                .Cells(j, vCol) = rsEvo.Fields(8)
                            End If
                    End If
                    If Cells(vLin, vCol) = "3" Then 'Fabricação
                        .Cells(j, vCol) = AchaRels(3, rsEvo.Fields(0), rsEvo.Fields(2), rsEvo.Fields(12), vPesoRel) '"R. Insp"
                    End If
                    If Cells(vLin, vCol) = "10" Then 'Pintura
                        .Cells(j, vCol) = AchaRels(10, rsEvo.Fields(0), rsEvo.Fields(2), rsEvo.Fields(12), vPesoRel) '"R. Pint"
                    End If
                    If Cells(vLin, vCol) = "11" Then 'Expedição
                        .Cells(j, vCol) = AchaRels(11, rsEvo.Fields(0), rsEvo.Fields(2), rsEvo.Fields(12), vPesoRel) '"R. Exp"
                    End If
                    vCol = vCol + 1
                Next
            End If
            vCol = 9
            
            rsEvo.MoveNext
            If Not rsEvo.EOF Then
            
                If rsEvo.Fields(1) = .Cells(j, 3) And rsEvo.Fields(3) = .Cells(j, 2) And rsEvo.Fields(6) = .Cells(j, 5) Then
                    j = j
                Else
                    j = j + 1
                End If
            End If
    Wend
    End With

    rsEvo.Close
    
    Plan.Range("A1").Select
    
    Plan.Columns("C:BD").EntireColumn.AutoFit 'Ajusta as colunas
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.ActiveWorkbook.SaveAs cdg.FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    'Plan.Close
    Set Plan = Nothing
    
    mobjMsg.Abrir "Dados exportados com sucesso", Ok, informacao, "Atenção"
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, O MSOffice não esta instalado nesse computador!", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Function AchaRels(vQual As Integer, vFCE As Integer, vLM As Integer, vSeq As Integer, vRetornaPeso As Double)
    Dim rsAchaRels As New ADODB.Recordset
    Dim SqlAchaRels As String
    SqlAchaRels = "select a.fce,a.codlm,a.codseq,sum(a.pesolib) from tbRelInspExpItens as a where a.fce = '" & vFCE & "' and a.codlm = '" & vLM & "' and a.codseq = '" & vSeq & "' and a.status = '" & vQual & "' group by a.fce,a.codlm,a.codseq order by a.codlm,a.codseq"
    rsAchaRels.Open SqlAchaRels, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaRels.RecordCount > 0 Then
        vRetornaPeso = rsAchaRels.Fields(3)
    Else
        vRetornaPeso = 0
    End If
    AchaRels = vRetornaPeso
    rsAchaRels.Close
    Set rsAchaRels = Nothing
End Function

'Daki para baixo: ROP - RELATORIO DE PARADAS

Private Sub preparaParada()
    SkinLabel1.Visible = True
    mobjMsg.Abrir "Esse procedimento pode demorar alguns minutos.", Ok, critico, "Atenção"
    Timer1.Enabled = True
End Sub

Private Sub preparaHA()
    SkinLabel1.Visible = True
    mobjMsg.Abrir "Esse procedimento pode demorar alguns minutos.", Ok, critico, "Atenção"
    Timer2.Enabled = True
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        DTPicker1.Value = ""
        converteSemana Val(Text2.Text), DTPicker1, Text3.Text
        If DTPicker1.Value = "" Then
            mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
            Exit Sub
        Else
            DTPicker2 = DTPicker1.Value + 6
        End If
    End If
End Sub

Private Sub Text2_LostFocus()
    DTPicker1.Value = ""
    converteSemana Val(Text2.Text), DTPicker1, Text3.Text
    If DTPicker1.Value = "" Then
        mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
        Exit Sub
    Else
        DTPicker2 = DTPicker1.Value + 6
    End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        DTPicker1.Value = ""
        converteSemana Val(Text2.Text), DTPicker1, Text3.Text
        If DTPicker1.Value = "" Then
            mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
            Exit Sub
        Else
            DTPicker2 = DTPicker1.Value + 6
        End If
    End If
End Sub

Private Sub Text3_LostFocus()
    DTPicker1.Value = ""
    converteSemana Val(Text2.Text), DTPicker1, Text3.Text
    If DTPicker1.Value = "" Then
        mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
        Exit Sub
    Else
        DTPicker2 = DTPicker1.Value + 6
    End If
End Sub

Private Sub Timer1_Timer()
    'excluiTabelaStopControl 1
    'criaTabelaStopControl 1
    transfDados
    somaTemposCC
    'Timer1.Enabled = False
    'SkinLabel1.Visible = False
    'FCRRop.Show 1
End Sub

Private Sub Timer2_Timer()
    ''deletaDadosStopControl
    
    excluiTabelaStopControl 2 'Tirar essa linha depois que os dois relatorios estiverem funcionando juntos
    criaTabelaStopControl 2 'Tirar essa linha depois que os dois relatorios estiverem funcionando juntos
    
    ''Retrabalho
    transfDadosHA
    somaTemposCSRetrabalho
    somaTemposPlanejadoCC
    
    ''paradas
    transfDados
    somaTemposCC
    
    somaTemposProgramadosCC
    
    DeletaExcesso
    
    
    Timer2.Enabled = False
    SkinLabel1.Visible = False
    FCRHApropriadas.Show 1
End Sub

Private Sub DeletaExcesso()
    Dim rsAcertaDados As New ADODB.Recordset
    Dim SqlAcertaDados As String
    SqlAcertaDados = "Delete from tbApropriaControle where substring(centrocusto,1,4) = '1000'"
    rsAcertaDados.Open SqlAcertaDados, cnBanco
End Sub

Private Sub excluiTabelaStopControl(vIndice As Integer)
On Error Resume Next
    Dim rsExcluirTb As New ADODB.Recordset
    Dim SqlExcluirTb As String
    If vIndice = 1 Then
        SqlExcluirTb = "Drop table tbApropriaControle"
        rsExcluirTb.Open SqlExcluirTb, cnBanco
    ElseIf vIndice = 2 Then
        SqlExcluirTb = "Drop table tbApropriaControle"
        rsExcluirTb.Open SqlExcluirTb, cnBanco
    End If
End Sub

Private Sub deletaDadosStopControl(vIndice As Integer)
    'Deleta todos os dados da tabela deletaDadosStopControl
    'para que possam ser inserido novos dados
    Dim rsDeletatbApropriaControle As New ADODB.Recordset
    Dim SqlDeletatbApropriaControle As String
    If vIndice = 1 Then
        SqlDeletatbApropriaControle = "Delete from tbApropriaControle"
        rsDeletatbApropriaControle.Open SqlDeletatbApropriaControle, cnBanco
    ElseIf vIndice = 2 Then
    End If
End Sub

Private Sub criaTabelaStopControl(vIndice As Integer)
    If vIndice = 1 Then
'        cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbApropriaControle(" & _
'        "registro NUMERIC NOT NULL," & _
'        "nome VARCHAR(100) NOT NULL," & _
'        "centrocusto VARCHAR(100) NOT NULL," & _
'        "dataentrada DATETIME NOT NULL," & _
'        "horaentrada DATETIME NOT NULL," & _
'        "datasaida DATETIME NOT NULL," & _
'        "horasaida DATETIME NOT NULL," & _
'        "idparada NUMERIC NOT NULL," & _
'        "nmparada VARCHAR(100) NOT NULL," & _
'        "tempoparada DATETIME NOT NULL," & _
'        "tempototalcc VARCHAR(30) NULL," & _
'        "tempototalpcc VARCHAR(30) NULL," & _
'        "tempototalparada VARCHAR(30) NULL," & _
'        "tempototal VARCHAR(30) NULL," & _
'        "percentualtotalparada VARCHAR(30) NULL," & _
'        "retrabalho VARCHAR(10) NULL," & _
'        "TempoCRetrabalho VARCHAR(30) NULL," & _
'        "TempoSRetrabalho VARCHAR(30) NULL," & _
'        "TempoCSRetrabalho VARCHAR(30) NULL)"
    
    ElseIf vIndice = 2 Then
        cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbApropriaControle(" & _
        "registro VARCHAR(50) NOT NULL,nome VARCHAR(100) NOT NULL," & _
        "centrocusto VARCHAR(100) NOT NULL,dataentrada DATETIME NOT NULL," & _
        "horaentrada DATETIME NOT NULL,datasaida DATETIME NOT NULL," & _
        "horasaida DATETIME NOT NULL,idparada NUMERIC NOT NULL," & _
        "nmparada VARCHAR(100) NOT NULL,tempoApropriado DATETIME NOT NULL," & _
        "retrabalho VARCHAR(10) NULL,TempoSRetrabalho VARCHAR(30) NULL," & _
        "TempoCSRetrabalho VARCHAR(30) NULL,TempoTotalApropriacao VARCHAR(30) NULL," & _
        "TempoTotalGeral VARCHAR(30) NULL,TempoPlanejadoCC VARCHAR(30) NULL," & _
        "TempoPlanejadoTotal VARCHAR(30) NULL,TempoTotalCarteira VARCHAR(30) NULL," & _
        "TempoGeralCarteira VARCHAR(30) NULL,tempoparada DATETIME NOT NULL," & _
        "tempototalcc VARCHAR(30) NULL,tempototalpcc VARCHAR(30) NULL," & _
        "tempototalparada VARCHAR(30) NULL,tempototal VARCHAR(30) NULL," & _
        "percentualtotalparada VARCHAR(30) NULL,PPSPorCC VARCHAR(30) NULL," & _
        "AtrasoPorCC VARCHAR(30) NULL," & _
        "PPSTotal VARCHAR(30) NULL," & _
        "AtrasoTotal VARCHAR(30) NULL," & _
        "PPSeAtrasoPorCC VARCHAR(30) NULL," & _
        "PPSeAtrasoSoma VARCHAR(30) NULL," & _
        "PPSRealPorCC VARCHAR(30) NULL," & _
        "PPSRealTotalPorCC VARCHAR(30) NULL," & _
        "ExtraPPSRealPorCC VARCHAR(30) NULL," & _
        "ExtraPPSRealTotalPorCC VARCHAR(30) NULL," & _
        "ExtraPPSRealSoma VARCHAR(30) NULL," & _
        "TempoTotalRealizado VARCHAR(30) NULL)"

'        cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbApropriaControle(" & _
'        "registro NUMERIC NOT NULL," & _
'        "nome VARCHAR(100) NOT NULL," & _
'        "centrocusto VARCHAR(100) NOT NULL," & _
'        "dataentrada DATETIME NOT NULL," & _
'        "horaentrada DATETIME NOT NULL," & _
'        "datasaida DATETIME NOT NULL," & _
'        "horasaida DATETIME NOT NULL," & _
'        "idparada NUMERIC NOT NULL," & _
'        "nmparada VARCHAR(100) NOT NULL," & _
'        "tempoApropriado DATETIME NOT NULL," & _
'        "retrabalho VARCHAR(10) NULL," & _
'        "TempoSRetrabalho VARCHAR(30) NULL," & _
'        "TempoCSRetrabalho VARCHAR(30) NULL," & _
'        "TempoTotalApropriacao VARCHAR(30) NULL," & _
'        "TempoTotalGeral VARCHAR(30) NULL," & _
'        "TempoPlanejadoCC VARCHAR(30) NULL," & _
'        "TempoPlanejadoTotal VARCHAR(30) NULL," & _
'        "TempoTotalCarteira VARCHAR(30) NULL," & _
'        "TempoGeralCarteira VARCHAR(30) NULL)"
    End If
End Sub

Private Sub transfDados()
    'Transfere dados referente à Paradas
'    SqlApropriacao = "select A.CHAPA,C.NOME,e.NOME as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada " & _
'                     "from tbOsMov AS A INNER JOIN tbParadas AS B ON a.idparada<> 'ERRO' and A.idparada = B.codigo inner join CORPORERM.dbo.PFUNC as C on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = C.CHAPA inner join CORPORERM.dbo.PFRATEIOFIXO as d on a.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = d.CHAPA " & _
'                     "inner join CORPORERM.dbo.GCCUSTO as e on d.CODCCUSTO = e.CODCCUSTO where A.idparada in(9001,9002,9003,9004,9005,9006,9007,9008,9009,9010,9011,9012,9013,9014,9015,9016,9017,9018,9019,9020) and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' ORDER BY E.NOME,A.dataent,A.horaent"
    SqlApropriacao = "select A.CHAPA,C.NOME,e.NOME as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada from tbOsMov AS A INNER JOIN tbParadas AS B ON a.idparada<> 'ERRO' and A.idparada = B.codigo inner join  " & vBancoTotvs & ".dbo.PFUNC as C on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = C.CHAPA " & _
                     "inner join  " & vBancoTotvs & ".dbo.PFRATEIOFIXO as d on a.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = d.CHAPA inner join " & vBancoTotvs & ".dbo.GCCUSTO as e on d.CODCCUSTO = e.CODCCUSTO where A.idparada in(9001,9002,9003,9004,9005,9006,9007,9008,9009,9010,9011,9012,9013,9014,9015,9016,9017,9018,9019,9020) and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' union " & _
                     "SELECT A.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AI,C.NOME COLLATE SQL_Latin1_General_CP1_CI_AI,c.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada from tbOsMov AS A INNER JOIN tbParadas AS B ON a.idparada<> 'ERRO' and " & _
                     "A.idparada = B.codigo inner join tbTerceirizados as C on A.chapa = C.chapa where A.idparada in(9001,9002,9003,9004,9005,9006,9007,9008,9009,9010,9011,9012,9013,9014,9015,9016,9017,9018,9019,9020) and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' ORDER BY CentroCusto,DataEntrada,HoraEntrada"
    rsApropriacao.Open SqlApropriacao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsApropriacao.RecordCount > 0 Then
        LocalizaParada
    End If
    rsApropriacao.Close
    Set rsApropriacao = Nothing
End Sub

Private Sub transfDadosHA()
    'Transfere dados referente à Horas Apropriadas
    Dim rsTempoParada As New ADODB.Recordset
    Dim SqlTempoParada As String
    Dim vHoraEntrada As String
    Dim vHoraSaida As String

    Dim vDifHora As String
    
    'Seleciona os dados
'    SqlApropriacao = "select A.CHAPA,C.NOME,e.NOME as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada,a.codigobarra,h.idretrabalho " & _
'                     "from tbOsMov AS A INNER JOIN tbParadas AS B ON A.idparada = B.codigo inner join CORPORERM.dbo.PFUNC as C on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = C.CHAPA inner join CORPORERM.dbo.PFRATEIOFIXO as d on a.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = d.CHAPA inner join CORPORERM.dbo.GCCUSTO as e on d.CODCCUSTO = e.CODCCUSTO " & _
'                     "left join tbMPItens as f on a.codigobarra = f.codigobarra left join tbRetrabalho as h on f.idprogramacao = h.idprogramacao " & _
'                     "where substring(e.NOME,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
'                     "and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and substring(e.nome,1,4) = '3000' ORDER BY E.NOME,A.dataent,A.horaent"
    SqlApropriacao = "select A.CHAPA,C.NOME,e.NOME as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada,a.codigobarra,h.idretrabalho from tbOsMov AS A INNER JOIN tbParadas AS B ON A.idparada = B.codigo inner join " & vBancoTotvs & ".dbo.PFUNC as C on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = C.CHAPA " & _
                     "inner join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as d on a.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = d.CHAPA inner join " & vBancoTotvs & ".dbo.GCCUSTO as e on d.CODCCUSTO = e.CODCCUSTO left join tbMPItens as f on a.codigobarra = f.codigobarra left join tbRetrabalho as h on f.idprogramacao = h.idprogramacao " & _
                     "where substring(e.NOME,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                     "and A.dataent BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and substring(e.nome,1,4) = '3000' union select A.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AI as CHAPA,c.NOME COLLATE SQL_Latin1_General_CP1_CI_AI as NOME,c.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI as CentroCusto,CONVERT (VARCHAR, a.dataent, 103) as DataEntrada,CONVERT (VARCHAR, a.horaent, 108) as HoraEntrada, " & _
                     "CONVERT (VARCHAR, a.datasai, 103) as DataSaida,CONVERT (VARCHAR, a.horasai, 108) as HoraSaida,A.idparada,B.nmparada,a.codigobarra,h.idretrabalho from tbOsMov AS A INNER JOIN tbParadas AS B ON A.idparada = B.codigo inner join tbTerceirizados as C on a.chapa = C.CHAPA left join tbMPItens as f on a.codigobarra = f.codigobarra left join tbRetrabalho as h on f.idprogramacao = h.idprogramacao where A.dataent " & _
                     "BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' ORDER BY CentroCusto,DataEntrada,HoraSaida"
    rsApropriacao.Open SqlApropriacao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsApropriacao.RecordCount <= 0 Then
        rsApropriacao.Close
        Set rsApropriacao = Nothing
        Exit Sub
    End If

    'Abaixo: transfere os dados selecionados para a tabela abaixo
    SqlTempoParada = "Select * from tbApropriaControle"
    rsTempoParada.Open SqlTempoParada, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsApropriacao.RecordCount > 0 Then Principal.ProgressBar1.Max = rsApropriacao.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Transferindo dados para tabela temporária..."
    Do While Not rsApropriacao.EOF
        Principal.ProgressBar1.Value = vProgress
        rsTempoParada.AddNew
        rsTempoParada.Fields(0) = rsApropriacao.Fields(0)
        rsTempoParada.Fields(1) = rsApropriacao.Fields(1)
        rsTempoParada.Fields(2) = rsApropriacao.Fields(2)
        rsTempoParada.Fields(3) = rsApropriacao.Fields(3)
        rsTempoParada.Fields(4) = rsApropriacao.Fields(4)
        rsTempoParada.Fields(5) = rsApropriacao.Fields(5)
        rsTempoParada.Fields(6) = rsApropriacao.Fields(6)
        rsTempoParada.Fields(7) = rsApropriacao.Fields(7)
        rsTempoParada.Fields(8) = rsApropriacao.Fields(8)
        vHoraEntrada = Format(rsApropriacao.Fields(6), "hh:mm")
        vHoraSaida = Format(rsApropriacao.Fields(4), "hh:mm")
        vDifHora = Format(TimeValue(vHoraSaida) - TimeValue(vHoraEntrada), "hh:mm")
        rsTempoParada.Fields(9) = vDifHora
        If IsNull(rsApropriacao.Fields(10)) Then
            rsTempoParada.Fields(10) = "S.Ret"
        Else
            rsTempoParada.Fields(10) = "C.Ret"
        End If
        
        rsTempoParada.Fields(19) = "00:00"
        
        rsApropriacao.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Acertando dados..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsTempoParada.Update
    rsTempoParada.Close
    rsApropriacao.Close
    
    
    Dim rsAchaPlan As New ADODB.Recordset
    Dim SqlAchaPlan As String
    
    
    SqlAchaPlan = "select codreduzido from tbFormula where codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') group by codreduzido order by codreduzido"
    rsAchaPlan.Open SqlAchaPlan, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaPlan.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaPlan.RecordCount
    vProgress = 0
    Do While Not rsAchaPlan.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaPlan.Fields(0)
        SqlTempoParada = "select * from tbApropriaControle where SUBSTRING(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsTempoParada.Open SqlTempoParada, cnBanco, adOpenKeyset, adLockOptimistic
        If rsTempoParada.RecordCount = 0 Then
            rsTempoParada.AddNew
            rsTempoParada.Fields(0) = "0" 'Registro
            rsTempoParada.Fields(1) = "-" 'Nome do colaborador
            rsTempoParada.Fields(2) = vCentroCusto ' Nome Centro do custo
            rsTempoParada.Fields(3) = DTPicker1.Value 'Data de entrada
            rsTempoParada.Fields(4) = "00:00" 'hora de entrada
            rsTempoParada.Fields(5) = DTPicker1.Value 'Data de Saida
            rsTempoParada.Fields(6) = "00:00" 'Hora de Saida
            rsTempoParada.Fields(7) = "9019" 'Identificador de parada
            rsTempoParada.Fields(8) = "FIM DE EXPEDIENTE" 'Nome da parada
            rsTempoParada.Fields(9) = "00:00" 'Tempo apropriado
            rsTempoParada.Fields(10) = "S.Ret" 'Tipo de apropriação
            rsTempoParada.Fields(15) = "00:00" 'Tempo apropriado
            rsTempoParada.Fields(19) = "00:00" 'Tempo apropriado
            rsTempoParada.Update
        End If
        rsTempoParada.Close
        rsAchaPlan.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaPlan.Close
    Set rsAchaPlan = Nothing
    
    'rsTempoParada.Update
    'rsTempoParada.Close
    'rsApropriacao.Close
    Set rsApropriacao = Nothing
End Sub

Private Sub somaTemposCC()
    Dim rsAchaCC As New ADODB.Recordset
    Dim SqlAchaCC As String
    Dim rsSomaTempoCC As New ADODB.Recordset
    Dim SqlSomaTempoCC As String
    
    Dim rsInsereTempoTotalCC As New ADODB.Recordset
    Dim SqlInsereTempoTotalCC As String
    
    
    Dim vCentroCusto As String
    Dim vIdParada As Integer
    
    '1ª Parte - ENCONTRA TEMPO TOTAL DE PARADAS DO CENTRO DE CUSTO
    SqlAchaCC = "select centrocusto from tbApropriaControle group by centrocusto order by centrocusto"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly

    If rsAchaCC.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaCC.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Calculando paradas por Centro de Custo..."
    Do While Not rsAchaCC.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaCC.Fields(0)
        SqlSomaTempoCC = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoparada as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.centrocusto = '" & vCentroCusto & "'"
        rsSomaTempoCC.Open SqlSomaTempoCC, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalCC = "Update tbApropriaControle set tempototalcc = '" & rsSomaTempoCC.Fields(0) & "' where centrocusto = '" & vCentroCusto & "'"
        rsInsereTempoTotalCC.Open SqlInsereTempoTotalCC, cnBanco
                                 
        rsSomaTempoCC.Close
        rsAchaCC.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo total por Centro de Custo"
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    Set rsSomaTempoCC = Nothing
    
    
    '2ª Parte - ENCONTRA TEMPO TOTAL DE CADA PARADA DENTRO DO CENTRO DE CUSTO
    SqlAchaCC = "select centrocusto,idparada from tbApropriaControle group by centrocusto,idparada order by centrocusto,idparada"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaCC.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaCC.RecordCount
    vProgress = 0
    Do While Not rsAchaCC.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaCC.Fields(0)
        vIdParada = rsAchaCC.Fields(1)
        SqlSomaTempoCC = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoparada as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.centrocusto = '" & vCentroCusto & "' and a.idparada = '" & vIdParada & "'"
        rsSomaTempoCC.Open SqlSomaTempoCC, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalCC = "Update tbApropriaControle set tempototalpcc = '" & rsSomaTempoCC.Fields(0) & "' where centrocusto = '" & vCentroCusto & "' and idparada = '" & vIdParada & "'"
        rsInsereTempoTotalCC.Open SqlInsereTempoTotalCC, cnBanco
                                 
        rsSomaTempoCC.Close
        rsAchaCC.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo de parada por período..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    Set rsSomaTempoCC = Nothing
    
    '3ª Parte - ENCONTRA TEMPO TOTAL DE CADA PARADA NO PERÍODO
    SqlAchaCC = "select centrocusto,idparada from tbApropriaControle group by centrocusto,idparada order by idparada,centrocusto"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaCC.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaCC.RecordCount
    vProgress = 0
    Do While Not rsAchaCC.EOF
        Principal.ProgressBar1.Value = vProgress
        vIdParada = rsAchaCC.Fields(1)
        SqlSomaTempoCC = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoparada as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.idparada = '" & vIdParada & "'"
        rsSomaTempoCC.Open SqlSomaTempoCC, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalCC = "Update tbApropriaControle set tempototalparada = '" & rsSomaTempoCC.Fields(0) & "' where idparada = '" & vIdParada & "'"
        rsInsereTempoTotalCC.Open SqlInsereTempoTotalCC, cnBanco
                                 
        rsSomaTempoCC.Close
        rsAchaCC.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo total do período..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    Set rsSomaTempoCC = Nothing
    
    '4ª Parte - ENCONTRA TEMPO TOTAL PERÍODO
    SqlAchaCC = "select centrocusto,idparada from tbApropriaControle group by centrocusto,idparada order by idparada,centrocusto"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaCC.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaCC.RecordCount
    vProgress = 0
    Do While Not rsAchaCC.EOF
        Principal.ProgressBar1.Value = vProgress
        vIdParada = rsAchaCC.Fields(1)
        SqlSomaTempoCC = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoparada as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoparada as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoparada as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a " 'where a.idparada = '" & vIdParada & "'"
        rsSomaTempoCC.Open SqlSomaTempoCC, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalCC = "Update tbApropriaControle set tempototal = '" & rsSomaTempoCC.Fields(0) & "'"
        rsInsereTempoTotalCC.Open SqlInsereTempoTotalCC, cnBanco
                                 
        rsSomaTempoCC.Close
        rsAchaCC.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando percentual por parada..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    Set rsSomaTempoCC = Nothing
    
    '5ª Parte - PERCENTUAL TOTAL POR PARADA
    SqlAchaCC = "select centrocusto,idparada from tbApropriaControle group by centrocusto,idparada order by idparada,centrocusto"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaCC.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaCC.RecordCount
    vProgress = 0
    Do While Not rsAchaCC.EOF
        Principal.ProgressBar1.Value = vProgress
        vIdParada = rsAchaCC.Fields(1)
        SqlSomaTempoCC = "select  case when a.tempototalparada <> '0000:00' then (cast(dbo.FN_CONVHORA(REPLACE(a.tempototalparada,':',':')) AS money)*100/cast(dbo.FN_CONVHORA(REPLACE(a.tempototal,':',':')) as money)) else '-' end as percentualtotalparada " & _
                         "from tbApropriaControle as a where a.idparada = '" & vIdParada & "'"
        rsSomaTempoCC.Open SqlSomaTempoCC, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalCC = "Update tbApropriaControle set percentualtotalparada = '" & rsSomaTempoCC.Fields(0) & "' where idparada = '" & vIdParada & "'"
        rsInsereTempoTotalCC.Open SqlInsereTempoTotalCC, cnBanco
                                 
        rsSomaTempoCC.Close
        rsAchaCC.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    Set rsSomaTempoCC = Nothing
End Sub

Private Sub somaTemposCSRetrabalho()
    Dim rsAchaHA As New ADODB.Recordset
    Dim SqlAchaHA As String
    Dim rsSomaTempoHA As New ADODB.Recordset
    Dim SqlSomaTempoHA As String
    
    Dim rsInsereTempoTotalHA As New ADODB.Recordset
    Dim SqlInsereTempoTotalHA As String
    
    
    Dim vCentroCusto As String
    Dim vIdParada As Integer
    Dim vRetrabalho As String
    
    '1ª Parte - ENCONTRA TEMPO TOTAL DE HORAS APROPRIADAS SEM RETRABALHO POR CENTRO DE CUSTO
    SqlAchaHA = "select centrocusto from tbApropriaControle where substring(centrocusto,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                "and retrabalho = 'S.Ret' group by centrocusto order by centrocusto"
    rsAchaHA.Open SqlAchaHA, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaHA.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaHA.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Calculando apropriação S. Ret. por CC..."
    Do While Not rsAchaHA.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaHA.Fields(0)
        SqlSomaTempoHA = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoApropriado as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.retrabalho = 'S.Ret' and a.centrocusto = '" & vCentroCusto & "'"
        rsSomaTempoHA.Open SqlSomaTempoHA, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoSRetrabalho = '" & rsSomaTempoHA.Fields(0) & "' where retrabalho = 'S.Ret' and centrocusto = '" & vCentroCusto & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
                                 
        rsSomaTempoHA.Close
        rsAchaHA.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando retrabalho por CC..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaHA.Close
    Set rsAchaHA = Nothing
    Set rsSomaTempoHA = Nothing

    '2ª Parte - ENCONTRA TEMPO TOTAL DE HORAS APROPRIADAS RETRABALHADAS POR CENTRO DE CUSTO
    SqlAchaHA = "select centrocusto from tbApropriaControle where substring(centrocusto,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                "and retrabalho = 'C.Ret' group by centrocusto order by centrocusto"
    rsAchaHA.Open SqlAchaHA, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaHA.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaHA.RecordCount
    vProgress = 0
    Do While Not rsAchaHA.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaHA.Fields(0)
        SqlSomaTempoHA = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoApropriado as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.retrabalho = 'C.Ret' and a.centrocusto = '" & vCentroCusto & "'"
        rsSomaTempoHA.Open SqlSomaTempoHA, cnBanco, adOpenKeyset, adLockReadOnly
                                 
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoSRetrabalho = '" & rsSomaTempoHA.Fields(0) & "' where retrabalho = 'C.Ret' and centrocusto = '" & vCentroCusto & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
                                 
        rsSomaTempoHA.Close
        rsAchaHA.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo total..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaHA.Close
    Set rsAchaHA = Nothing
    Set rsSomaTempoHA = Nothing
    
    
    '3ª Parte - SOMA O TEMPO SEM RETRABALHO + TEMPO COM RETRABALHO
    SqlAchaHA = "select centrocusto from tbApropriaControle where substring(centrocusto,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                "group by centrocusto order by centrocusto"
    rsAchaHA.Open SqlAchaHA, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaHA.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaHA.RecordCount
    vProgress = 0
    Do While Not rsAchaHA.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = rsAchaHA.Fields(0)
        SqlSomaTempoHA = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoApropriado as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.centrocusto = '" & vCentroCusto & "'"
        rsSomaTempoHA.Open SqlSomaTempoHA, cnBanco, adOpenKeyset, adLockReadOnly
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoCSRetrabalho = '" & rsSomaTempoHA.Fields(0) & "' where centrocusto = '" & vCentroCusto & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
        rsSomaTempoHA.Close
        rsAchaHA.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo total..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaHA.Close
    Set rsAchaHA = Nothing
    Set rsSomaTempoHA = Nothing
    
    '4ª Parte - ENCONTRA TEMPO TOTAL DAS COLUNAS DE APROPRIAÇÃO
    SqlAchaHA = "select retrabalho from tbApropriaControle where substring(centrocusto,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                "group by retrabalho order by retrabalho"
    rsAchaHA.Open SqlAchaHA, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaHA.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaHA.RecordCount
    vProgress = 0
    Do While Not rsAchaHA.EOF
        Principal.ProgressBar1.Value = vProgress
        If Not IsNull(rsAchaHA.Fields(0)) Then vRetrabalho = rsAchaHA.Fields(0)
        SqlSomaTempoHA = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoApropriado as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where a.retrabalho = '" & vRetrabalho & "'"
        rsSomaTempoHA.Open SqlSomaTempoHA, cnBanco, adOpenKeyset, adLockReadOnly
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoTotalApropriacao = '" & rsSomaTempoHA.Fields(0) & "' where retrabalho = '" & vRetrabalho & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
        rsSomaTempoHA.Close
        rsAchaHA.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando tempo total de apropriação..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaHA.Close
    Set rsAchaHA = Nothing
    Set rsSomaTempoHA = Nothing
    
    '5ª Parte - ENCONTRA TEMPO TOTAL DE APROPRIAÇÃO (SEM E COM RETRABALHO)
    'SqlAchaHA = "select retrabalho from tbApropriaControle group by retrabalho order by retrabalho"
    'rsAchaHA.Open SqlAchaHA, cnBanco, adOpenKeyset, adLockReadOnly
    
    'Do While Not rsAchaHA.EOF
        'vRetrabalho = rsAchaHA.Fields(0)
        SqlSomaTempoHA = "Set datefirst 1 select right('0000' + rtrim(CONVERT(VARCHAR, Sum(datepart(hh,Cast(a.tempoApropriado as Datetime)))+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))/60)+(sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))/60)),4) + ':' + " & _
                         "right('00' + rtrim(CONVERT(VARCHAR, (sum(datepart(MINUTE,Cast(a.tempoApropriado as Datetime)))%60+(sum(datepart(SECOND,Cast(a.tempoApropriado as Datetime)))/60))%60)),2) " & _
                         "from tbApropriaControle as a where  substring(centrocusto,1,15) in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02')" 'where a.retrabalho = '" & vRetrabalho & "'"
        rsSomaTempoHA.Open SqlSomaTempoHA, cnBanco, adOpenKeyset, adLockReadOnly
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoTotalGeral = '" & rsSomaTempoHA.Fields(0) & "'" ' where retrabalho = '" & vRetrabalho & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
        rsSomaTempoHA.Close
        'rsAchaHA.MoveNext
    'Loop
    'rsAchaHA.Close
    'Set rsAchaHA = Nothing
    Set rsSomaTempoHA = Nothing
End Sub

Private Sub somaTemposPlanejadoCC()
    'On Error Resume Next
    Dim rsAchaPlanejado As New ADODB.Recordset
    Dim SqlAchaPlanejado As String
    Dim rsSomaTempoPlanejado As New ADODB.Recordset
    Dim SqlSomaTempoPlanejado As String
    
    Dim rsSomaTempoCarteira As New ADODB.Recordset
    Dim SqlSomaTempoCarteira As String
    
    Dim rsSomaTempoGeralCarteira As New ADODB.Recordset
    Dim SqlSomaTempoGeralCarteira As String
    
    
    Dim rsBaixaParcial As New ADODB.Recordset
    Dim SqlBaixaParcial As String
    
    
    Dim rsInsereTempoTotalHA As New ADODB.Recordset
    Dim SqlInsereTempoTotalHA As String
    
    Dim rsInsereTempoCarteira As New ADODB.Recordset
    Dim SqlInsereTempoCarteira As String
    Dim vCentroCusto As String
    Dim vSomaCarteiraCC As String
    
    '1ª Parte - ENCONTRA HORAS ORÇADAS DENTRO DO PERÍODO INFORMADO POR CENTRO DE CUSTO
    SqlAchaPlanejado = "select codreduzido from tbFormula where codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                       "group by codreduzido order by codreduzido"
    rsAchaPlanejado.Open SqlAchaPlanejado, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaPlanejado.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaPlanejado.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Calculando horas orçadas por CC..."
    Do While Not rsAchaPlanejado.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsAchaPlanejado.Fields(0), 1, 15)
        
        SqlSomaTempoPlanejado = "Declare @TempoTotal as VARCHAR(40) SET @TempoTotal = '' " & _
                                "SELECT @TempoTotal = dbo.FN_CONVMIN(sum((cast(replace(b.tempocalc,'.','') as money)/100))) from tbMP as a inner join tbMPItens as b on a.idprogramacao = b.idprogramacao " & _
                                "where a.dataprogramacao  BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and b.idcc = '" & vCentroCusto & "' select @TempoTotal as Tempo_Total"
        rsSomaTempoPlanejado.Open SqlSomaTempoPlanejado, cnBanco, adOpenKeyset, adLockReadOnly
        
        SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoPlanejadoCC = '" & rsSomaTempoPlanejado.Fields(0) & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
                                 
        rsSomaTempoPlanejado.Close
        rsAchaPlanejado.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando total de horas do período..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    Set rsSomaTempoPlanejado = Nothing
    
    '2ª Parte - ENCONTRA TOTAL DE HORAS ORÇADAS DENTRO DO PERÍODO INFORMADO
    SqlSomaTempoPlanejado = "Declare @TempoTotal as VARCHAR(40) SET @TempoTotal = '' " & _
                            "SELECT @TempoTotal = dbo.FN_CONVMIN(sum((cast(replace(b.tempocalc,'.','') as money)/100))) from tbMP as a inner join tbMPItens as b on a.idprogramacao = b.idprogramacao " & _
                            "where a.dataprogramacao  BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and  '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' and substring(b.idcc,1,4) = '3000' select @TempoTotal as Tempo_Total"
    rsSomaTempoPlanejado.Open SqlSomaTempoPlanejado, cnBanco, adOpenKeyset, adLockReadOnly
        
    SqlInsereTempoTotalHA = "Update tbApropriaControle set TempoPlanejadoTotal = '" & rsSomaTempoPlanejado.Fields(0) & "'"
    rsInsereTempoTotalHA.Open SqlInsereTempoTotalHA, cnBanco
    
    
    '3ª Parte - ENCONTRA HORAS TOTAIS QUE ESTAO COM STATUS DIFERENTE DE 3 (FECHADO) POR CENTRO DE CUSTO
    '           INDEPENDENTE DO PERÍODO
    rsAchaPlanejado.MoveFirst
    Dim VTESTE As Double
    
    If rsAchaPlanejado.RecordCount > 0 Then Principal.ProgressBar1.Max = rsAchaPlanejado.RecordCount
    vProgress = 0
    Do While Not rsAchaPlanejado.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsAchaPlanejado.Fields(0), 1, 15)
        vSomaCarteiraCC = "00:00"
        VTESTE = 0

        SqlSomaTempoCarteira = "select a.idcc,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,a.status,(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) from tbMPItens as a " & _
                               "WHERE a.idcc in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                               "and a.idcc = '" & vCentroCusto & "' and a.dataprevista is not null order by a.idcc,a.idos,a.idoperacao,a.dataprevista"
        rsSomaTempoCarteira.Open SqlSomaTempoCarteira, cnBanco, adOpenKeyset, adLockReadOnly
        
        
        Do While Not rsSomaTempoCarteira.EOF
            If rsSomaTempoCarteira.Fields(1) <> " " And rsSomaTempoCarteira.Fields(2) <> 3 Then
                    somaTempoPPSAtraso rsSomaTempoCarteira.Fields(1), vSomaCarteiraCC 'Horas tempo de carteira por CC
                    VTESTE = VTESTE + rsSomaTempoCarteira.Fields(3)
            End If
            rsSomaTempoCarteira.MoveNext
        Loop
        
        
'-------ABATE HORAS DE BAIXA PARCIAL (CARTEIRA)
        SqlBaixaParcial = "select b.idcc,(sum((cast(replace(replace(b.tempocalc,'.',''),',','.') as money) * a.percentualBaixado)/100)) as TBaixa_Parcial from tbMPBaixaParcial as a inner join tbMPItens as b on a.idos = b.idos and a.revisao = b.revisaoos and a.idoperacao = b.idoperacao " & _
                          "where b.idcc = '" & vCentroCusto & "' and b.status <> 3 and a.percentualBaixado is not null and b.idcc in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03') group by b.idcc"
        rsBaixaParcial.Open SqlBaixaParcial, cnBanco, adOpenKeyset, adLockReadOnly
        If rsBaixaParcial.RecordCount > 0 Then
            Dim rsAbateBaixa As New ADODB.Recordset
            Dim SqlAbateBaixa As String
            SqlAbateBaixa = "Select dbo.FN_CONVMIN(cast(replace(replace('" & VTESTE - rsBaixaParcial.Fields(1) & "','.',''),',','.') as money)) as Tempo_Convertido"
            rsAbateBaixa.Open SqlAbateBaixa, cnBanco, adOpenKeyset, adLockReadOnly
            vSomaCarteiraCC = rsAbateBaixa.Fields(0)
            rsAbateBaixa.Close
        End If
        rsBaixaParcial.Close
'-------------------------
        
        
        SqlInsereTempoCarteira = "Update tbApropriaControle set TempoTotalCarteira = '" & vSomaCarteiraCC & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoCarteira.Open SqlInsereTempoCarteira, cnBanco
                                 
        rsSomaTempoCarteira.Close
        rsAchaPlanejado.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando total de horas do período..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    Set rsSomaTempoCarteira = Nothing
    
    
    '4ª Parte - ENCONTRA HORAS TOTAIS QUE ESTAO COM STATUS DIFERENTE DE 3 (FECHADO) INDEPENDENTE DO CENTRO DE CUSTO E
    '           INDEPENDENTE DO PERÍODO
    
'-------ABATE HORAS DE BAIXA PARCIAL DO TOTAL GERAL (CARTEIRA)
        'Dim VTESTE1 As Double
        VTESTE = 0
        SqlBaixaParcial = "select (sum((cast(replace(replace(b.tempocalc,'.',''),',','.') as money) * a.percentualBaixado)/100)) as TBaixa_Parcial from tbMPBaixaParcial as a inner join tbMPItens as b on a.idos = b.idos and a.revisao = b.revisaoos and a.idoperacao = b.idoperacao " & _
                          "where b.status <> 3 and a.percentualBaixado is not null and b.idcc in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03')"
        rsBaixaParcial.Open SqlBaixaParcial, cnBanco, adOpenKeyset, adLockReadOnly
        If rsBaixaParcial.RecordCount > 0 Then
            If Not IsNull(rsBaixaParcial.Fields(0)) Then VTESTE = rsBaixaParcial.Fields(0) Else VTESTE = 0
        End If
        rsBaixaParcial.Close
'-------------------------
'"SELECT @TempoGeralCarteira = dbo.FN_CONVMIN((sum(((cast(replace(a.tempocalc,'.','') as money))/100)))- 100000.544)  from tbMPItens as a " & _

    
    SqlSomaTempoGeralCarteira = "Declare @TempoGeralCarteira as VARCHAR(40) SET @TempoGeralCarteira = '' " & _
                            "SELECT @TempoGeralCarteira = dbo.FN_CONVMIN((sum(((cast(replace(a.tempocalc,'.','') as money))/100)))- " & Replace(VTESTE, ",", ".") & ")  from tbMPItens as a " & _
                            "where a.status <> 3 and a.tempocalc <> ' ' and a.tempocalc <> '0' and a.dataprevista is not null select @TempoGeralCarteira as Tempo_GeralCarteira"
    rsSomaTempoGeralCarteira.Open SqlSomaTempoGeralCarteira, cnBanco, adOpenKeyset, adLockReadOnly
    
    
    
    
    SqlInsereTempoCarteira = "Update tbApropriaControle set TempoGeralCarteira = '" & rsSomaTempoGeralCarteira.Fields(0) & "'"
    rsInsereTempoCarteira.Open SqlInsereTempoCarteira, cnBanco
    
    rsSomaTempoPlanejado.Close
    rsAchaPlanejado.Close
    Set rsAchaPlanejado = Nothing
    Set rsSomaTempoPlanejado = Nothing
End Sub

Private Sub somaTemposProgramadosCC()
    'On Error Resume Next
    Dim rsSelecionaCCs As New ADODB.Recordset
    Dim SqlSelecionaCCs As String
    
    Dim rsAchaProgramados As New ADODB.Recordset
    Dim SqlAchaProgramados As String
    
    Dim rsAchaAtraso As New ADODB.Recordset
    Dim SqlAchaAtraso As String
    
    Dim rsSomaTempoPPSAtraso As New ADODB.Recordset
    Dim SqlSomaTempoPPSAtraso As String
    
    Dim rsSomaExtraPPSCC As New ADODB.Recordset
    Dim SqlSomaExtraPPSCC As String
    
    
    Dim rsInsereTempoProgramados As New ADODB.Recordset
    Dim SqlInsereTempoProgramados As String
    Dim rsInsereTempoAtraso As New ADODB.Recordset
    Dim SqlInsereTempoAtraso As String
    
    Dim vCentroCusto As String
    Dim vHorasPPSCC As String
    Dim vHorasAtrasoCC As String
    Dim vHorasPPSTotal As String
    Dim vHorasAtrasoTotal As String
    Dim vSomaHorasAtrasoPPS As String
    Dim vSomaHorasAtrasoPPSTotal As String
    Dim vHorasPPSRealCC As String
    Dim vHorasPPSRealTotal As String
    Dim vHorasExtraPPS As String
    Dim vHorasExtraPPSRealTotal As String
    Dim vSomaExtraPPSReal As String
    Dim vTempoTotalRealizado As String
    Dim vSemanaBaixada As Integer 'Armazena semana atual
    
    '1ª Parte - SELECIONA OS CENTROS DE CUSTO OS QUAIS IREMOS TRABALHAR
    SqlSelecionaCCs = "select codreduzido from tbFormula where codreduzido in('3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3103.SC-01','3000.3103.SC-02','3000.3104.SC-01','3000.3104.SC-02','3000.3105.SC-01','3000.3105.SC-02','3000.3106.SC-01','4000.4101.SC-03','7000.7103.SC-02') " & _
                       "group by codreduzido order by codreduzido"
    rsSelecionaCCs.Open SqlSelecionaCCs, cnBanco, adOpenKeyset, adLockReadOnly
    
    '2ª Parte - SOMA AS HORAS DE PPS POR CENTRO DE CUSTO DENTRO DO PERIODO INFORMADO
    SqlAchaProgramados = "Set datefirst 1 select a.idcc,CONVERT (VARCHAR,a.dataprevista, 103) as Data_Programacao,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,max(DATEPART(WK,b.dataprogramacao)) as SemanaPlanejada,max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,case when max(DATEPART(WK,a.databaixa)) is null then 0 else max(DATEPART(WK,a.databaixa)) end as SemanaBaixada," & _
                         "a.idos,max(a.idoperacao) as operacao,MAX(a.status) as status,max(d.fce) as fce,max(e.idretrabalho) as retrabalho from tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto and d.fce > 2000 left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                         "where A.dataprevista BETWEEN '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and '" & Format(DTPicker2.Value, "yyyy/mm/dd") & "' group by a.idcc,a.dataprevista,a.tempocalc,b.dataprogramacao,a.idos,a.idoperacao,a.status,d.fce,e.idretrabalho order by a.idcc,a.idos,a.idoperacao"
    cnBanco.CommandTimeout = 0
    rsAchaProgramados.Open SqlAchaProgramados, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSelecionaCCs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSelecionaCCs.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Calculando Horas de PPS por CC"
    Do While Not rsSelecionaCCs.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsSelecionaCCs.Fields(0), 1, 15)
        vHorasPPSCC = "00:00"
        vHorasPPSRealCC = "00:00"
        If rsAchaProgramados.RecordCount > 0 Then rsAchaProgramados.MoveFirst
        Do While Not rsAchaProgramados.EOF
            If vCentroCusto = rsAchaProgramados.Fields(0) Then
                Do While vCentroCusto = rsAchaProgramados.Fields(0) And Not rsAchaProgramados.EOF
                    
                    'A CONDIÇÃO ABAIXO É VÁLIDA APENAS PARA NO CASO DA SEMANA SOLICITADA SEJA IGUAL A SEMANA ATUAL. CONSIDERA SEMANA ATUAL CASO A SEMANA DA BAIXA SEJA ZERO
                    If rsAchaProgramados.Fields(5) = 0 And Val(Text2.Text) = Val(DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays)) Then
                        vSemanaBaixada = DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays) 'DatePart(WK, Date, vbMonday, vbFirstFourDays)
                    Else
                        vSemanaBaixada = rsAchaProgramados.Fields(5)
                    End If
                    '------------------------------------------------------------------------------------------------------
                    
                    'O IF abaixo passa por 3 condições que devem ser diferentes de zero. Se entrar no IF é PPS
                    If Val(Text2.Text) - Val(rsAchaProgramados.Fields(3)) <> 0 Or Val(rsAchaProgramados.Fields(4)) - Val(Text2.Text) <> 0 Or vSemanaBaixada - Val(Text2.Text) <> 0 Then
                        somaTempoPPSAtraso rsAchaProgramados.Fields(2), vHorasPPSCC 'Horas totais de PPS por CC
                        'If rsAchaProgramados.Fields(8) = 3 Then
                        '    If rsAchaProgramados.Fields(2) <> " " Then somaTempoPPSAtraso rsAchaProgramados.Fields(2), vHorasPPSRealCC 'Horas totais de PPS Realizado por CC
                        'End If
                        
                    End If
                    rsAchaProgramados.MoveNext
                    If rsAchaProgramados.EOF Then Exit Do
                Loop
                Exit Do
            Else
                rsAchaProgramados.MoveNext
            End If
        Loop
        SqlInsereTempoProgramados = "Update tbApropriaControle set PPSPorCC = '" & vHorasPPSCC & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        
        SqlInsereTempoProgramados = "Update tbApropriaControle set PPSRealPorCC = '" & vHorasPPSRealCC & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        
        rsSelecionaCCs.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando horas de Atraso por CC"
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaProgramados.Close
    Set rsAchaProgramados = Nothing
    
    '3ª Parte - SOMA AS HORAS DE ATRASO POR CENTRO DE CUSTO DENTRO DO PERIODO INFORMADO
    rsSelecionaCCs.MoveFirst
    SqlAchaAtraso = "Set datefirst 1 SELECT a.idcc,CONVERT (VARCHAR,a.dataprevista, 103) as Data_Programacao,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,max(DATEPART(WK,b.dataprogramacao)) as SemanaPlanejada,max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,case when max(DATEPART(WK,a.databaixa)) is null then 0 else max(DATEPART(WK,a.databaixa)) end as SemanaBaixada," & _
                    "a.idos,max(a.idoperacao) as operacao,a.status,max(d.fce) as fce,max(e.idretrabalho) as retrabalho FROM tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto and d.fce > 2000 left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                    "where A.dataprevista < '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and DATEPART(WK,GETDATE()) > DATEPART(WK,a.dataprevista) group by a.idcc,a.dataprevista,a.tempocalc,b.dataprogramacao,a.idos,a.idoperacao,a.status order by a.idcc,a.idos,a.idoperacao"
    cnBanco.CommandTimeout = 0
    rsAchaAtraso.Open SqlAchaAtraso, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSelecionaCCs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSelecionaCCs.RecordCount
    vProgress = 0
    Do While Not rsSelecionaCCs.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsSelecionaCCs.Fields(0), 1, 15)
        vHorasAtrasoCC = "00:00"
        vHorasPPSRealCC = "00:00"
        If rsAchaAtraso.RecordCount > 0 Then rsAchaAtraso.MoveFirst
        Do While Not rsAchaAtraso.EOF
            If vCentroCusto = rsAchaAtraso.Fields(0) Then
                Do While vCentroCusto = rsAchaAtraso.Fields(0) And Not rsAchaAtraso.EOF
                    
                    'A CONDIÇÃO ABAIXO É VÁLIDA APENAS PARA NO CASO DA SEMANA SOLICITADA SEJA IGUAL A SEMANA ATUAL. CONSIDERA SEMANA ATUAL CASO A SEMANA DA BAIXA SEJA ZERO
                    If rsAchaAtraso.Fields(5) = 0 And Val(Text2.Text) <> Val(DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays)) Then
                        vSemanaBaixada = DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays) 'DatePart(WK, Date)
                    Else
                        vSemanaBaixada = rsAchaAtraso.Fields(5)
                    End If
                    '------------------------------------------------------------------------------------------------------
                    
                    'INCLUIR FILTRO PARA ATRASO E ATRASO REALIZADO
                    'Realizado ----------------------------------------------
                    If (Val(rsAchaAtraso.Fields(3)) <> Val(Text2.Text) Or Val(rsAchaAtraso.Fields(4)) <> Val(Text2.Text) Or vSemanaBaixada <> Val(Text2.Text)) And vSemanaBaixada = Val(Text2.Text) Then
                        If rsAchaAtraso.Fields(2) <> " " Then somaTempoPPSAtraso rsAchaAtraso.Fields(2), vHorasPPSRealCC 'Horas totais de PPS Realizado por CC
                    End If
                    '--------------------------------------------------------
                    
                    If rsAchaAtraso.Fields(5) = 0 And Val(Text2.Text) <> Val(DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays)) Then
                        vSemanaBaixada = DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays) 'DatePart(WK, Date)
                    ElseIf rsAchaAtraso.Fields(5) = 0 And Val(Text2.Text) = Val(DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays)) Then
                        vSemanaBaixada = DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays) 'DatePart(WK, Date)
                    Else
                        vSemanaBaixada = rsAchaAtraso.Fields(5)
                    End If
                    
                    'If rsAchaAtraso.Fields(5) = 0 And Val(Text2.Text) = Val(DatePart("ww", CDate(Date))) Then
                    '    vSemanaBaixada = DatePart("ww", CDate(Date)) 'DatePart(WK, Date)
                    'Else
                    '    vSemanaBaixada = rsAchaAtraso.Fields(5)
                    'End If
                    
                    'Atraso -------------------------------------------------
                    If Val(rsAchaAtraso.Fields(4)) <= Val(Text2.Text) Then
                        If vSemanaBaixada >= Val(Text2.Text) Then
                            If rsAchaAtraso.Fields(2) <> " " Then somaTempoPPSAtraso rsAchaAtraso.Fields(2), vHorasAtrasoCC
                        Else
                            If rsAchaAtraso.Fields(8) < 3 And Val(Text2.Text) = vSemanaBaixada Then
                                If rsAchaAtraso.Fields(2) <> " " Then somaTempoPPSAtraso rsAchaAtraso.Fields(2), vHorasAtrasoCC
                            End If
                        End If
                    End If
                    '----------------------------------------------------------
                    
                    rsAchaAtraso.MoveNext
                    If rsAchaAtraso.EOF Then Exit Do
                Loop
                Exit Do
            Else
                rsAchaAtraso.MoveNext
            End If
        Loop
        SqlInsereTempoAtraso = "Update tbApropriaControle set AtrasoPorCC = '" & vHorasAtrasoCC & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoAtraso.Open SqlInsereTempoAtraso, cnBanco
        
        'TESTE
        Dim rsAchaHorasCC As New ADODB.Recordset
        Dim SqlAchaHorasCC As String
        
        SqlAchaHorasCC = "Select a.PPSRealPorCC from tbApropriaControle as a where substring(a.centrocusto,1,15) = '" & vCentroCusto & "'"
        rsAchaHorasCC.Open SqlAchaHorasCC, cnBanco, adOpenKeyset, adLockReadOnly
        If rsAchaHorasCC.RecordCount > 0 Then somaTempoPPSAtraso rsAchaHorasCC.Fields(0), vHorasPPSRealCC
        rsAchaHorasCC.Close
        Set rsAchaHorasCC = Nothing
        
        If vHorasPPSRealCC <> "00:00" Then
            SqlInsereTempoProgramados = "Update tbApropriaControle set PPSRealPorCC = '" & vHorasPPSRealCC & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
            rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        End If
        'TESTE
        
        rsSelecionaCCs.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando PPS + Atraso por CC"
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsAchaAtraso.Close
    Set rsAchaAtraso = Nothing
    
    '4ª Parte - SOMA TODAS AS HORAS DE ATRASO E TODAS AS HORAS DE PPS INDEPENDENTE DO CENTRO DE CUSTO
    rsSelecionaCCs.MoveFirst
    SqlSomaTempoPPSAtraso = "select a.centrocusto,a.ppsporcc,a.atrasoporcc,a.PPSRealPorCC from tbApropriaControle as a group by a.centrocusto,a.ppsporcc,a.atrasoporcc,a.PPSRealPorCC"


    cnBanco.CommandTimeout = 0
    rsSomaTempoPPSAtraso.Open SqlSomaTempoPPSAtraso, cnBanco, adOpenKeyset, adLockReadOnly

    vHorasPPSTotal = "00:00"
    vHorasAtrasoTotal = "00:00"
    vHorasPPSRealTotal = "00:00"
    If rsSomaTempoPPSAtraso.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSomaTempoPPSAtraso.RecordCount
    vProgress = 0
    Do While Not rsSomaTempoPPSAtraso.EOF
        Principal.ProgressBar1.Value = vProgress
        If Not IsNull(rsSomaTempoPPSAtraso.Fields(1)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(1), vHorasPPSTotal
        If Not IsNull(rsSomaTempoPPSAtraso.Fields(2)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(2), vHorasAtrasoTotal
        If Not IsNull(rsSomaTempoPPSAtraso.Fields(3)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(3), vHorasPPSRealTotal

        rsSomaTempoPPSAtraso.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando total de horas de programação"
    Principal.StatusBar1.Panels(3).Text = Legenda

    SqlInsereTempoAtraso = "Update tbApropriaControle set PPSTotal = '" & vHorasPPSTotal & "'"
    rsInsereTempoAtraso.Open SqlInsereTempoAtraso, cnBanco

    SqlInsereTempoAtraso = "Update tbApropriaControle set AtrasoTotal = '" & vHorasAtrasoTotal & "'"
    rsInsereTempoAtraso.Open SqlInsereTempoAtraso, cnBanco

    SqlInsereTempoAtraso = "Update tbApropriaControle set PPSRealTotalPorCC = '" & vHorasPPSRealTotal & "'"
    rsInsereTempoAtraso.Open SqlInsereTempoAtraso, cnBanco

    rsSomaTempoPPSAtraso.Close
    Set rsSomaTempoPPSAtraso = Nothing
    
    '5ª Parte - SOMA HORAS DE ATRASO + HORAS DE PPS POR CENTRO DE CUSTO
    rsSelecionaCCs.MoveFirst
    SqlSomaTempoPPSAtraso = "select a.centrocusto,a.ppsporcc,a.atrasoporcc from tbApropriaControle as a group by a.centrocusto,a.ppsporcc,a.atrasoporcc"
    cnBanco.CommandTimeout = 0
    rsSomaTempoPPSAtraso.Open SqlSomaTempoPPSAtraso, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSelecionaCCs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSelecionaCCs.RecordCount
    vProgress = 0
    Do While Not rsSelecionaCCs.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsSelecionaCCs.Fields(0), 1, 15)
        vSomaHorasAtrasoPPS = "00:00"
        If rsSomaTempoPPSAtraso.RecordCount > 0 Then rsSomaTempoPPSAtraso.MoveFirst
        Do While Not rsSomaTempoPPSAtraso.EOF
            If vCentroCusto = Mid$(rsSomaTempoPPSAtraso.Fields(0), 1, 15) Then
                Do While vCentroCusto = Mid$(rsSomaTempoPPSAtraso.Fields(0), 1, 15)
                    If rsSomaTempoPPSAtraso.Fields(1) <> " " Or Not IsNull(rsSomaTempoPPSAtraso.Fields(1)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(1), vSomaHorasAtrasoPPS
                    If rsSomaTempoPPSAtraso.Fields(2) <> " " Or Not IsNull(rsSomaTempoPPSAtraso.Fields(2)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(2), vSomaHorasAtrasoPPS
                    rsSomaTempoPPSAtraso.MoveNext
                    If rsSomaTempoPPSAtraso.EOF Then Exit Do
                Loop
                Exit Do
            Else
                rsSomaTempoPPSAtraso.MoveNext
            End If
        Loop
        
        SqlInsereTempoProgramados = "Update tbApropriaControle set PPSeAtrasoPorCC = '" & vSomaHorasAtrasoPPS & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        rsSelecionaCCs.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando total de horas de programação..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsSomaTempoPPSAtraso.Close
    Set rsSomaTempoPPSAtraso = Nothing
    
    '6ª Parte - SOMA HORAS DE ATRASO + HORAS DE PPS GERAL
    rsSelecionaCCs.MoveFirst
    SqlSomaTempoPPSAtraso = "select a.centrocusto,a.PPSeAtrasoPorCC from tbApropriaControle as a group by a.centrocusto,a.PPSeAtrasoPorCC"
    cnBanco.CommandTimeout = 0
    rsSomaTempoPPSAtraso.Open SqlSomaTempoPPSAtraso, cnBanco, adOpenKeyset, adLockReadOnly
    vSomaHorasAtrasoPPSTotal = "00:00"
    
    If rsSelecionaCCs.RecordCount > 0 And rsSomaTempoPPSAtraso.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSomaTempoPPSAtraso.RecordCount
    vProgress = 0
    Do While Not rsSomaTempoPPSAtraso.EOF
        Principal.ProgressBar1.Value = vProgress
        If Not IsNull(rsSomaTempoPPSAtraso.Fields(1)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(1), vSomaHorasAtrasoPPSTotal
        rsSomaTempoPPSAtraso.MoveNext
        vProgress = vProgress + 1
    Loop
    SqlInsereTempoAtraso = "Update tbApropriaControle set PPSeAtrasoSoma = '" & vSomaHorasAtrasoPPSTotal & "'"
    rsInsereTempoAtraso.Open SqlInsereTempoAtraso, cnBanco
    
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando Extra PPS por Centro de CUsto"
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsSomaTempoPPSAtraso.Close
    Set rsSomaTempoPPSAtraso = Nothing
    
    '7ª Parte - EXTRA PPS POR CENTRO DE CUSTO
    SqlSomaExtraPPSCC = "Set datefirst 1 select a.idcc,CONVERT (VARCHAR,a.dataprevista, 103) as Data_Programacao,dbo.FN_CONVMIN(cast(replace(replace(a.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido,max(DATEPART(WK,b.dataprogramacao)) as SemanaPlanejada,max(DATEPART(WK,a.dataprevista)) as SemanaProgramada,case when max(DATEPART(WK,a.databaixa)) is null then DATEPART(WK,GETDATE()) else max(DATEPART(WK,a.databaixa)) end as SemanaBaixada," & _
                        "a.idos,max(a.idoperacao) as operacao,MAX(a.status) as status,max(d.fce) as fce,max(e.idretrabalho) as retrabalho from tbMPItens as a Inner join tbMP as b on a.idprogramacao=b.idprogramacao inner join tbProjetos as d on b.codprojeto = d.codprojeto and d.fce > 2000 left join tbRetrabalho as e on b.idprogramacao = e.idprogramacao " & _
                        "where A.dataprevista < '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "' and DATEPART(WK,GETDATE()) > DATEPART(WK,a.dataprevista) group by a.idcc,a.dataprevista,a.tempocalc,b.dataprogramacao,a.idos,a.idoperacao,a.status,d.fce,e.idretrabalho order by a.idcc,a.idos,a.idoperacao,a.dataprevista"
    cnBanco.CommandTimeout = 0
    rsSomaExtraPPSCC.Open SqlSomaExtraPPSCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSelecionaCCs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSelecionaCCs.RecordCount
    vProgress = 0
    Principal.StatusBar1.Panels(3).Text = "Calculando Horas de Extra PPS por CC"
    vHorasExtraPPSRealTotal = "00:00"
    Do While Not rsSelecionaCCs.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsSelecionaCCs.Fields(0), 1, 15)
        vHorasExtraPPS = "00:00"
        If rsSomaExtraPPSCC.RecordCount > 0 Then rsSomaExtraPPSCC.MoveFirst
        Do While Not rsSomaExtraPPSCC.EOF
            If vCentroCusto = rsSomaExtraPPSCC.Fields(0) Then
                Do While vCentroCusto = rsSomaExtraPPSCC.Fields(0) And Not rsSomaExtraPPSCC.EOF
                    If rsSomaExtraPPSCC.Fields(2) <> " " Then
                        If Val(rsSomaExtraPPSCC.Fields(3)) = Val(Text2.Text) And Val(rsSomaExtraPPSCC.Fields(4)) = Val(Text2.Text) And Val(rsSomaExtraPPSCC.Fields(5)) = Val(Text2.Text) Then
                        'If (Val(Text2.Text) * 4) - Val(Text2.Text) - rsSomaExtraPPSCC.Fields(5) - rsSomaExtraPPSCC.Fields(4) - rsSomaExtraPPSCC.Fields(3) Or rsSomaExtraPPSCC.Fields(5) - Val(Text2.Text) > 0 And rsSomaExtraPPSCC.Fields(5) - rsSomaExtraPPSCC.Fields(4) < 0 Then
                            somaTempoPPSAtraso rsSomaExtraPPSCC.Fields(2), vHorasExtraPPS 'Horas totais de PPS por CC
                            somaTempoPPSAtraso rsSomaExtraPPSCC.Fields(2), vHorasExtraPPSRealTotal 'Horas totais de PPS
                        End If
                    End If
                    rsSomaExtraPPSCC.MoveNext
                    If rsSomaExtraPPSCC.EOF Then Exit Do
                Loop
                Exit Do
            Else
                rsSomaExtraPPSCC.MoveNext
            End If
        Loop
        SqlInsereTempoProgramados = "Update tbApropriaControle set ExtraPPSRealPorCC = '" & vHorasExtraPPS & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        
        SqlInsereTempoProgramados = "Update tbApropriaControle set ExtraPPSRealTotalPorCC = '" & vHorasExtraPPSRealTotal & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        
        rsSelecionaCCs.MoveNext
        vProgress = vProgress + 1
    Loop
    Principal.ProgressBar1.Value = 0
    Legenda = "Calculando horas realizadas por Centro de Custo..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    rsSomaExtraPPSCC.Close
    Set rsSomaExtraPPSCC = Nothing
    
'------------------------------------------------------------
'------------------------------------------------------------

    '8ª Parte - SOMA TODAS AS HORAS REALIZADAS + HORAS DE EXTRA PPS POR CENTRO DE CUSTO (SEMELHANTE A 5ª PARTE)
    rsSelecionaCCs.MoveFirst
    SqlSomaTempoPPSAtraso = "select a.centrocusto,a.PPSRealPorCC,a.ExtraPPSRealPorCC from tbApropriaControle as a group by a.centrocusto,a.PPSRealPorCC,a.ExtraPPSRealPorCC"
    cnBanco.CommandTimeout = 0
    rsSomaTempoPPSAtraso.Open SqlSomaTempoPPSAtraso, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsSelecionaCCs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsSelecionaCCs.RecordCount
    vProgress = 0
    vTempoTotalRealizado = "00:00"
    Do While Not rsSelecionaCCs.EOF
        Principal.ProgressBar1.Value = vProgress
        vCentroCusto = Mid$(rsSelecionaCCs.Fields(0), 1, 15)
        vSomaExtraPPSReal = "00:00"
        If rsSomaTempoPPSAtraso.RecordCount > 0 Then rsSomaTempoPPSAtraso.MoveFirst
        Do While Not rsSomaTempoPPSAtraso.EOF
            If vCentroCusto = Mid$(rsSomaTempoPPSAtraso.Fields(0), 1, 15) Then
                Do While vCentroCusto = Mid$(rsSomaTempoPPSAtraso.Fields(0), 1, 15)
                    If rsSomaTempoPPSAtraso.Fields(1) <> " " Or Not IsNull(rsSomaTempoPPSAtraso.Fields(1)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(1), vSomaExtraPPSReal
                    If rsSomaTempoPPSAtraso.Fields(2) <> " " Or Not IsNull(rsSomaTempoPPSAtraso.Fields(2)) Then somaTempoPPSAtraso rsSomaTempoPPSAtraso.Fields(2), vSomaExtraPPSReal
                    rsSomaTempoPPSAtraso.MoveNext
                    If rsSomaTempoPPSAtraso.EOF Then Exit Do
                Loop
                Exit Do
            Else
                rsSomaTempoPPSAtraso.MoveNext
            End If
        Loop
        somaTempoPPSAtraso vSomaExtraPPSReal, vTempoTotalRealizado
        
        SqlInsereTempoProgramados = "Update tbApropriaControle set ExtraPPSRealSoma = '" & vSomaExtraPPSReal & "' where substring(centrocusto,1,15) = '" & vCentroCusto & "'"
        rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
        rsSelecionaCCs.MoveNext
        vProgress = vProgress + 1
    Loop
    
    SqlInsereTempoProgramados = "Update tbApropriaControle set TempoTotalRealizado = '" & vTempoTotalRealizado & "'"
    rsInsereTempoProgramados.Open SqlInsereTempoProgramados, cnBanco
    
    Principal.ProgressBar1.Value = 0
    Legenda = vGuardaLegenda
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    rsSomaTempoPPSAtraso.Close
    Set rsSomaTempoPPSAtraso = Nothing

'------------------------------------------------------------
   
    rsSelecionaCCs.Close
    Set rsSelecionaCCs = Nothing
End Sub

Private Sub LocalizaParada()
'On Error Resume Next
    Dim vChapa As String
    Dim vNome As String
    Dim vCentroCusto As String
    Dim vDataEntrada As String
    Dim vHoraEntrada As String
    Dim vDataSaida As String
    Dim vHoraSaida As String
    Dim vIdParada As String
    Dim vNmParada As String
    Dim rsTempoParada As New ADODB.Recordset
    Dim SqlTempoParada As String
    
    Dim vDifHora As String
    
    'SqlTempoParada = "delete from tbApropriaControle"
    'rsTempoParada.Open SqlTempoParada, cnBanco
    
    rsApropriacao.MoveFirst
    Do While Not rsApropriacao.EOF
        'ListView1.ListItems.Item(X).Selected = True
        If rsApropriacao.Fields(7) <> "9014" And rsApropriacao.Fields(7) <> "9018" And rsApropriacao.Fields(7) <> "9019" Then 'And rsApropriacao.Fields(7) <> "9020" Then
            
            vChapa = rsApropriacao.Fields(0)
            vNome = rsApropriacao.Fields(1)
            vCentroCusto = rsApropriacao.Fields(2)
            vDataEntrada = rsApropriacao.Fields(5)
            vHoraEntrada = Format(rsApropriacao.Fields(6), "hh:mm")
            vIdParada = rsApropriacao.Fields(7)
            vNmParada = rsApropriacao.Fields(8)
            
            'TESTE
            If vCentroCusto <> "" And vCentroCusto <> rsApropriacao.Fields(2) Then
                incluiParadasVazias vCentroCusto, vDataEntrada, vDataSaida
            End If
            'TESTE
            
            rsApropriacao.MoveNext
            If rsApropriacao.EOF = True Then Exit Do
            
            'SE AS PARADAS ENCONTRADAS ESTIVEREM NO MESMO DIA
            
            'TESTE
            SqlTempoParada = "Select * from tbApropriaControle"
            rsTempoParada.Open SqlTempoParada, cnBanco, adOpenKeyset, adLockOptimistic
            rsTempoParada.AddNew
            'TESTE
            
            If rsApropriacao.Fields(3) = vDataEntrada And rsApropriacao.Fields(1) = vNome Then
                vDataSaida = rsApropriacao.Fields(3)
                vHoraSaida = Format(rsApropriacao.Fields(4), "hh:mm")
                
                vDifHora = Format(TimeValue(vHoraSaida) - TimeValue(vHoraEntrada), "hh:mm")
                
                rsApropriacao.MovePrevious
            Else
                rsApropriacao.MovePrevious
                vDataSaida = rsApropriacao.Fields(3)
                vHoraSaida = Format(achaHorarioSaida(vChapa), "hh:mm") 'Busca o horário de saida do colaborador que esta cadastrado na folha de pagamento
                
                vDifHora = Format(TimeValue(vHoraSaida) - TimeValue(vHoraEntrada), "hh:mm")
            End If
            rsTempoParada.Fields(0) = vChapa
            rsTempoParada.Fields(1) = vNome
            rsTempoParada.Fields(2) = vCentroCusto
            rsTempoParada.Fields(3) = vDataEntrada
            rsTempoParada.Fields(4) = vHoraEntrada
            rsTempoParada.Fields(5) = vDataSaida
            rsTempoParada.Fields(6) = vHoraSaida
            rsTempoParada.Fields(7) = vIdParada
            rsTempoParada.Fields(8) = vNmParada
            rsTempoParada.Fields(9) = "00:00" 'Tempo apropriado
            rsTempoParada.Fields(10) = "S.Ret" 'Tipo de apropriação
            rsTempoParada.Fields(11) = "0000:00"
            rsTempoParada.Fields(12) = "0000:00"
            rsTempoParada.Fields(13) = "0000:00"
            rsTempoParada.Fields(14) = "0000:00"
            
            rsTempoParada.Fields(19) = vDifHora
            
            rsTempoParada.Update
            rsTempoParada.Close
            
            If rsApropriacao.EOF = True Then Exit Sub
        End If
        rsApropriacao.MoveNext
    Loop
    'If vCentroCusto <> "" And vCentroCusto <> rsApropriacao.Fields(2) Then
        incluiParadasVazias vCentroCusto, vDataEntrada, vDataSaida
    'End If
End Sub

Private Sub incluiParadasVazias(vCC As String, vDE As String, vDS As String)
On Error Resume Next
    Dim rsTP As New ADODB.Recordset
    Dim SqlTP As String
    
    Dim rsParadas As New ADODB.Recordset
    Dim sqlParadas As String
        
    Dim vIP As String, vNP As String
    
    sqlParadas = "select a.codigo,a.nmparada from tbParadas as a where " & _
                 "a.codigo in(9001,9002,9003,9004,9005,9006,9007,9008,9009,9010,9011,9012,9013,9014,9015,9016,9017,9018,9019,9020) order by a.codigo"
    rsParadas.Open sqlParadas, cnBanco, adOpenKeyset, adLockReadOnly
    
    'TESTE
    SqlTP = "Select * from tbApropriaControle"
    rsTP.Open SqlTP, cnBanco, adOpenKeyset, adLockOptimistic
    'TESTE
    
    Do While Not rsParadas.EOF
        vIP = rsParadas.Fields(0)
        vNP = rsParadas.Fields(1)
        rsTP.AddNew
        rsTP.Fields(0) = 0
        rsTP.Fields(1) = "-"
        rsTP.Fields(2) = vCC
        rsTP.Fields(3) = vDE
        rsTP.Fields(4) = "17:00:00"
        rsTP.Fields(5) = vDS
        rsTP.Fields(6) = "17:00:00"
        rsTP.Fields(7) = vIP
        rsTP.Fields(8) = vNP
        rsTP.Fields(9) = "00:00:00"
        rsTP.Fields(19) = "00:00:00"
        
        rsTP.Fields(10) = "C.Ret" 'Tipo de apropriação
        rsTP.Fields(11) = "0000:00"
        rsTP.Fields(12) = "0000:00"
        rsTP.Fields(13) = "0000:00"
        rsTP.Fields(14) = "0000:00"
        
        rsParadas.MoveNext
    Loop
    rsTP.Update
    rsTP.Close
End Sub

Private Function achaHorarioSaida(vRegistro As String)
On Error GoTo Err
    Dim rsHorarioAlmoco As New ADODB.Recordset
    Dim SqlHorarioAlmoco As String
    'SqlHorarioAlmoco = "use CORPORERM " & _
    '    "DECLARE @Horario VARCHAR(4000) " & _
    '    "SET @Horario = '' " & _
    '    "SELECT @Horario = RTRIM(@Horario) + RTRIM((REPLICATE('0', 2 - LEN(CAST((a.BATIDA /60) AS VARCHAR))) + CAST((a.BATIDA /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((a.BATIDA %60) AS VARCHAR))) + CAST((a.BATIDA %60) AS VARCHAR))) + ';' FROM ABATHOR as a " & _
    '    "where a.INDICE = 1 AND A.BATIDA <> 0 GROUP BY A.CODHORARIO,A.INDICE, A.BATIDA " & _
    '    "select a.CHAPA,b.NOME,c.CODHORARIO,c.INDICE,SUBSTRING(@Horario,1,5) ENT1,SUBSTRING(@Horario,7,5) SAI1,SUBSTRING(@Horario,13,5) ENT2,SUBSTRING(@Horario,19,5) SAI2 from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
    '    "a.CODPESSOA = b.CODIGO inner join CORPORERM.dbo.ABATHOR as c on a.CODHORARIO = c.CODHORARIO where c.INDICE = 1 AND c.BATIDA <> 0 and a.CHAPA = '" & Format(vRegistro, "00000") & "' GROUP BY a.CHAPA,b.NOME,c.CODHORARIO,c.INDICE order by b.NOME"
    
    SqlHorarioAlmoco = "use " & vBancoTotvs & " " & _
                       "DECLARE @Horario VARCHAR(4000) " & _
                       "SET @Horario = '' " & _
                       "SELECT @Horario = RTRIM(@Horario) + RTRIM((REPLICATE('0', 2 - LEN(CAST((a.BATIDA /60) AS VARCHAR))) + CAST((a.BATIDA /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((a.BATIDA %60) AS VARCHAR))) + CAST((a.BATIDA %60) AS VARCHAR))) + ';' FROM ABATHOR as a where a.INDICE = 1 AND A.BATIDA <> 0 " & _
                       "GROUP BY A.CODHORARIO,A.INDICE, A.BATIDA select a.CHAPA, b.NOME, c.CODHORARIO,c.INDICE,SUBSTRING(@Horario,1,5) ENT1,SUBSTRING(@Horario,7,5) SAI1,SUBSTRING(@Horario,13,5) ENT2,SUBSTRING(@Horario,19,5) SAI2 from " & vBancoTotvs & ".dbo.PFUNC as a inner join " & vBancoTotvs & ".dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and a.CODPESSOA = b.CODIGO " & _
                       "inner join " & vBancoTotvs & ".dbo.ABATHOR as c on a.CODHORARIO = c.CODHORARIO where c.INDICE = 1 AND c.BATIDA <> 0 and a.CHAPA = '" & Format(vRegistro, "00000") & "' GROUP BY a.CHAPA,b.NOME,c.CODHORARIO,c.INDICE union select b.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AI,b.NOME COLLATE SQL_Latin1_General_CP1_CI_AI,'001' as CODHORARIO,'1' as INDICE, " & _
                       "SUBSTRING(@Horario,1,5) ENT1,SUBSTRING(@Horario,7,5) SAI1,SUBSTRING(@Horario,13,5) ENT2,SUBSTRING(@Horario,19,5) SAI2 from " & sDatabaseName & ".dbo.tbTerceirizados as b where b.CHAPA = '" & vRegistro & "' order by b.nome"
    rsHorarioAlmoco.Open SqlHorarioAlmoco, cnBanco, adOpenKeyset, adLockReadOnly
    If rsHorarioAlmoco.RecordCount > 0 Then
        achaHorarioSaida = rsHorarioAlmoco.Fields(7)
    Else
        achaHorarioSaida = "17:00"
    End If
    Exit Function
Err:
    achaHorarioSaida = "17:00"
End Function
