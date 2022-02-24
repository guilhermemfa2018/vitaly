VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmVariaveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FO - Variáveis"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstListView 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglstListImages"
      SmallIcons      =   "imglstListImages"
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0000
            Key             =   "S"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":015C
            Key             =   "Flag"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":02B8
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0414
            Key             =   "A"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0570
            Key             =   "Bolt"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":06CC
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0828
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0984
            Key             =   "Mail_New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0AE0
            Key             =   "Mail_Read"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVariaveis.frx":0C3C
            Key             =   "R"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVariaveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    listview_cabecalho
End Sub

Private Sub listview_cabecalho()
    lstListView.ColumnHeaders.Clear
    lstListView.ColumnHeaders.Add , , "VARIÁVEL", lstListView.Width / 1.5
    Compoe_Listview
End Sub

Private Sub Compoe_Listview()
    
    'Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    'ListView1.ListItems.Clear 'Limpa o listview
    'Set ItemLst = ListView1.ListItems.Add(, , Format(rsListview(0), "000000"))
    'ItemLst.SubItems(1) = "" & rsListview.Fields(1)
    'ItemLst.SubItems(2) = "" & rsListview.Fields(3)
    
    Dim lstEntry As ListItem
    Set lstEntry = lstListView.ListItems.Add(, , "FO_IPI_ALIQUOTA")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_IPI_VALOR")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_LUCRO_PERCENTUAL_COM_IPI")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_LUCRO_PERCENTUAL_SEM_IPI")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_LUCRO_POR_KG")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_LUCRO_VALOR")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_PESO_KG")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_QUANTIDADE")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_VALOR_BASE")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_VALOR_KG")
    Set lstEntry = lstListView.ListItems.Add(, , "FO_VALOR_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "IMPOSTOS_SOMA_PESO")
    Set lstEntry = lstListView.ListItems.Add(, , "IMPOSTOS_SOMA_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_AREA")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_AREA_M2_TON")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_AREA_UNIT")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_CRED_ICMS_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_CRED_IPI_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_CRED_IPI_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_CUSTO_MATERIAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PERCENTUAL_?")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PERCENTUAL_ICMS")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PERCENTUAL_IPI")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PERCENTUAL_PERDA")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PESO_SUBTOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PESO_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PESO_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PESO_UNIT")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PL_SUBTOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PL_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_PL_UNIT")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_QUANTIDADE_CJ")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_BRUTO")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_LIQUIDO")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_MEDIO")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_UNITARIO")
    Set lstEntry = lstListView.ListItems.Add(, , "MP_VALOR_UNITARIO")
    Set lstEntry = lstListView.ListItems.Add(, , "PINTURA_AREA")
    Set lstEntry = lstListView.ListItems.Add(, , "PINTURA_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "PINTURA_VALOR")
    Set lstEntry = lstListView.ListItems.Add(, , "RESUMOMP_FRETE_PERCENTUAL")
    Set lstEntry = lstListView.ListItems.Add(, , "RESUMOMP_FRETE_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "RESUMOMP_SOMA_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "RESUMOMP_VALOR_LIQUIDO")
    Set lstEntry = lstListView.ListItems.Add(, , "RESUMOMP_VALOR_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TESTES_ENSAIOS_VALOR_KG")
    Set lstEntry = lstListView.ListItems.Add(, , "TESTES_ENSAIOS_VALOR_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_BALDE")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_BALDE_MT2")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_BALDE_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_GALAO")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_GALAO_MT2SOLVENTE")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_GALAO_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_LATA")
    Set lstEntry = lstListView.ListItems.Add(, , "TINTAS_LATA_TOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANPOSTE_MP_TOTAL_CARRETAS")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_MP_TOTAL_SUBTOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_MP_TOTAL_VALORKG")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_PI_MADIA_PESOPOR_CARRETA")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_CARRETAS")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_SUBTOTAL")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_VALORKG")
    Set lstEntry = lstListView.ListItems.Add(, , "TRANSPORTEMP_MADIA_PESOPOR_CARRETA")

    
    lstListView.Refresh
End Sub
