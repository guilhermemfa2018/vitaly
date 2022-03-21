Attribute VB_Name = "Module1"
'-----------------------------------------------------
' Declarações para alteração da cor de fundo do componente SSTAB
Global gHookHWND As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type PAINTSTRUCT
    hdc                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type


Private Const GWL_WNDPROC = (-4)
Private Const STRETCHMODE = vbPaletteModeContainer

Private Const WM_PAINT = &HF

Private Declare Function BeginPaint Lib "user32" (ByVal HWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal HWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal HWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal HWnd As Long, ByVal lpString As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal HWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private pRect As RECT

'-----------------------------------------------------

Option Explicit
Public vLeftPadrao As Integer

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public CaminhoSkin As String

Public CorFundo As Long '--------------

Public RefreshVendas As Boolean
Public RefreshOS As Boolean

Public AgendaAberta As String, Servidor As String, CaMinho As String
Public vcaminhodosom As String  'caminho ao arquivo do som para o form de alerta/aviso
Public CodCompromisso As String  'Passa o código do compromisso ao clicar na mensagem
Public NumPopUp As Integer
Public NumAlturaPopUp As Integer
Public NumDistPopUp As Integer

'MsgBox
Public Onde As String
Public Onde1 As String
'Valor da resposta da msgbox
Public Tp As Integer
'Verificar se input tem valor de retorno ou não
Public Res As Boolean
'Valor da resposta da inputmsg
Public Inp As String

Public mobjMsg As Msgbox

'Public Tema As Msgbox
'Public SkinAtual As Msgbox

Public ResX As Single
Public ResY As Single
Public OldX As Single
Public OldY As Single
Public resolucao As Boolean
Public rsConf As New ADODB.Recordset
'muda data e símbolo de R$
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = 20
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

' muda resolução do vídeo
'Public Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type

Public Declare Function GetClipCursor Lib "user32.dll" (lprc As RECT) As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias _
"EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias _
"ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Public Enum TipoOS
    Fabricação
    Manutenção
    Usinagem
End Enum

Dim DevM As DEVMODE
Public Data30 As Date
Public vSubstituto As String, vMantemExpressao As String, vTituloFiltro As String
Public vIdFiltro As Integer

Public vColorThema(20) As String

Public objText(29, 29) As TextBox
Public objFrame(29, 29) As Frame
Public objCombo(29, 29) As ComboBox
Public objButton(29, 29) As VBControlExtender
Public objLabel(29, 29) As Label
Public objListview(29, 29) As MSComctlLib.Listview
Public objPicture(29, 29) As PictureBox
Public objImage As Image
Public vFramePrincipal As Frame
Public objButton1(29, 29) As VBControlExtender
Public objClose(29, 29) As VBControlExtender
Public vListViewPrincipal As Listview
Public vClosePrincipal As chameleonButton
Public vAcaoTab As String
Public vLabelPrincipal As Label
Public vPicBgPrincipal As PictureBox
Public tabAberta As Boolean
Public mo_Events As Collection

Public Sub ChangeRes(iWidth As Single, iHeight As Single)
   Dim A As Boolean
   Dim i As Long
   Do
      A = EnumDisplaySettings(0&, i, DevM)
      i = i + 1
   Loop Until (A = False)

   Dim b As Long
   DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
   DevM.dmPelsWidth = iWidth
   DevM.dmPelsHeight = iHeight
   b = ChangeDisplaySettings(DevM, 0)
End Sub

Public Function Img()
CaMinho = Servidor & ":"
Set Principal.Image1.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")
End Function

Public Function AplicarSkin(Frm As Form, Skin As Skin)
    CaminhoSkin = App.Path & "\MySkin.Skn"
    Skin.LoadSkin CaminhoSkin
    Skin.ApplySkin Frm.HWnd
    
    Set mobjMsg = New Msgbox
    'mobjMsg.Skin App.Path & "\MySkin.skn"
End Function

'funcao para ler valor de Chave
'Public Function iniReadKey(FileName As String, section As String, Key As String) As String
'    Dim RetVal As String * 255, v As Long
'    v = GetPrivateProfileString(section, Key, "", RetVal, 255, FileName)
'    iniReadKey = Left(RetVal, v)
'End Function

'Public Sub TemaMenu()
'    Tema = (iniReadKey(App.Path & "\config.ini", "TEMA", "NomeTema"))
'End Sub

Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.HWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.HWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Function

Function Lang_pt_br()
On Error GoTo Err
    Dim rsLang_pt_br As New ADODB.Recordset
    Dim SqlLang_pt_br As String

    SqlLang_pt_br = "SET LANGUAGE 'Brazilian'"
    rsLang_pt_br.Open SqlLang_pt_br, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
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

Public Function filtroPadrao()
    'Dim rsFiltroPadrao As New ADODB.Recordset
    'Dim sqlFiltroPadrao As String
    FiltroGeral = "Todos"
End Function


'-------------------------------------------------- teste

Public Function constroiTabs(vSSTab1 As SSTab, vCheck As Boolean)
    constroiTabs = False
    If verificaTabAberta(vSSTab1) = True Then
        constroiTabs = True
        Exit Function
    End If
    
'    If contaTabsAbertas(vSSTab1) + 1 > vOpenTabs Then
    If contaTabsAbertas(vSSTab1) = True Then
        mobjMsg.Abrir "Limite configurado para abertura de abas foi atingido. Feche uma aba para abrir uma nova aba", Ok, critico, "Atenção"
        constroiTabs = True
        Exit Function
    End If
    
    Dim vProximaTab As Integer, x As Integer, y As Integer
    x = 29
    vProximaTab = 0
    For y = 0 To x
        If vSSTab1.TabVisible(y) = False Then
            vProximaTab = y
            Exit For
        Else
            'vProximaTab = Y
            'vSSTab1.TabVisible(vProximaTab) = True
            'vSSTab1.Tab = vProximaTab
            'Exit Function
        End If
    Next
    If vProximaTab <= 29 Then
        vSSTab1.TabVisible(vProximaTab) = True
        vSSTab1.Tab = vProximaTab
        construirControles vProximaTab, vSSTab1, vCheck
'        construirBotoes vProximaTab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes vProximaTab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes vProximaTab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes vProximaTab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes vProximaTab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir.jpg", 360, 8040, 615, 615, "Admitir", True
'        construirBotoes vProximaTab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes vProximaTab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes vProximaTab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza.jpg", 360, 9840, 615, 615, "Atualizar", True
'        construirBotoes vProximaTab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado.jpg", 360, 10440, 615, 615, "Afastamento", True
'        construirBotoes vProximaTab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog.jpg", 360, 11040, 615, 615, "Programação", True
    End If
    tabAberta = True
    Set vSSTab1 = Nothing
End Function

Public Function verificaTabAberta(vSSTab1 As SSTab)
    verificaTabAberta = False
    Dim vProximaTab As Integer, x As Integer, y As Integer
    x = 29
    vProximaTab = 0
    For y = 0 To x
        If vSSTab1.TabCaption(y) = Formulario Then
            vSSTab1.TabVisible(y) = True
            vSSTab1.Tab = y
            verificaTabAberta = True
            Exit Function
        End If
    Next
End Function

Public Function contaTabsAbertas(vSSTab1 As SSTab) As Boolean
    contaTabsAbertas = False
    Dim vContaTabsAbertas As Integer, x As Integer, y As Integer
    x = 29
    For y = 0 To x
        If vSSTab1.TabVisible(y) = True Then
            vContaTabsAbertas = vContaTabsAbertas + 1
        End If
    Next
    ''MENSAGEM ABAIXO APENAS PARA REALIZAÇÃO DE TESTE
    'Debug.Print "Tabs abertas: " & vContaTabsAbertas + 1
    If vContaTabsAbertas + 1 > vOpenTabs Then
        Formulario = LegendaExc
        apontaLV = frmPesqGeralTeste2.picBg(vSSTab1.Tab).Tag
        
            'vSSTab1.TabVisible(frmPesqGeralTeste2.picBg(vSSTab1.Tab)) = True
            'vSSTab1.Tab = frmPesqGeralTeste2.picBg(vSSTab1.Tab)
        
        contaTabsAbertas = True
    End If
    If vContaTabsAbertas = 1 Then tabAberta = False
    Exit Function
End Function

Function desconstroiTabs(vSSTab1 As SSTab)
    Dim i As Long
    For i = 0 To 29
        vSSTab1.TabVisible(i) = False
    Next
End Function

Public Function setaComponentesTab(vSSTab1 As SSTab)
    Select Case vSSTab1.Caption
        Case Is = "Clientes"
            Formulario = "Clientes"
            Set vListViewPrincipal = vListviewClientes
            apontaLV = 1
            Set chamaForm = New frmClientes
            'contaColLVTeste vListviewClientes
        Case Is = "Tipo de Material"
            Set vListViewPrincipal = vListViewTipoMaterial
            Formulario = "Tipo de Material"
            apontaLV = 0
            Set chamaForm = New frmTipoMat
            'contaColLVTeste vListViewTipoMaterial
        Case Is = "Fórmula de Produtos"
            Formulario = "Fórmula de Produtos"
            Set vListViewPrincipal = vListviewFormulaPRD
            apontaLV = 4
            'contaColLVTeste vListviewFormulaPRD
        Case Is = "Paradas - OS"
            Formulario = "Paradas - OS"
            Set vListViewPrincipal = vListViewParadas
            apontaLV = 2
            Set chamaForm = New frmAtividades
            'contaColLVTeste vListViewParadas
        Case Is = "Transportadoras"
            Formulario = "Transportadoras"
            Set vListViewPrincipal = vListviewTransportadoras
            apontaLV = 3
            Set chamaForm = New frmTransportes
            'contaColLVTeste vListviewTransportadoras
        Case Is = "Desenhos"
            Formulario = "Desenhos"
            Set vListViewPrincipal = vListviewDesenhos
            apontaLV = 7
            Set chamaForm = New frmDesenhos
            'contaColLVTeste vListviewDesenhos
        Case Is = "Fórmula - Centro de Custo"
            Formulario = "Fórmula - Centro de Custo"
            Set vListViewPrincipal = vListviewFormulaCC
            apontaLV = 11
            Set chamaForm = New frmFormulaCC
            'contaColLVTeste vListviewFormulaCC
        Case Is = "FO"
            Formulario = "FO"
            Set vListViewPrincipal = vListviewComercial
            apontaLV = 5
            Set chamaForm = New frmFO
            'contaColLVTeste vListviewComercial
        Case Is = "Faturamento por FCE"
            Formulario = "Faturamento por FCE"
            Set vListViewPrincipal = vListviewFaturamentoFCE
            apontaLV = 20
            Set chamaForm = New FCRFatFCE
            'contaColLVTeste vListviewFaturamentoFCE
        Case Is = "FCE"
            Formulario = "FCE"
            Set vListViewPrincipal = vListviewFCE
            apontaLV = 6
            Set chamaForm = New frmFCECons
            'contaColLVTeste vListviewFCE
        Case Is = "LM"
            Formulario = "LM"
            Set vListViewPrincipal = vListviewLM
            apontaLV = 8
            Set chamaForm = New frmLM
            'contaColLVTeste vListviewLM
        Case Is = "MP"
            Formulario = "MP"
            Set vListViewPrincipal = vListviewMP
            apontaLV = 9
            Set chamaForm = New frmMPCompleto
            'contaColLVTeste vListviewMP
        Case Is = "Controle de Desenhos"
            Formulario = "Controle de Desenhos"
            Set vListViewPrincipal = vListviewControleDesenhos
            apontaLV = 10
            Set chamaForm = New frmCD
            'contaColLVTeste vListviewControleDesenhos
        Case Is = "RNCF"
            Formulario = "RNCF"
            Set vListViewPrincipal = vListviewRNCF
            apontaLV = 12
            Set chamaForm = New frmRNCF
            'contaColLVTeste vListviewRNCF
        Case Is = "Relatório de Inspeção"
            Formulario = "Relatório de Inspeção"
            Set vListViewPrincipal = vListviewRelInsp
            apontaLV = 16
            Set chamaForm = New frmRelInsp
            'contaColLVTeste vListviewRelInsp
        Case Is = "Imp. Rel. de Inspeção"
            Formulario = "Imp. Rel. de Inspeção"
            Set vListViewPrincipal = vListviewImpInspecao
            apontaLV = 19
            Set chamaForm = New FCRLibFab
            'contaColLVTeste vListviewImpInspecao
        Case Is = "Relatório de Expedição"
            Formulario = "Relatório de Expedição"
            Set vListViewPrincipal = vListviewRelExpedicao
            apontaLV = 17
            Set chamaForm = frmRelExp
            'contaColLVTeste vListviewRelExpedicao
        Case Is = "Imp. Rel. de Expedição"
            Formulario = "Imp. Rel. de Expedição"
            Set vListViewPrincipal = vListviewImpExpedicao
            apontaLV = 18
            Set chamaForm = New FCRExpedicao
            'contaColLVTeste vListviewImpExpedicao
        Case Is = "Grupos"
            Formulario = "Grupos"
            Set vListViewPrincipal = vListviewGrupos
            apontaLV = 14
            'Set chamaForm = New frmGrupos
            contaColLVTeste vListviewGrupos
        Case Is = "Usuários"
            Formulario = "Usuários"
            Set vListViewPrincipal = vListviewUsuarios
            apontaLV = 13
            Set chamaForm = New frmUsuarios
            'contaColLVTeste vListviewUsuarios
        Case Is = "OS Permissões"
            Formulario = "OS Permissões"
            Set vListViewPrincipal = vListviewPermissoes
            apontaLV = 15
            Set chamaForm = New frmPerColab
            'contaColLVTeste vListviewPermissoes
        Case Is = "Terceiros"
            Formulario = "Terceiros"
            Set vListViewPrincipal = vListviewTerceiros
            apontaLV = 21
            Set chamaForm = New frmTerceirizados
            'contaColLVTeste vListviewTerceiros
    End Select
    
End Function

Function construirControles(vTab As Integer, vSSTab1 As SSTab, vCheckBox As Boolean)
On Error GoTo Err
    Set objFrame(vTab, 0) = frmPesqGeralTeste2.Controls.Add("VB.Frame", "Frame1" + Trim(Str(vTab)), vSSTab1)
    With objFrame(vTab, 0)
        .Visible = True
        .Top = 360
        .Left = 120
        .Width = 20000
        .Height = 9015
        .BackColor = &HB7B7B7
        .Caption = "Informações"
    End With
    Set vFramePrincipal = objFrame(vTab, 0)
   
'    Set objClose(vTab, 0) = frmPesqGeralTeste2.Controls.Add("zeus.chameleonButton", "btnClose" + Trim(Str(vTab)), vSSTab1)
'    With objClose(vTab, 0)
'        .Visible = True
'        .Top = 60
'        .Left = vLeft '2160
'        .Width = 280
'        .Height = 280
'        .Caption = ""
'        .ButtonType = 11
'        .PictureNormal = frmPesqGeralTeste2.ImageList.ListImages(40).Picture
'        .Tag = "Fechar aba"
'        .ToolTipText = "Fechar aba: " & Trim(Str(vTab))
'        .BackColor = &HB7B7B7
'        .ZOrder (0)
'    End With
   
    Set objFrame(vTab, 1) = frmPesqGeralTeste2.Controls.Add("VB.Frame", "Frame0" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 1)
        .Visible = True
        .Top = 240
        .Left = 2760
        .Width = 5175
        .Height = 735
        .BackColor = &HB7B7B7
        .Caption = "Pesquisa"
    End With
    
    Set objCombo(vTab, 0) = frmPesqGeralTeste2.Controls.Add("VB.ComboBox", "Combo" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objCombo(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 120
        .Width = 2175
    End With
    
    Set objText(vTab, 0) = frmPesqGeralTeste2.Controls.Add("VB.TextBox", "Text" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objText(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 2400
        .Width = 2055
        .Height = 285
    End With

    Set objButton1(vTab, 0) = frmPesqGeralTeste2.Controls.Add("zeus.chameleonButton", "chameleonButton0" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objButton1(vTab, 0)
        .Visible = True
        .Top = 120
        .Left = 4560
        .Width = 495
        .Height = 495
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = Principal.ImageList.ListImages(30).Picture
        .Tag = "Pesquisar"
        .ToolTipText = "Pesquisar"
        .BackColor = &HB7B7B7
    End With

    Set objPicture(vTab, 0) = frmPesqGeralTeste2.Controls.Add("VB.PictureBox", "picBg" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objPicture(vTab, 0)
        .Visible = False
        .Top = 360
        .Left = 15600
        .Width = 855
        .Height = 495
    End With

    Set objFrame(vTab, 2) = frmPesqGeralTeste2.Controls.Add("VB.Frame", "Frame3" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 2)
        .Visible = True
        .Top = 240
        .Left = 12360
        .Width = 7000
        .Height = 735
        .Caption = "Filtro "
        .Appearance = 0
        .BackColor = &HB7B7B7
    End With

    Set objLabel(vTab, 0) = frmPesqGeralTeste2.Controls.Add("VB.Label", "Label1" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 120
        .Width = 735
        .Height = 255
        .Caption = "Status: "
        .BackColor = &HB7B7B7
    End With

    Set objLabel(vTab, 1) = frmPesqGeralTeste2.Controls.Add("VB.Label", "Label3" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 1)
        .Visible = True
        .Top = 240
        .Left = 1720
        .Width = 735
        .Height = 255
        .Caption = "Período: "
        .BackColor = &HB7B7B7
    End With

    Set objLabel(vTab, 2) = frmPesqGeralTeste2.Controls.Add("VB.Label", "Label2" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 2)
        .Visible = True
        .Top = 240
        .Left = 900
        .Width = 735
        .Height = 255
        .Caption = "-"
        .BackColor = &HB7B7B7
    End With

    Set objLabel(vTab, 3) = frmPesqGeralTeste2.Controls.Add("VB.Label", "Label4" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 3)
        .Visible = True
        .Top = 240
        .Left = 2520
        .Width = 735
        .Height = 255
        .Caption = "-"
        .BackColor = &HB7B7B7
    End With


    Set objLabel(vTab, 4) = frmPesqGeralTeste2.Controls.Add("VB.Label", "Label5" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objLabel(vTab, 4)
        .Visible = True
        .Alignment = 2
        .Font = "Calibri"
        .FontSize = 16
        .Top = 7100
        .Left = 200
        .Width = 20000
        .Height = 400
        .Caption = "Aguarde, Carregando dados..."
        .BackColor = &HB7B7B7
    End With
    Set vLabelPrincipal = objLabel(vTab, 4)

'    Set objListview(vTab, 0) = frmPesqGeralTeste2.Controls.Add("MSComctlLib.ListViewCtrl.2", "Listview2" + Trim(Str(vTab)), objFrame(vTab, 0))
'    With objListview(vTab, 0)
'        .Visible = True
'        .Top = 1080
'        .Left = 120
'        .Width = 16455
'        .Height = 7695
'        .Gridlines = True
'        .FullRowSelect = True
'        .LabelEdit = lvwManual
'        .LabelWrap = True
'        .SortKey = 0
'        .SortOrder = lvwAscending
'        .View = lvwReport
'        .BackColor = &H80000018
'        .ForeColor = &H800000
'        .SmallIcons = frmPesqGeralTeste2.ImgList
'        .Icons = frmPesqGeralTeste2.ImgList
'        .HideSelection = True
'        .Sorted = True
'        .AllowColumnReorder = True
'    End With
'    Set vListViewPrincipal = objListview(vTab, 0)
    
    
    Load frmPesqGeralTeste2.ListView2(vTab)
    With frmPesqGeralTeste2.ListView2(vTab)
        '.Visible = True
        .Top = 0
        .Left = 220
        .Width = 16455
        .Height = 7695
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .LabelWrap = True
        .SortKey = 0
        .SortOrder = lvwAscending
        .View = lvwReport
        .BackColor = &H80000018
        .ForeColor = &H800000
        
        If vColectionIcons = 1 Then
            .SmallIcons = frmPesqGeralTeste2.ImgList
            .Icons = frmPesqGeralTeste2.ImgList
        ElseIf vColectionIcons = 2 Then
            .SmallIcons = frmPesqGeralTeste2.ImgList1
            .Icons = frmPesqGeralTeste2.ImgList1
        ElseIf vColectionIcons = 3 Then
            .SmallIcons = frmPesqGeralTeste2.ImgList2
            .Icons = frmPesqGeralTeste2.ImgList2
        ElseIf vColectionIcons = 4 Then
            .SmallIcons = frmPesqGeralTeste2.ImgList3
            .Icons = frmPesqGeralTeste2.ImgList3
        ElseIf vColectionIcons = 5 Then
            .SmallIcons = frmPesqGeralTeste2.ImgList4
            .Icons = frmPesqGeralTeste2.ImgList
        End If
        
        
        '.SmallIcons = frmPesqGeralTeste2.ImgList
        '.Icons = frmPesqGeralTeste2.ImgList
        .HideSelection = True
        .Sorted = False
        .AllowColumnReorder = True
        .CheckBoxes = vCheckBox
    End With
    frmPesqGeralTeste2.ListView2(vTab).ZOrder (0)
    
'    Load frmPesqGeralTeste2.cmdClose(vTab)
'    With frmPesqGeralTeste2.cmdClose(vTab)
'        .Visible = True
'        .Top = 60
'        .Left = vLeft
'        .Width = 280
'        .Height = 280
'        .Caption = ""
'        .ButtonType = 11
'        .Tag = "Fechar aba"
'        .ToolTipText = "Fechar aba: " & Trim(Str(vTab))
'        .BackColor = &HB7B7B7
'        .ZOrder (0)
'    End With
'    frmPesqGeralTeste2.cmdClose(vTab).ZOrder (0)
    
    Load frmPesqGeralTeste2.picBg(vTab)
    With frmPesqGeralTeste2.picBg(vTab)
        .Visible = False
        .Tag = apontaLV
    End With
    frmPesqGeralTeste2.picBg(vTab).ZOrder (0)
    
 '   Set vClosePrincipal = frmPesqGeralTeste2.cmdClose(vTab)
    Set vPicBgPrincipal = frmPesqGeralTeste2.picBg(vTab)
    Set vListViewPrincipal = frmPesqGeralTeste2.ListView2(vTab)
Err:
    If Err.Number = 727 Then
        vTab = vTab + 1
        Resume
    End If
End Function

Public Function construirBotoes(vTab As Long, vBotao As Long, vCaminho As ImageList, vIndiceImage As Integer, vTop As Integer, vLeft As Integer, vWidth As Integer, vHeight As Integer, vTag As String, vVisible As Boolean)
On Error Resume Next
    Set objImage = frmPesqGeralTeste2.Controls.Add("VB.Image", "objImage" & vTab & vBotao, objFrame(vTab, 0))
    With objImage
        .Visible = vVisible
        .Top = vTop
        .Left = vLeft
        .Width = vWidth
        .Height = vHeight
        .Picture = vCaminho.ListImages(vIndiceImage).Picture
        .Tag = vTag
        .ToolTipText = vTag & vTab & vBotao
    End With
    mo_Events.Add New cEvents
    mo_Events(Val(vTab & vBotao)).Add_Image objImage, Val(vTab & vBotao)
End Function

Public Function desconstruirBotao(vTab As Long)
On Error Resume Next
    Dim x As Integer
    For x = 1 To 10
        frmPesqGeralTeste2.Controls.Remove ("objImage" & vTab & x)
    Next
End Function

Public Function desconstruirBotaoClose(vSSTab1 As SSTab)
On Error GoTo Err
    Dim i As Long
    For i = 0 To 29
        Unload frmPesqGeralTeste2.cmdClose(i)
    Next
    Exit Function
Err:
    'Debug.Print Err.Description & " - tab:" & i
    Resume Next
End Function

Public Function construirBotaoClose(vSSTab1 As SSTab)
On Error GoTo Err
    Dim x As Integer, y As Integer, vcontaTabAberta As Integer
    'desconstruirBotaoClose frmPesqGeralTeste2.SSTab1
    x = 29
    y = 0
    vcontaTabAberta = 0
    'Debug.Print String(50, 13)

  
    For y = y To x
        If vSSTab1.TabVisible(y) = True Then
            'vSSTab1.Tab = Y
            Load frmPesqGeralTeste2.cmdClose(y)
            Debug.Print "ENCONTROU TAB (Y): " & y
            With frmPesqGeralTeste2.cmdClose(y)
                .Visible = True
                .Top = 60
                .Left = leftDoBotaoFecharDaAba(vcontaTabAberta)
                .Width = 280
                .Height = 280
                .Caption = ""
                .ButtonType = 11
                .Tag = "Fechar aba"
                .ToolTipText = "Fechar aba: " & Trim(Str(y))
                .BackColor = &HB7B7B7
                '.ZOrder (0)
            End With
            'frmPesqGeralTeste2.cmdClose(Y).ZOrder (0)
        End If
        vcontaTabAberta = vcontaTabAberta + 1
        'Y = Y + 1
    Next
    Exit Function
Err:
    If Err.Number = 360 And vAcaoTab = "CLOSE" Then
        If frmPesqGeralTeste2.cmdClose(y).Left <> leftDoBotaoFecharDaAba(vcontaTabAberta) Then
            frmPesqGeralTeste2.SSTab1.Tab = y
            frmPesqGeralTeste2.cmdClose(y).Left = leftDoBotaoFecharDaAba(vcontaTabAberta)
        End If
'    If Err.Number = 360 And vAcaoTab = "OPEN" Then
    End If
    Debug.Print Err.Number & " - " & Err.Description
    'frmPesqGeralTeste2.cmdClose(Y).Left = leftDoBotaoFecharDaAba(vcontaTabAberta)
    'Unload frmPesqGeralTeste2.cmdClose(Y)
    vcontaTabAberta = vcontaTabAberta + 1
    y = y + 1
    'vSSTab1.Tab = Y
    Resume
End Function


Public Function leftDoBotaoFecharDaAba(vTabAtiva As Integer)
    Dim vLeftDoBotaoFechar As Integer
    'Debug.Print vTabAtiva
    
    If vTabAtiva = 0 Then
        vLeftDoBotaoFechar = 2160
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    
    
    If vTabAtiva = 1 Then
        vLeftDoBotaoFechar = 4680
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 2 Then
        vLeftDoBotaoFechar = 7200
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 3 Then
        vLeftDoBotaoFechar = 9720
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 4 Then
        vLeftDoBotaoFechar = 12240
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 5 Then
        vLeftDoBotaoFechar = 14760
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 6 Then
        vLeftDoBotaoFechar = 17280
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 7 Then
        vLeftDoBotaoFechar = 19800
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 8 Then
        vLeftDoBotaoFechar = 22320
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 9 Then
        vLeftDoBotaoFechar = 29880
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 10 Then
        vLeftDoBotaoFechar = 32400
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 11 Then
        vLeftDoBotaoFechar = 34920
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 12 Then
        vLeftDoBotaoFechar = 37440
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 13 Then
        vLeftDoBotaoFechar = 39960
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 14 Then
        vLeftDoBotaoFechar = 42480
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 15 Then
        vLeftDoBotaoFechar = 45000
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 16 Then
        vLeftDoBotaoFechar = 47520
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 17 Then
        vLeftDoBotaoFechar = 50040
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 18 Then
        vLeftDoBotaoFechar = 52560
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 19 Then
        vLeftDoBotaoFechar = 55080
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    If vTabAtiva = 20 Then
        vLeftDoBotaoFechar = 57600
'        vLeftDoBotaoFechar = vLeftPadrao
'        vLeftPadrao = vLeftPadrao + 2520
    End If
    leftDoBotaoFecharDaAba = vLeftDoBotaoFechar
    Debug.Print ".............LEFT:" & leftDoBotaoFecharDaAba
    Debug.Print String(1, 13)
End Function

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Abaixo: Funcões para alteração das cores de fundo do componente SSTAB
Public Function SubClassSSTAB(MySSTAB As SSTab, Pct As PictureBox)
    Pct.AutoRedraw = True
    Pct.AutoSize = True
    Pct.ScaleMode = vbPixels
    Pct.BackColor = Pct.BackColor
    
    'Save Grid fontname to use with DC's
    SetProp MySSTAB.HWnd, "lpPROC", SetWindowLong(MySSTAB.HWnd, GWL_WNDPROC, AddressOf MySubclassedGrid)
    SetProp MySSTAB.HWnd, "PctOBJ", ObjPtr(Pct)      'Save a pointer to PictureBox
    SetProp MySSTAB.HWnd, "GridOBJ", ObjPtr(MySSTAB)  'Save a pointer to Control
End Function

Public Sub UnSubClassSSTAB(ByVal hw As Long)
    Dim retval As Long
    retval = SetWindowLong(hw, GWL_WNDPROC, GetProp(hw, "lpPROC")) 'unsubclass Control
    'Clean up windows database
    RemoveProp hw, "lpPROC"
    RemoveProp hw, "PctOBJ"
    RemoveProp hw, "GridOBJ"
End Sub

Private Function MySubclassedGrid(ByVal hw As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim picTemp As PictureBox
Dim PicBACKGROUND As PictureBox
Dim GridTEMP As SSTab, GridREAL As SSTab

    gHookHWND = hw
    
    'Make GridTEMP a illegal reference - do not press END - Crash
    CopyMemory GridTEMP, GetProp(hw, "GridOBJ"), 4
    
    'Make it legal
    Set GridREAL = GridTEMP
    
    'Destroy illegal - no more crash
    CopyMemory GridTEMP, 0&, 4
    
    'Same story for PicTEMP
    CopyMemory picTemp, GetProp(hw, "PctOBJ"), 4
    Set PicBACKGROUND = picTemp
    CopyMemory picTemp, 0&, 4

    Select Case lMsg
         Case Is = WM_PAINT
            
            'We must do all the painting job
            Dim controlDC As Long, TempDC As Long, intDC As Long, tempBMP, intBMP As Long
            Dim aPS As PAINTSTRUCT
            Dim aDC As Long
            Dim Altura As Long
            Dim tppX, tppY As Long
            Dim backBuffDC, backBuffBmp As Long
            GetClientRect hw, pRect
            
                        
            'Start painting control ...
            Call BeginPaint(hw, aPS)
            aDC = aPS.hdc 'store painting DC
            
            'Prepare Double buffering ...No flickering
            backBuffDC = CreateCompatibleDC(aDC)
            backBuffBmp = CreateCompatibleBitmap(aDC, pRect.Right, pRect.Bottom)
            DeleteObject SelectObject(backBuffDC, backBuffBmp)
            
            'This is the big thing ! We are sendind WM_PAINT to our backbuffer
            MySubclassedGrid = CallWindowProc(GetProp(hw, "lpPROC"), hw, lMsg, ByVal backBuffDC, 0&)
                    
            With pRect
              'We just want to place a background picture, so let's Strech it
              Call SetStretchBltMode(backBuffDC, STRETCHMODE)
                    
              Call StretchBlt(backBuffDC, tppX, tppY, pRect.Right, pRect.Bottom, _
                    PicBACKGROUND.hdc, 0, 0, PicBACKGROUND.ScaleWidth, PicBACKGROUND.ScaleHeight, vbSrcAnd)
            End With
            
            'We have all the changes into backbuffer. Let's bring in back to control.hDc
            With aPS.rcPaint
               BitBlt aDC, .Left, .Top, .Right - .Left, .Bottom - .Top, backBuffDC, .Left, .Top, vbSrcCopy
            End With
            
            DeleteDC backBuffDC
            DeleteObject backBuffBmp
            Call EndPaint(hw, aPS)
            MySubclassedGrid = 0 'When a function intercepts WM_PAINT it must return 0
            
        Case Else
            'Call default windows procedure, stored in windows database in propertie lpPROC
            MySubclassedGrid = CallWindowProc(GetProp(hw, "lpPROC"), hw, lMsg, wParam, lParam)
    End Select
End Function


'---CONSTRUCAO DE BOTOES ESPECIFICOS POR TABS

Public Function contruirBotoesPorModulo(vQualLV As Integer)
    Dim vImageList As ImageList
    If vColectionIcons = 1 Then
        Set vImageList = Principal.ImageList
    ElseIf vColectionIcons = 2 Then
        Set vImageList = Principal.ImageList1
    ElseIf vColectionIcons = 3 Then
        Set vImageList = Principal.ImageList2
    ElseIf vColectionIcons = 4 Then
        Set vImageList = Principal.ImageList7
    ElseIf vColectionIcons = 5 Then
        Set vImageList = Principal.ImageList9
    End If

'-- TIPO DE MATERIAIS
    If vQualLV = 0 Then
        Formulario = "Tipo de Material"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
        
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If
'--CLIENTES
    If vQualLV = 1 Then
        Formulario = "Clientes"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
        
        
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
        'construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'--PARADAS
    If vQualLV = 2 Then
        Formulario = "Paradas - OS"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    
    
    End If

'--TRANSPORTADORAS
    If vQualLV = 3 Then
        Formulario = "Transportadoras"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    End If

'--FÓRMULAS PRODUTOS
    If vQualLV = 4 Then
        Formulario = "Fórmula de Produtos"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    End If

'--ORÇAMENTOS
    If vQualLV = 5 Then
        Formulario = "FO"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 7, 360, 8040, 615, 615, "Receber FO", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 22, 360, 9840, 615, 615, "Editar FCE", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 41, 360, 10440, 615, 615, "Impostos e Serviços", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 44, 360, 11040, 615, 615, "Receitas e Despesas", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\baixar.jpg", 360, 8040, 615, 615, "Receber FO", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\fce.jpg", 360, 9840, 615, 615, "Editar FCE", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imposto-02.jpg", 360, 10440, 615, 615, "Impostos e Serviços", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\ReceitaDespesa.jpg", 360, 11040, 615, 615, "Receitas e Despesas", True
    
    End If

'--FCE - Ficha de Controle de Encomenda
    If vQualLV = 6 Then
        Formulario = "FCE"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\pesquisar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    End If

'--DESENHOS
    If vQualLV = 7 Then
        Formulario = "Desenhos"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    End If

'--LM - LISTA DE MATERIAIS
    If vQualLV = 8 Then
        Formulario = "LM"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    
    End If

'--MP - Métodos e Processos
    If vQualLV = 9 Then
        Formulario = "MP"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 16, 360, 8040, 615, 615, "CD - Comunicação de Desvio", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 29, 360, 9840, 615, 615, "Atualizar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 14, 360, 10440, 615, 615, "Baixa Parcial", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 43, 360, 11040, 615, 615, "Programação", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\cd.jpg", 360, 8040, 615, 615, "CD - Comunicação de Desvio", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\retrabalho.jpg", 360, 9840, 615, 615, "Atualizar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\baixaParcial.jpg", 360, 10440, 615, 615, "Baixa Parcial", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog.jpg", 360, 11040, 615, 615, "Programação", True
    End If

'-- CONTROLE DE DESENHOS
    If vQualLV = 10 Then
        Formulario = "Controle de Desenhos"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'--FÓRMULA CENTRO DE CUSTO
    If vQualLV = 11 Then
        Formulario = "Fórmula - Centro de Custo"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- QUALIDADE - RNCF (Registro de Não Conformidade de Fabricação)
    If vQualLV = 12 Then
        Formulario = "RNCF"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 18, 360, 8040, 615, 615, "Causais", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\causais.jpg", 360, 8040, 615, 615, "Admitir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- USUÁRIOS
    If vQualLV = 13 Then
        Formulario = "Usuários"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- GRUPOS
    If vQualLV = 14 Then
        Formulario = "Grupos"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- OS FECHAMENTO - PERMISSÃO DE COLABORADORES
    If vQualLV = 15 Then
        Formulario = "OS Permissões"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
     
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
     End If

'-- LF - Relatório Liberação de Fabricação
    If vQualLV = 16 Then
        Formulario = "Relatório de Inspeção"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 42, 360, 720, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\inspecao.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 720, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\pintura.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- RO - Relatório de Expedição
    If vQualLV = 17 Then
        Formulario = "Relatório de Expedição"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 28, 360, 120, 615, 615, "Relatório de Expedição", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 27, 360, 720, 615, 615, "Relatório de Expedição Avulso", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 42, 360, 1320, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\transporte2.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\transporte1.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 1320, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- IMPRESSAO DOS RELATÓRIOS DE EXPEDIÇÃO
    If vQualLV = 18 Then
        Formulario = "Imp. Rel. de Expedição"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 42, 360, 720, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 720, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- IMPRESSAO DOS RELATÓRIOS DE INSPEÇÃO (QUALIDADE)
    If vQualLV = 19 Then
        Formulario = "Imp. Rel. de Inspeção"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 42, 360, 720, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 720, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

'-- FATURAMENTO POR FCE
    If vQualLV = 20 Then
        Formulario = "Faturamento por FCE"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 42, 360, 120, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 42, 360, 720, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 42, 360, 1320, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 29, 360, 11040, 615, 615, "Atualizar", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 120, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 720, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 1320, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza.jpg", 360, 11040, 615, 615, "Programação", True
    End If

'-- TERCEIRIZADOS
    If vQualLV = 21 Then
        Formulario = "Terceiros"
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, vImageList, 31, 360, 120, 615, 615, "Novo", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, vImageList, 32, 360, 720, 615, 615, "Editar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, vImageList, 33, 360, 1320, 615, 615, "Excluir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, vImageList, 34, 360, 1920, 615, 615, "Sair", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, vImageList, 42, 360, 8040, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, vImageList, 36, 360, 8640, 615, 615, "Filtrar", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, vImageList, 37, 360, 9240, 615, 615, "Imprimir", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, vImageList, 42, 360, 9840, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, vImageList, 42, 360, 10440, 615, 615, "", True
        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, vImageList, 42, 360, 11040, 615, 615, "", True
    
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo.jpg", 360, 120, 615, 615, "Novo", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar.jpg", 360, 720, 615, 615, "Editar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir.jpg", 360, 1320, 615, 615, "Excluir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair.jpg", 360, 1920, 615, 615, "Sair", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 8040, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro.jpg", 360, 8640, 615, 615, "Filtrar", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir.jpg", 360, 9240, 615, 615, "Imprimir", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 9840, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 10440, 615, 615, "", True
'        construirBotoes frmPesqGeralTeste2.SSTab1.Tab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\clear.jpg", 360, 11040, 615, 615, "", True
    End If

End Function


Public Function quantidadeDeFrom(vQualLV As Integer)
    If apontaLV = 0 Then vQdtFrom = 1
    If apontaLV = 1 Then vQdtFrom = 1
    If apontaLV = 2 Then vQdtFrom = 1
    If apontaLV = 3 Then vQdtFrom = 1
    If apontaLV = 4 Then vQdtFrom = 1
    If apontaLV = 5 Then vQdtFrom = 4
    If apontaLV = 6 Then vQdtFrom = 3
    If apontaLV = 7 Then vQdtFrom = 3
    If apontaLV = 8 Then vQdtFrom = 3
    If apontaLV = 9 Then vQdtFrom = 3
    If apontaLV = 10 Then vQdtFrom = 1
    If apontaLV = 11 Then vQdtFrom = 1
    If apontaLV = 12 Then vQdtFrom = 1
    If apontaLV = 13 Then vQdtFrom = 1
    If apontaLV = 14 Then vQdtFrom = 1
    If apontaLV = 15 Then vQdtFrom = 3
    If apontaLV = 16 Then vQdtFrom = 1
    If apontaLV = 17 Then vQdtFrom = 1
    If apontaLV = 18 Then vQdtFrom = 1
    If apontaLV = 19 Then vQdtFrom = 1
    If apontaLV = 20 Then vQdtFrom = 7
    If apontaLV = 21 Then vQdtFrom = 1
''-- TIPO DE MATERIAIS
'    If vQualLV = 0 Then
'    End If
'
''--CLIENTES
'    If vQualLV = 1 Then
'    End If
'
''--PARADAS
'    If vQualLV = 2 Then
'    End If
'
''--TRANSPORTADORAS
'    If vQualLV = 3 Then
'    End If
'
''--FÓRMULAS PRODUTOS
'    If vQualLV = 4 Then
'    End If
'
''--ORÇAMENTOS
'    If vQualLV = 5 Then
'    End If
'
''--FCE - Ficha de Controle de Encomenda
'    If vQualLV = 6 Then
'    End If
'
''--DESENHOS
'    If vQualLV = 7 Then
'    End If
'
''--LM - LISTA DE MATERIAIS
'    If vQualLV = 8 Then
'    End If
'
''--MP - Métodos e Processos
'    If vQualLV = 9 Then
'    End If
'
''-- CONTROLE DE DESENHOS
'    If vQualLV = 10 Then
'    End If
'
''--FÓRMULA CENTRO DE CUSTO
'    If vQualLV = 11 Then
'    End If
'
''-- QUALIDADE - RNCF (Registro de Não Conformidade de Fabricação)
'    If vQualLV = 12 Then
'    End If
'
''-- USUÁRIOS
'    If vQualLV = 13 Then
'    End If
'
''-- GRUPOS
'    If vQualLV = 14 Then
'    End If
'
''-- OS FECHAMENTO - PERMISSÃO DE COLABORADORES
'    If vQualLV = 15 Then
'     End If
'
''-- LF - Relatório Liberação de Fabricação
'    If vQualLV = 16 Then
'    End If
'
''-- RO - Relatório de Expedição
'    If vQualLV = 17 Then
'    End If
'
''-- IMPRESSAO DOS RELATÓRIOS DE EXPEDIÇÃO
'    If vQualLV = 18 Then
'    End If
'
''-- IMPRESSAO DOS RELATÓRIOS DE INSPEÇÃO (QUALIDADE)
'    If vQualLV = 19 Then
'    End If
'
''-- FATURAMENTO POR FCE
'    If vQualLV = 20 Then
'    End If
'
''-- TERCEIRIZADOS
'    If vQualLV = 21 Then
'    End If
End Function

Public Function Compoe_ListviewVariaveis(vLV As Listview)
    Dim lstEntry As ListItem
    Set lstEntry = vLV.ListItems.Add(, , "FO_IPI_ALIQUOTA")
    Set lstEntry = vLV.ListItems.Add(, , "FO_IPI_VALOR")
    Set lstEntry = vLV.ListItems.Add(, , "FO_LUCRO_PERCENTUAL_COM_IPI")
    Set lstEntry = vLV.ListItems.Add(, , "FO_LUCRO_PERCENTUAL_SEM_IPI")
    Set lstEntry = vLV.ListItems.Add(, , "FO_LUCRO_POR_KG")
    Set lstEntry = vLV.ListItems.Add(, , "FO_LUCRO_VALOR")
    Set lstEntry = vLV.ListItems.Add(, , "FO_PESO_KG")
    Set lstEntry = vLV.ListItems.Add(, , "FO_QUANTIDADE")
    Set lstEntry = vLV.ListItems.Add(, , "FO_VALOR_BASE")
    Set lstEntry = vLV.ListItems.Add(, , "FO_VALOR_KG")
    Set lstEntry = vLV.ListItems.Add(, , "FO_VALOR_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "IMPOSTOS_SOMA_PESO")
    Set lstEntry = vLV.ListItems.Add(, , "IMPOSTOS_SOMA_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_AREA")
    Set lstEntry = vLV.ListItems.Add(, , "MP_AREA_M2_TON")
    Set lstEntry = vLV.ListItems.Add(, , "MP_AREA_UNIT")
    Set lstEntry = vLV.ListItems.Add(, , "MP_CRED_ICMS_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_CRED_IPI_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_CRED_IPI_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_CUSTO_MATERIAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PERCENTUAL_?")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PERCENTUAL_ICMS")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PERCENTUAL_IPI")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PERCENTUAL_PERDA")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PESO_SUBTOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PESO_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PESO_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PESO_UNIT")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PL_SUBTOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PL_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_PL_UNIT")
    Set lstEntry = vLV.ListItems.Add(, , "MP_QUANTIDADE_CJ")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_BRUTO")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_LIQUIDO")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_MEDIO")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_UNITARIO")
    Set lstEntry = vLV.ListItems.Add(, , "MP_VALOR_UNITARIO")
    Set lstEntry = vLV.ListItems.Add(, , "PINTURA_AREA")
    Set lstEntry = vLV.ListItems.Add(, , "PINTURA_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "PINTURA_VALOR")
    Set lstEntry = vLV.ListItems.Add(, , "RESUMOMP_FRETE_PERCENTUAL")
    Set lstEntry = vLV.ListItems.Add(, , "RESUMOMP_FRETE_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "RESUMOMP_SOMA_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "RESUMOMP_VALOR_LIQUIDO")
    Set lstEntry = vLV.ListItems.Add(, , "RESUMOMP_VALOR_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TESTES_ENSAIOS_VALOR_KG")
    Set lstEntry = vLV.ListItems.Add(, , "TESTES_ENSAIOS_VALOR_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_BALDE")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_BALDE_MT2")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_BALDE_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_GALAO")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_GALAO_MT2SOLVENTE")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_GALAO_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_LATA")
    Set lstEntry = vLV.ListItems.Add(, , "TINTAS_LATA_TOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TRANPOSTE_MP_TOTAL_CARRETAS")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_MP_TOTAL_SUBTOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_MP_TOTAL_VALORKG")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_PI_MADIA_PESOPOR_CARRETA")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_CARRETAS")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_SUBTOTAL")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTE_PI_TOTAL_VALORKG")
    Set lstEntry = vLV.ListItems.Add(, , "TRANSPORTEMP_MADIA_PESOPOR_CARRETA")
    vLV.Refresh
End Function

Public Function Compoe_ListviewMatrizes(vLV As Listview)
    Dim lstEntry As ListItem
    Set lstEntry = vLV.ListItems.Add(, , "MT_DESPESASCREDITOS(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_PINTURA(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TESTES_ENSAIOS(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TRANSPORTES_MP(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TRANSPORTES_PI(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TINTA_LATA(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TINTA_GALAO(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_TINTA_BALDE(Linha,Coluna)")
    Set lstEntry = vLV.ListItems.Add(, , "MT_MP_(Linha,Coluna)")
    vLV.Refresh
End Function

Public Function carregaImagemBotao(vBtn As CommandButton, vIndex As Integer, vIcon As Integer)
    Dim vImageList As ImageList
    If vColectionIcons = 1 Then
        Set vImageList = Principal.ImageList
    ElseIf vColectionIcons = 2 Then
        Set vImageList = Principal.ImageList1
    ElseIf vColectionIcons = 3 Then
        Set vImageList = Principal.ImageList2
    ElseIf vColectionIcons = 4 Then
        Set vImageList = Principal.ImageList7
    ElseIf vColectionIcons = 5 Then
        Set vImageList = Principal.ImageList9
    End If
    vBtn.Picture = vImageList.ListImages(vIcon).Picture
    vBtn.BackColor = &HB7B7B7
End Function

Public Function montaTabMenu()
On Error GoTo Err
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsCopia As New ADODB.Recordset
    Dim sqlCopia As String
    
    
    Dim rsMenuExpert As New ADODB.Recordset
    Dim sqlMenuExpert As String
    
    sqlMenuExpert = "Select * from tbMenuConf order by idsub"
    rsMenuExpert.Open sqlMenuExpert, cnBanco, adOpenKeyset, adLockReadOnly
10  cnBanco.BeginTrans
    sqlDeletar = "Delete from tbMenu"
    rsDeletar.Open sqlDeletar, cnBanco
    
    If rsMenuExpert.RecordCount > 0 Then
        sqlCopia = "Select * into tbConfGrupoCOPIA from tbConfGrupo"
        rsCopia.Open sqlCopia, cnBanco

        sqlDeletar = "Delete from tbConfGrupo where tbconfgrupo.tipo <> 'CHK' and tbconfgrupo.idgrupo = '" & XCodGrp & "'"
        rsDeletar.Open sqlDeletar, cnBanco
        While Not rsMenuExpert.EOF
            SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values('" & rsMenuExpert.Fields(0) & "','" & rsMenuExpert.Fields(1) & "','" & rsMenuExpert.Fields(2) & "','" & rsMenuExpert.Fields(3) & "','" & rsMenuExpert.Fields(5) & "')"
            rsMenu.Open SqlMenu, cnBanco

            SqlMenu = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & XCodGrp & ",'" & rsMenuExpert.Fields(0) & "','" & rsMenuExpert.Fields(1) & "','" & rsMenuExpert.Fields(2) & "','" & rsMenuExpert.Fields(3) & "','S','" & rsMenuExpert.Fields(5) & "','" & rsMenuExpert.Fields(6) & "')"
            rsMenu.Open SqlMenu, cnBanco

            rsMenuExpert.MoveNext
        Wend
        rsMenuExpert.Close
        Set rsMenuExpert = Nothing

        'Restaurando Permissões
        sqlCopia = "Select * from tbConfGrupoCOPIA"
        rsCopia.Open sqlCopia, cnBanco, adOpenKeyset, adLockReadOnly
        While Not rsCopia.EOF
            SqlMenu = "Update tbConfGrupo set status = '" & rsCopia.Fields(5) & "',incluir = '" & rsCopia.Fields(9) & "',editar = '" & rsCopia.Fields(10) & "',excluir = '" & rsCopia.Fields(11) & "',salvar = '" & rsCopia.Fields(12) & "',imprimir = '" & rsCopia.Fields(13) & "',filtrar = '" & rsCopia.Fields(14) & "' where idgrupo = '" & rsCopia.Fields(0) & "' and idmenu = '" & rsCopia.Fields(1) & "' and idsub = '" & rsCopia.Fields(2) & "'"
            rsMenu.Open SqlMenu, cnBanco
            rsCopia.MoveNext
        Wend
        rsCopia.Close
        Set rsCopia = Nothing

        sqlCopia = "Drop table tbConfGrupoCOPIA"
        rsCopia.Open sqlCopia, cnBanco
    
    Else
        SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','TAB','Cadastros','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','CAT','Primários','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'02','CAT','Secundários','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0101','BUT','Ramo de atividades','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0102','BUT','Clientes','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0103','BUT','Transportadoras','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0104','BUT','Tipo material','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0205','BUT','Materiais','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0206','BUT','Itens verificação','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0207','BUT','Projetos','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0208','BUT','Processos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'02','TAB','Orçamentos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'11','CAT','Vendas','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'1111','BUT','Serviços','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'03','TAB','Planejamento','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'21','CAT','Planejamento e Controle da Produção','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2121','BUT','FCE','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2122','BUT','LM','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2123','BUT','LD','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2124','BUT','OS','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2125','BUT','Controle de Desenhos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'04','TAB','Produção','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'31','CAT','Acompanhamento de Produção','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3131','BUT','OS Acompanhamento','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3132','BUT','Evolução','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'05','TAB','Inspeção/Expedição','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'41','CAT','Emissão de Relatórios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4141','BUT','Emitir relatório','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4142','BUT','Imprimir relatório','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'06','TAB','Configurações','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'51','CAT','Parametrizações','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'52','CAT','Aparência','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5151','BUT','Sistema','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5152','BUT','Grupos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5153','BUT','Usuários','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5254','BUT','Menu','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5255','BUT','Skin','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5256','BUT','Fundo','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'07','TAB','Sobre','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'61','CAT','Sobre','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'6161','BUT','Sobre ZEUS','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'6162','BUT','Ajuda do ZEUS','" & vCodcoligada & "');"
        
        rsMenu.Open SqlMenu, cnBanco
    End If
    cnBanco.CommitTrans
    Set rsMenu = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Public Function abreConfMenu()
On Error GoTo Err

    Dim SqlConf As String
    SqlConf = "Select * from tbconfgrupo Where tbconfgrupo.idgrupo = '" & XCodGrp & "' and codcoligada = " & vCodcoligada & " order by idsub"
    rsConf.Open SqlConf, cnBanco, adOpenKeyset, adLockReadOnly
    Exit Function
Err:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Public Function fechaConfMenu()
    rsConf.Close
    Set rsConf = Nothing
End Function

Public Function montaMenu(vRB As XTREMERibbon)
    Dim vMenu As String
    While Not rsConf.EOF
        If rsConf.Fields(5) = "S" Then
            If rsConf.Fields(3) <> "CHK" Then
                If rsConf.Fields(3) = "TAB" Then
                    vRB.AddTab rsConf.Fields(1), rsConf.Fields(4)
                End If
                If rsConf.Fields(3) = "CAT" Then
                    vRB.AddCat Right$(rsConf.Fields(2), 2), rsConf.Fields(1), rsConf.Fields(4), False
                End If
                If rsConf.Fields(3) = "BUT" Then
                    If Len(rsConf.Fields(2)) = 4 Then
                        vRB.AddButton Right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 1, 2), rsConf.Fields(4), rsConf.Fields(8)
                    Else
                        vMenu = Val(Mid$(rsConf.Fields(2), 3, 3))
                        If Len(vMenu) <> 3 Then
                            vRB.AddButton Right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 4, 2), rsConf.Fields(4), rsConf.Fields(8)
                        Else
                            vRB.AddButton Right$(rsConf.Fields(2), 3), Mid$(rsConf.Fields(2), 3, 3), rsConf.Fields(4), rsConf.Fields(8)
                        End If
                    End If
                End If
            End If
        End If
        rsConf.MoveNext
    Wend
    vRB.Refresh
End Function

'FUNCAO PARA MUDAR TOOLTIPS
Public Sub MudaTool()
    On Error Resume Next
    Dim ctl As Control
    Dim i As Integer
    With chamaForm.cIpToolTips1
        .Create
        .Title = "Atenção:" 'Titulo do tooltip
        .MyIcon = itInfoIcon 'Icone do tooltip
        .BackColor = &H80000018  'Cor de fundo
        .ForeColor = &H800000    'Cor da letra e bordas
        For Each ctl In chamaForm.Controls
            If ctl.Tag <> "" Then
                .AddTool ctl, tfAbsolute, Replace(ctl.Tag, "|", vbCrLf)
            End If
        Next
    End With
End Sub

Public Sub inicializa_tabs(vSSTab As SSTab, vPic As PictureBox)
    vSSTab.Tab = 0
    SubClassSSTAB vSSTab, vPic
End Sub
