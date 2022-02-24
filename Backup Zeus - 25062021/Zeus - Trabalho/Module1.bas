Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public CaminhoSkin As String

Public CorFundo As Long '--------------

Public RefreshVendas As Boolean
Public RefreshOS As Boolean

Public AgendaAberta As String
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

'muda data e símbolo de R$
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = 20
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

' muda resolução do vídeo
Public Type RECT
   Left As Long
   Top As Long
   right As Long
   bottom As Long
End Type

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

Dim DevM As DEVMODE
Public Data30 As Date

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
    Dim rsLang_pt_br As New ADODB.Recordset
    Dim SqlLang_pt_br As String

    SqlLang_pt_br = "SET LANGUAGE 'Brazilian'"
    rsLang_pt_br.Open SqlLang_pt_br, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
End Function
