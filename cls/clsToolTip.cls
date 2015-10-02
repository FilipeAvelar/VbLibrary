VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' #Const WIN32_WINNT = &H600                ' Windows Vista Only

' ______________
' * Constantes *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯
  Const WM_USER As Long = (&H400)
' _______________
' * Enumerações *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' Tool Tip Messages
  Private Enum eTOOLTIPMessage
    TTM_ACTIVATE = (WM_USER + 1)
    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_ADDTOOL = (WM_USER + 4)
    TTM_DELTOOL = (WM_USER + 5)
    TTM_NEWTOOLRECT = (WM_USER + 6)
    TTM_RELAYEVENT = (WM_USER + 7)
    TTM_GETTOOLINFO = (WM_USER + 8)
    TTM_SETTOOLINFO = (WM_USER + 9)
    TTM_HITTEST = (WM_USER + 10)
    TTM_GETTEXT = (WM_USER + 11)
    TTM_UPDATETIPTEXT = (WM_USER + 12)
    TTM_GETTOOLCOUNT = (WM_USER + 13)
    TTM_ENUMTOOLS = (WM_USER + 14)
    TTM_GETCURRENTTOOL = (WM_USER + 15)
    TTM_WINDOWFROMPOINT = (WM_USER + 16)
    TTM_TRACKACTIVATE = (WM_USER + 17)      ' wParam = TRUE/FALSE start end  lparam = LPTOOLINFO
    TTM_TRACKPOSITION = (WM_USER + 18)      ' lParam = dwPos
    TTM_SETTIPBKCOLOR = (WM_USER + 19)
    TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    TTM_GETDELAYTIME = (WM_USER + 21)
    TTM_GETTIPBKCOLOR = (WM_USER + 22)
    TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
    TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    TTM_SETMARGIN = (WM_USER + 26)          ' lParam = lprc
    TTM_GETMARGIN = (WM_USER + 27)          ' lParam = lprc
    TTM_POP = (WM_USER + 28)
    TTM_UPDATE = (WM_USER + 29)
    TTM_GETBUBBLESIZE = (WM_USER + 30)
    TTM_ADJUSTRECT = (WM_USER + 31)
    TTM_SETTITLE = (WM_USER + 32)           ' wParam = TTI_*, lParam = char* szTitle
    TTM_POPUP = (WM_USER + 34)              ' (_WIN32_WINNT >= &H0501)
    TTM_GETTITLE = (WM_USER + 35)           ' wParam = 0, lParam = TTGETTITLE* (_WIN32_WINNT >= &H0501)
  End Enum

' In TOOLINFO, flags that control the ToolTip display. In NMTTDISPINFO, flags that indicates how to interpret the idFrom member of the included NMHDR structure")]
  Enum eTOOLINFOFlag
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_RTLREADING = &H4
    TTF_SUBCLASS = &H10
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_PARSELINKS = &H1000                 ' Comctl32.dll version 6 (manifest)
    TTF_DI_SETITEM = &H8000                 ' Valid only on the TTN_NEEDTEXT callback
  End Enum

' TOOLTIPSTYLES Styles uFlags
  Enum eTOOLTIPStyle
    TTS_ALWAYSTIP = &H1
    TTS_NOPREFIX = &H2
    TTS_NOANIMATE = &H10
    TTS_NOFADE = &H20
    TTS_BALLOON = &H40
    TTS_CLOSE = &H80
    TTS_USEVISUALSTYLE = &H100              ' Use themed hyperlinks (_WIN32_WINNT >= &H0600) */
  End Enum

' ToolTip Icons (Set with TTM_SETTITLE)
  Enum eTOOLTIPIcon
    TTI_NONE = 0
    TTI_INFO = 1
    TTI_WARNING = 2
    TTI_ERROR = 3
    #If WIN32_WINNT >= &H600 Then
        TTI_INFO_LARGE = 4
        TTI_WARNING_LARGE = 5
        TTI_ERROR_LARGE = 6
    #End If
  End Enum

' ______________
' * Estruturas *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯
  Private Type NMHDR
    hwndFrom        As Long
    idfrom          As Long
    code            As Long
  End Type

  Private Type NMTTDISPINFO
    hdr             As NMHDR
    lpszText        As String
    szText(0 To 79) As Byte
    hinst           As Long
    uFlags          As eTOOLINFOFlag
    lParam          As Long
    hbmp            As Long
  End Type

  Private Type TOOLINFO
    cbSize          As Long
    uFlags          As eTOOLINFOFlag
    hwnd            As Long
    uId             As Long
    RECT            As RECT
    hinst           As Long
    lpszText        As String
    lParam          As Long
  End Type

  Private Type TTGETTITLE
    cch             As Long
    dwSize          As Long
    pszTitle        As String
    uTitleBitmap    As Long
  End Type

' _________________________________________
' * Declarações - Funções e Procedimentos *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'' Sends the specified message to a window or windows.
'  Private Declare Function _
'    SendMessage Lib "user32" Alias "SendMessageA" _
'       (ByVal hwnd As Long, _
'        ByVal wMsg As eTOOLTIPMessage, _
'        ByVal wParam As Long, _
'        lParam As Any) _
'    As Long


' _______________________
' * Variáveis de módulo *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Internas
  Dim m_lHInstance        As Long
  Dim m_oToolInfo         As TOOLINFO
  Dim m_oRect             As RECT
  Dim m_lResult           As Long
  Public oParent          As Object

' Expostas via Get/Let

' Read_Only
  Dim m_lToolTipHwnd      As Long

' Read-Write
  Dim m_lLeft             As Long
  Dim m_lTop              As Long
  Dim m_sMensagem         As String
  Dim m_sTitulo           As String
  Dim m_lIcone            As eTOOLTIPIcon
  Dim m_lAparencia        As eTOOLTIPStyle
  Dim m_lEstilo           As eTOOLINFOFlag
'
'

'
' _____________________
' * Eventos da classe *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Class_Initialize()
'   Call InitVisualStyles
    m_lHInstance = App.hInstance
End Sub

Private Sub Class_Terminate()
    If m_lToolTipHwnd Then
       DestroyWindow m_lToolTipHwnd
       m_lToolTipHwnd = 0
    End If
End Sub

' ____________________
' * Funções privadas *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub VerifyApiError()
    If Err.LastDllError Then Debug.Print "Erro: " & Err.LastDllError & "-" & GetApiError & vbNewLine & "lResult: " & CBool(m_lResult)
End Sub

' ______________________________
' * Métodos públicos da classe *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Public Function Create(ByVal lHwndParent As Long, _
                       ByVal sMsg As String, _
                       Optional lAutomatico As Boolean = True, _
                       Optional ByVal sTitulo As String, _
                       Optional lIcone As eTOOLTIPIcon = TTI_NONE) As Long
'                      Optional lAparencia As eTOOLTIPStyle = (TTS_NOPREFIX Or TTS_BALLOON)) As Long

  ' Destroi qualquer janela anterior
    Call Class_Terminate

  ' Obtenho coordenadas do controle pai
    Call GetWindowRect(lHwndParent, m_oRect)
    VerifyApiError

  ' Create tooltip window and set it's properties
    m_lToolTipHwnd = CreateWindowEx(0&, _
                   "Tooltips_Class32", _
                   vbNullString, _
                   TTS_NOPREFIX Or TTS_BALLOON, _
                   CW_USEDEFAULT, _
                   CW_USEDEFAULT, _
                   CW_USEDEFAULT, _
                   CW_USEDEFAULT, _
                   lHwndParent, _
                   0&, _
                   m_lHInstance, _
                   0&)
    VerifyApiError
   ' TTS_NOPREFIX Or TTS_ALWAYSTIP Or TTS_BALLOON Or TTS_CLOSE,

    SetWindowPos m_lToolTipHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    VerifyApiError

  ' Initialize members of the toolinfo structure
  ' m_lLeft = m_oRect.Left + ((m_oRect.Right - m_oRect.Left) / 2)
  ' m_lTop = m_oRect.Top + ((m_oRect.Bottom - m_oRect.Top) / 2)
    
    m_lLeft = m_oRect.Left + ((m_oRect.Right - m_oRect.Left) / 2)
    m_lTop = m_oRect.Top + (m_oRect.Bottom - m_oRect.Top)
    
    m_sMensagem = sMsg
    m_lIcone = lIcone

    With m_oToolInfo
        .hinst = m_lHInstance
        .hwnd = lHwndParent
        .lpszText = m_sMensagem
         GetClientRect lHwndParent, .RECT
        .uFlags = TTF_SUBCLASS Or IIf(lAutomatico, 0, TTF_TRACK) 'Or TTF_CENTERTIP ' TTF_ABSOLUTE ' TTF_CENTERTIP  'Or TTF_ABSOLUTE 'Or TTF_CENTERTIP 'Or TTF_PARSELINKS
        .uId = 0&
        .cbSize = Len(m_oToolInfo)
    End With
    VerifyApiError

  ' Triga API do título
    Titulo = sTitulo

  ' Send an addtool message
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_ADDTOOL, 0&, m_oToolInfo)
    VerifyApiError
  
  ' Posiciona
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_TRACKPOSITION, 0&, ByVal MakeLong(m_lLeft, m_lTop))
    VerifyApiError

    Create = m_lToolTipHwnd

End Function

Public Function Ativar(bAtivar As Boolean) As Boolean
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_TRACKACTIVATE, bAtivar, m_oToolInfo)
End Function

' ___________________________________
' * Propriedades públicas da classe *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' Read-only
' ---------
Public Property Get hwnd() As Long: hwnd = m_lToolTipHwnd: End Property
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hwnd.VB_HelpID = 1000
Attribute hwnd.VB_UserMemId = -515

' Read-Write
' ----------

'=== Icone
Public Property Get Icone() As eTOOLTIPIcon: Icone = m_lIcone: End Property
Public Property Let Icone(ByVal lIcone As eTOOLTIPIcon)
    m_lIcone = lIcone
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_SETTITLE, m_lIcone, ByVal m_sTitulo)
End Property

'=== Titulo
Public Property Get Titulo() As String: Titulo = m_sTitulo: End Property
Public Property Let Titulo(ByVal sTitulo As String)
    m_sTitulo = sTitulo
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_SETTITLE, m_lIcone, ByVal m_sTitulo)
End Property

'=== Mensagem
Public Property Get Mensagem() As String: Mensagem = m_sMensagem: End Property
Public Property Let Mensagem(ByVal sMsg As String)
    m_sMensagem = sMsg
    With m_oToolInfo
        .lpszText = sMsg
        .cbSize = Len(m_oToolInfo)
    End With
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_UPDATETIPTEXT, 0&, m_oToolInfo)
End Property

'=== Left
Public Property Get Left() As Long: Left = m_lLeft: End Property
Public Property Let Left(ByVal lLeft As Long)
    m_lLeft = lLeft
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_TRACKPOSITION, 0&, ByVal MakeLong(m_lLeft, m_lTop))
End Property

'=== Top
Public Property Get Top() As Long: Top = m_lTop: End Property
Public Property Let Top(ByVal lTop As Long)
    m_lTop = lTop
    m_lResult = SendMessage(m_lToolTipHwnd, TTM_TRACKPOSITION, 0&, ByVal MakeLong(m_lLeft, m_lTop))
End Property

'=== Aparencia
Public Property Get Aparencia() As eTOOLTIPStyle: Aparencia = m_lAparencia: End Property
Public Property Let Aparencia(ByVal lAparencia As eTOOLTIPStyle)
    m_lAparencia = lAparencia
End Property

'=== Estilo
Public Property Get Estilo() As eTOOLINFOFlag: Estilo = m_lEstilo: End Property
Public Property Let Estilo(ByVal lEstilo As eTOOLINFOFlag)
    m_lEstilo = lEstilo
End Property


'Public Sub Testa()
'
'    Dim lResult As ScaleModeConstants
'    Dim oSubParent As Object
'    Dim lCont As Long
'
'    If Not (oParent Is Nothing) Then
'       Set oSubParent = oParent
'       On Error Resume Next
'       lResult = CallByName(oSubParent, "ScaleMode", VbGet)
'       If Err.Number > 0 Then
'          On Error Resume Next
'          lCont = lCont + 1
'          Set oSubParent = oSubParent.Parent
'          If Err.Number > 0 Then Exit Sub
'       End If
'    End If
'
'End Sub



'   SendMessage lToolTipHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal -1&
'   SendMessage lToolTipHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal -1&

