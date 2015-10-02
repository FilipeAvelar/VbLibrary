Attribute VB_Name = "basUser32"
Option Explicit

' WM_APP(-32768) to WM_USER
  Public sArrayMSG(0 To 1024) As String

' ______________
' * Constantes *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯

  Public Const WM_USER         As Long = (&H400)
  Public Const CW_USEDEFAULT   As Long = (&H80000000)

' GetSystemMetrics codes
  Enum eGetSystemMetrics
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    '#if(WINVER >= 0x0400)
    SM_CXFIXEDFRAME = SM_CXDLGFRAME           '/* ;win40 name change */
    SM_CYFIXEDFRAME = SM_CYDLGFRAME           '/* ;win40 name change */
    SM_CXSIZEFRAME = SM_CXFRAME               '/* ;win40 name change */
    SM_CYSIZEFRAME = SM_CYFRAME               '/* ;win40 name change */
    SM_SECURE = 44
    SM_CXEDGE = 45
    SM_CYEDGE = 46
    SM_CXMINSPACING = 47
    SM_CYMINSPACING = 48
    SM_CXSMICON = 49
    SM_CYSMICON = 50
    SM_CYSMCAPTION = 51
    SM_CXSMSIZE = 52
    SM_CYSMSIZE = 53
    SM_CXMENUSIZE = 54
    SM_CYMENUSIZE = 55
    SM_ARRANGE = 56
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
    SM_NETWORK = 63
    SM_CLEANBOOT = 67
    SM_CXDRAG = 68
    SM_CYDRAG = 69
    '#endif /* WINVER >= 0x0400 */
    SM_SHOWSOUNDS = 70
    '#if(WINVER >= 0x0400)
    SM_CXMENUCHECK = 71           '/* Use instead of GetMenuCheckMarkDimensions()! */
    SM_CYMENUCHECK = 72
    SM_SLOWMACHINE = 73
    SM_MIDEASTENABLED = 74
    '#endif /* WINVER >= 0x0400 */
    '#if (WINVER >= 0x0500) || (_WIN32_WINNT >= 0x0400)
    SM_MOUSEWHEELPRESENT = 75
    '#End If
    '#if(WINVER >= 0x0500)
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    SM_CMONITORS = 80
    SM_SAMEDISPLAYFORMAT = 81
    '#endif /* WINVER >= 0x0500 */
    '#if(_WIN32_WINNT >= 0x0500)
    SM_IMMENABLED = 82
    '#endif /* _WIN32_WINNT >= 0x0500 */
    '#if(_WIN32_WINNT >= 0x0501)
    SM_CXFOCUSBORDER = 83
    SM_CYFOCUSBORDER = 84
    '#endif /* _WIN32_WINNT >= 0x0501 */
    
    '#if(_WIN32_WINNT >= 0x0501)
    SM_TABLETPC = 86
    SM_MEDIACENTER = 87
    SM_STARTER = 88
    SM_SERVERR2 = 89
    '#endif /* _WIN32_WINNT >= 0x0501 */
    
    '#if(_WIN32_WINNT >= 0x0600)
    SM_MOUSEHORIZONTALWHEELPRESENT = 91
    SM_CXPADDEDBORDER = 92
    '#endif /* _WIN32_WINNT >= 0x0600 */
    
    '#if (WINVER < 0x0500) && (!defined(_WIN32_WINNT) || (_WIN32_WINNT < 0x0400))
'    SM_CMETRICS = 76
    '#elseif WINVER = 0x500
'    SM_CMETRICS = 83
    '#elseif WINVER = 0x501
'    SM_CMETRICS = 90
    '#Else
    SM_CMETRICS = 93
    '#End If
    
    '#if(WINVER >= 0x0500)
    SM_REMOTESESSION = &H1000
    '#if(_WIN32_WINNT >= 0x0501)
    SM_SHUTTINGDOWN = &H2000
    '#endif /* _WIN32_WINNT >= 0x0501 */
    '#if(WINVER >= 0x0501)
    SM_REMOTECONTROL = &H2001
    '#endif /* WINVER >= 0x0501 */
    '#if(WINVER >= 0x0501)
    SM_CARETBLINKINGENABLED = &H2002
    '#endif /* WINVER >= 0x0501 */
    '#endif /* WINVER >= 0x0500 */
End Enum




' _______________
' * Enumerações *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' ShowWindow Flags
  Enum eShowWindowFlags
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
    SW_MAX = 11
  End Enum

' Window Styles
  Enum eWINStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000               ' WS_BORDER | WS_DLGFRAME
    WS_CHILD = &H40000000
    WS_CHILDWINDOW = &H40000000         ' Common Window Styles
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_ICONIC = &H20000000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0
    WS_OVERLAPPEDWINDOW = &HCF0000      ' Common Window Styles
    WS_POPUP = &H80000000
    WS_POPUPWINDOW = &H80880000         ' Common Window Styles
    WS_SIZEBOX = &H40000                ' Obsolete style names
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_TILED = &H0                      ' Obsolete style names
    WS_TILEDWINDOW = &HCF0000           ' Obsolete style names
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
  End Enum

' Window sizing and positioning flags
  Enum eWINDOWPOSFlags
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20             ' The frame changed: send WM_NCCALCSIZE
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOOWNERZORDER = &H200           ' Don't do owner Z ordering
    SWP_NOSENDCHANGING = &H400
    SWP_DRAWFRAME = &H20
    SWP_NOREPOSITION = &H200
    SWP_DEFERERASE = &H2000
    SWP_ASYNCWINDOWPOS = &H4000
  End Enum

' SetWindowPos() hwndInsertAfter field values
  Enum eWINInsertAfter
    HWND_TOP = 0
    HWND_BOTTOM = 1
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
  End Enum

' ______________
' * Estruturas *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯

  Type POINT
    x               As Long
    y               As Long
  End Type

  Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
  End Type

' _________________________________________
' * Declarações - Funções e Procedimentos *
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' Sends the specified message to a window or windows.
  Declare Function _
    SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) _
    As Long

' Changes the size, position, and Z order of a child, pop-up, or top-level window.
  Declare Function _
    SetWindowPos Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As eWINInsertAfter, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As eWINDOWPOSFlags) _
    As Long

' Retrieves the dimensions of the bounding rectangle of the specified window.
  Declare Function _
    GetWindowRect Lib "user32" _
       (ByVal hwnd As Long, _
        lpRect As RECT) _
    As Long

' Retrieves the coordinates of a window's client area.
  Declare Function _
    GetClientRect Lib "user32" _
       (ByVal hwnd As Long, _
        lpRect As RECT) _
    As Long

' Creates an overlapped, pop-up, or child window with an extended window style
  Declare Function _
    CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
       (ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As eWINStyles, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        lpParam As Any) _
    As Long

' Destroys the specified window.
  Declare Function _
    DestroyWindow Lib "user32" _
       (ByVal hwnd As Long) _
    As Long

' Enables you to produce special effects when showing or hiding windows.
  Declare Function _
    AnimateWindow Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal dwTime As Long, _
        ByVal dwFlags As Long) _
    As Long

' Retrieves the name of the class to which the specified window belongs.
  Declare Function _
     GetClassName Lib "user32" Alias "GetClassNameA" _
       (ByVal hwnd As Long, _
        ByVal lpClassName As String, _
        ByVal nMaxCount As Long) _
    As Long

  Declare Function _
    UpdateWindow Lib "user32.dll" _
       (ByVal hwnd As Long) _
    As Long

  Declare Function _
    ShowWindow Lib "user32.dll" _
       (ByVal hwnd As Long, _
        ByVal nCmdShow As Long) _
    As Long



