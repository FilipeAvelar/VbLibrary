Attribute VB_Name = "basKernel32"
Option Explicit

' Retrieve API error messages
  Declare Function _
    FormatMessage Lib "kernel32" Alias "FormatMessageA" _
       (ByVal dwFlags As Long, _
        lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        Arguments As Long) _
    As Long

' Retrieves the fully-qualified path for the file that contains the specified module.
  Declare Function _
    GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" _
       (ByVal hModule As Long, _
        ByVal lpFileName As String, _
        ByVal nSize As Long) _
    As Long
  
' Retrieves the fully-qualified path for the file that contains the specified module.
  Declare Function _
    GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
       (ByVal lpFileName As String) _
    As Long
  
  Declare Function _
    GetProcAddress Lib "kernel32" _
       (ByVal hModule As Long, _
        ByVal lpProcName As String) _
    As Long
  
Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeLong = ((HiWord * &H10000) + LoWord)
End Function

Public Function GetApiError(Optional lErrorId As Long) As String
    
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Const LANG_NEUTRAL = &H0
    
    If lErrorId = 0 Then lErrorId = Err.LastDllError
    GetApiError = String$(250, Chr$(0))
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrorId, LANG_NEUTRAL, GetApiError, 250, ByVal 0&

End Function

'Protótipo para funções que retornam a quantidade de caracteres obtida
'Private Function IDE_VB() As Boolean
'    Dim ModuleName As String
'    ModuleName = String$(128, Chr$(0))
'    ModuleName = Left$(ModuleName, GetModuleFileName(0&, ModuleName, Len(ModuleName)))
'End Function

