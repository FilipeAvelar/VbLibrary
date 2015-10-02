Attribute VB_Name = "basComCtl32"
Option Explicit

Const ICC_USEREX_CLASSES As Long = &H200

  Type tagInitCommonControlsEx
    lngSize     As Long
    lngICC      As Long
  End Type

' Receives DLL-specific version information. It is used with the DllGetVersion function.
  Type DLLVERSIONINFO
    cbSize          As Long
    dwMajorVersion  As Long
    dwMinorVersion  As Long
    dwBuildNumber   As Long
    dwPlatformID    As Long
  End Type

  Type ULONGLONG
    wQFE            As Integer
    wBuild          As Integer
    wMinorVersion   As Integer
    wMajorVersion   As Integer
  End Type
  
  Type DLLVERSIONINFO2
    info1           As DLLVERSIONINFO
    dwFlags         As Long
    ullVersion      As ULONGLONG
  End Type

  Declare Sub _
    InitCommonControls Lib "comctl32" ()
  
  Declare Function _
    InitCommonControlsEx Lib "comctl32" _
        (iccex As tagInitCommonControlsEx) _
    As Boolean

  Declare Function _
    DllGetVersion Lib "comctl32" _
        (dwVersion As DLLVERSIONINFO2) As Long

Public Function InitVisualStyles() As Boolean

   On Error Resume Next  ' this will fail if Comctl not available
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex

End Function

    
Public Function ComCtl32LoadedVersion() As String

    Dim oVersion    As DLLVERSIONINFO2
    Dim lHModule    As Long

    lHModule = GetModuleHandle("COMCTL32")
    
    If lHModule <> 0 Then      
     
        If GetProcAddress(lHModule, "DllGetVersion") <> 0 Then
           oVersion.info1.cbSize = Len(oVersion)
           Call DllGetVersion(oVersion)

           With oVersion.ullVersion
                ComCtl32LoadedVersion = .wMajorVersion & "." & _
                                        .wMinorVersion & "." & _
                                        .wBuild & "." & _
                                        .wQFE
           End With
        End If
    End If

End Function

