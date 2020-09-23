Attribute VB_Name = "IniFile"
'APIs to access INI files and retrieve data
'Tom Pydeski
'
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)

Function GetSettingIni(AppName As String, ByVal Section As String, ByVal Key As String, Optional DefValue As String) As Variant
'usage
'EncoderRes = GetSettingIni(App.Title, "Settings", "EncoderRes", 720)
'Returns info from an INI file
Dim Buffer As String
Dim iniFileName$
iniFileName$ = App.Path & "\" & AppName & ".ini"
If Dir(iniFileName$) = "" Then
    eMess$ = iniFileName$ & " not found!"
    GetSettingIni = DefValue
    MsgBox eMess$, vbCritical
    Exit Function
End If
'we should use an ini file instead of the registry
Buffer = String$(255, 0)
lReturn = GetPrivateProfileString(Section, Key, DefValue, Buffer, Len(Buffer), iniFileName$)
If lReturn = 0 Then
    GetSettingIni = ""
Else
    GetSettingIni = Left(Buffer, InStr(Buffer, Chr(0)) - 1)
End If
End Function

Function SaveSettingIni(AppName As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String) As Long
'Function returns 0 if successful and error number if unsuccessful
'usage
'SaveSettingIni App.Title, "Settings", "TallyLeft", frmTally.Left
Dim iniFileName$
SaveSettingIni = 1
iniFileName$ = App.Path & "\" & AppName & ".ini"
If Dir(iniFileName$) = "" Then
    eMess$ = iniFileName$ & " not found!"
    'MsgBox EMess$, vbCritical
    'Exit Function
End If
WritePrivateProfileString Section, Key, KeyValue, iniFileName$
SaveSettingIni = 0
End Function
