Attribute VB_Name = "MIni"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Ret As String

Public Sub WriteINI(FileName As String, Section As String, Key As String, Text As String)
    WritePrivateProfileString Section, Key, Text, FileName
End Sub

Public Function ReadINI(FileName As String, Section As String, Key As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), FileName)
    Ret = Left$(Ret, RetLen)
    ReadINI = Ret
End Function

