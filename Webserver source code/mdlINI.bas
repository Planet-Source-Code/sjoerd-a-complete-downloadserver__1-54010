Attribute VB_Name = "mdlINI"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    '*** DESCRIPTION:Reads from an *.INI fil
    '     e strFileName (full path &  file name)
    '*** RETURNS:The string stored in [strSe
    '     ctionHeader], line beginning strVariableName=
    '*** NOTE: Requires declaration of API c
    '     all     GetPrivateProfileString
    'Initialise variable
    Dim strReturn As String
    'Blank the return string
    strReturn = String(255, Chr(0))
    'Get requested information, trimming the
    '     returned string
    mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function


