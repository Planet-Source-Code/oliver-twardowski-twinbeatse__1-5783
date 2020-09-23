Attribute VB_Name = "Module2"
Option Explicit
'----- Registry -----
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

    'Predifined constants for registry IO
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_DYN_DATA = &H80000006
    
    Public Const REG_SZ = 1                         ' Unicode nul terminated string
    Public Const REG_BINARY = 3                     ' Free form binary
    Public Const REG_DWORD = 4                      ' 32-bit number
    
    Public hCurKey As Long 'Needed it public to many needed it
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim lngValueType As Long
    Dim strBuffer As String
    Dim lngDataBufferSize As Long
    Dim intZeroPos As Integer

    ' Set up default value
    If Not IsEmpty(Default) Then
      GetSettingString = Default
    Else
      GetSettingString = ""
    End If

    ' Open the key and get length of string
    RegOpenKey hKey, strPath, hCurKey
    RegQueryValueEx hCurKey, strValue, 0&, lngValueType, ByVal 0&, lngDataBufferSize

    If lngValueType = REG_SZ Then
        ' initialise string buffer and retrieve string
        strBuffer = String(lngDataBufferSize, " ")
        RegQueryValueEx hCurKey, strValue, 0&, 0&, ByVal strBuffer, lngDataBufferSize
    
        intZeroPos = InStr(strBuffer, Chr$(0)) 'Finds null terminator
        
        If intZeroPos > 0 Then 'If null term is their then
            GetSettingString = Left$(strBuffer, intZeroPos - 1) 'Remove it
        Else
            GetSettingString = strBuffer
        End If
    End If

    RegCloseKey hCurKey 'Closes the current open key
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    RegCreateKey hKey, strPath, hCurKey 'Creates key , if exists opens key
    RegSetValueEx hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey hCurKey 'Closes the current open key
End Sub

