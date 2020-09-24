Attribute VB_Name = "Registry"
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Const REG_DWORD = 4

Enum REG
    HKEY_CURRENT_USER = &H80000001
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Enum TypeStringValue
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
End Enum

Public Function DeleteValue(hKey As REG, Subkey As String, lpValName As String) As Long
Dim Ret As Long

    On Error Resume Next
    RegOpenKey hKey, Subkey, Ret
    DeleteValue = RegDeleteValue(Ret, lpValName)
    RegCloseKey Ret
    
End Function

Public Function CreateStringValue(hKey As REG, Subkey As String, RTypeStringValue As TypeStringValue, strValueName As String, strData As String) As Long
    
    On Error Resume Next
    Dim Ret As Long
    
    RegCreateKey hKey, Subkey, Ret
    CreateStringValue = RegSetValueEx(Ret, strValueName, 0, RTypeStringValue, ByVal strData, Len(strData))
    RegCloseKey Ret
    
End Function

