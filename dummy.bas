Attribute VB_Name = "Module1"
Option Explicit

Public Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Boolean
End Type

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function ExitWindowsEx Lib "User32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long




Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_DWORD = 4
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CLASSES_ROOT = &H80000000


Public state_read As String * 3
Public retval As Long
Public datatype As Long
Public Lvalue_read As Long
Public Svalue_read As String * 255
Public opened As Long
Public secattr As SECURITY_ATTRIBUTES
Public subkey As String
Public flag As Integer

Public Sub main()
    Form1.cmd_apply.Enabled = False
    retval = GetProfileString("intl", "sType", "N", state_read, 3)
    If Left$(state_read, 1) = "0" Then
        Form1.Check1.Value = 0
        Form1.Show
    Else
        Form1.Check1.Value = vbChecked
        frmLogin.Show
    End If
    frmSplash.Show 1
End Sub


Public Function encrypt(pass As String) As String
    Dim i, k As Integer
    Dim encrypted As String
    encrypted = ""
    For i = 1 To Len(pass)
        k = Asc(Mid$(pass, i, 1))
        k = k * 8 - 34
        encrypted = encrypted + Trim(Str(k))
    Next
    encrypt = encrypted
End Function

Public Function decrypt(pass As String) As String
    Dim i As Integer
    Dim decrypted As String
    Dim k As Integer
    decrypted = ""
    For i = 1 To Len(pass) Step 3
        k = Val(Mid$(pass, i, 3))
        k = (k + 34) / 8
        decrypted = decrypted + Chr$(k)
    Next
    decrypt = decrypted
End Function
