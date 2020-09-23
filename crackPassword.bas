Attribute VB_Name = "crackPassword"
Public Type POINT
    X As Long
    Y As Long
End Type

Public Targeting As Boolean
Public CursorPosition As POINT

Global Const MainTitle = "crackPassword"
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Public Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As POINT) As Long
Public Declare Function WindowFromPoint Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetParent Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function IsWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sendmessagebystring Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Function GetTopLevelParent(ByVal hwndNum As Long) As Long
    Dim ParentHwnd As Long
    Dim tmpHwnd As Long
    
    tmpHwnd = hwndNum
 If 0 <> IsWindow(tmpHwnd) Then ' make sure the input hWnd refers to a window
            ParentHwnd = GetParent(tmpHwnd)
            tmpHwnd = ParentHwnd
   
 End If
    
    GetTopLevelParent = hwndNum 'suposed to be parent hwnd
End Function

