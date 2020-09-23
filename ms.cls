VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Sub get_owner()

    subkey = "Software\Microsoft\Windows\CurrentVersion"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegQueryValueEx(opened, "RegisteredOwner", 0, datatype, ByVal Svalue_read, 255)
    '*********** error1
    Form1.txt_owner = Trim(Svalue_read)
    
    retval = RegQueryValueEx(opened, "RegisteredOrganization", 0, datatype, ByVal Svalue_read, 255)
    '*********** error1
    
    Form1.txt_place = Trim(Svalue_read)
End Sub
Public Sub get_mouse()
    subkey = "Control Panel\Mouse"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "ActiveWindowTracking", 0, datatype, Lvalue_read, 4)
    '*********** error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_mouse(0).Value = vbChecked
        Lvalue_read = 0
    End If
    
    subkey = "Software\Microsoft\Office\8.0\Word\Options"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "LiveScrolling", 0, datatype, ByVal Svalue_read, 255)
    '*********** error1
    
    
    If (Left$(Svalue_read, 1) = "1") Then
        Form1.ms_mouse(1).Value = vbChecked
        Svalue_read = "0"
    End If
        
End Sub

Public Sub get_menu_graphics()
    
'**************** READING THE SPEED *****************
    
    subkey = "Control Panel\Desktop"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegQueryValueEx(opened, "MenuShowDelay", 0, datatype, ByVal Svalue_read, 255)
    '*********** error1
    
    Form1.menu_speed.Value = Val(Left$(Svalue_read, 3))

'******************* FINISHED READING THE SPEED *****************

'******************* READING SMART MENU *******************
    
    subkey = "Software\Microsoft\Office\9.0\Common\Toolbars"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    If retval <> 0 Then GoTo skip
    
    retval = RegQueryValueEx(opened, "AdaptiveMenus", 0, datatype, Lvalue_read, 4)
    '************* error1

    If (Lvalue_read = 0) Then
        Form1.ms_mnu(0).Value = vbChecked
    End If
'******************* FINISH READING SMART MENU ****************

skip:
'******************* READING HIDE DESKTOP *********************
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegQueryValueEx(opened, "NoDesktop", 0, datatype, Lvalue_read, 4)
    '************* error1

    If (Lvalue_read = 1) Then
        Form1.ms_mnu(1).Value = vbChecked
    End If

'****************** FINISH READING DESKTOP *****************

'******************* READING AUTO REFRESH ************

    subkey = "System\CurrentControlSet\Control\Update"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegQueryValueEx(opened, "UpdateMode", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 7) Then
        Form1.ms_mnu(2).Value = vbChecked
    End If

'********************** FINISH READING AUTO REFRESH ***************

'********************* READING THUMBNAIL *******************

    subkey = "Paint.Picture\DefaultIcon"
    
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "", 0, datatype, ByVal Svalue_read, 255)
    '************* error1

    If (Left$(Svalue_read, 2) = "%1") Then
        Form1.ms_mnu(3).Value = vbChecked
    End If

'********************** FINISHED READING THUMBNAIL **************

'********************* READING ARROW **********************
    
    subkey = "lnkfile"
    
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "IsShortcut", 0, datatype, ByVal Svalue_read, 255)
    If (retval <> 0) Then
        Form1.ms_mnu(4).Value = vbChecked
    End If
   
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "IsShortcut", 0, datatype, ByVal Svalue_read, 255)
    If (retval <> 0) Then
        Form1.ms_mnu(4).Value = vbChecked
    End If
    
'****************** FINISHED READING ARROW **************

'******************** READING FULL WINDOWS ***************
    
    subkey = "Control Panel\Desktop"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "DragFullWindows", 0, datatype, ByVal Svalue_read, 255)
    '************* error1

    If (Left$(Svalue_read, 1) = "1") Then
        Form1.ms_mnu(5).Value = vbChecked
    End If

'********************* FINISHED READING FULL WINDOWS **************

'********************* READING CLICK HERE *************************
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "NoStartBanner", 0, datatype, Lvalue_read, 4)
    '************* error1

    If (Lvalue_read = 1) Then
        Form1.ms_mnu(6).Value = vbChecked
    End If

'********************* FINISHED READING CLICK HERE ****************

'******************* READING NEW MENU ************************
    subkey = "Software\CLASSES\Directory\Background\shellex\ContextMenuHandlers\New"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "", 0, datatype, ByVal Svalue_read, 255)
    '************* error1
    
    If (Left(Svalue_read, 38) <> "{D969A300-E7FF-11d0-A93B-00A0C90F2719}") Then
        Form1.ms_mnu(7).Value = vbChecked
    End If
'******************* FINSHED READING NEW MENU *************

'*************** READING SEND TO MENU ****************
    subkey = "Software\CLASSES\AllFilesystemObjects\shellex\ContextMenuHandlers\Send To"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegQueryValueEx(opened, "", 0, datatype, ByVal Svalue_read, 255)
    '************* error1
    
    If (Left(Svalue_read, 38) <> "{7BA4C740-9E81-11CF-99D3-00AA004AE837}") Then
        Form1.ms_mnu(8).Value = vbChecked
    End If
'**************** FINISHED READING SEND TO MENU *****************
End Sub

Public Sub get_active_desktop()
    Dim create_open As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
    
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    
    retval = RegQueryValueEx(opened, "NoChangingWallpaper", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(0).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoComponents", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(1).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoAddingComponents", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(2).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoDeletingComponents", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(3).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoEditingComponents", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(4).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoCloseDragDropBands", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(5).Value = vbChecked
    End If
    
    retval = RegQueryValueEx(opened, "NoMovingBands", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(6).Value = vbChecked
    End If
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    
    retval = RegQueryValueEx(opened, "NoActiveDesktopChanges", 0, datatype, Lvalue_read, 4)
    '************* error1
    
    If (Lvalue_read = 1) Then
        Form1.ms_ad(7).Value = vbChecked
    End If
    
End Sub

Public Sub set_owner()
    Dim name_to_write As String
    Dim place_to_write As String
    
    subkey = "Software\Microsoft\Windows\CurrentVersion"
    
    name_to_write = Trim(Form1.txt_owner)
    place_to_write = Trim(Form1.txt_place)
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegSetValueEx(opened, "RegisteredOwner", 0, 1, ByVal name_to_write, Len(name_to_write))
    '*********** error1
    
    retval = RegSetValueEx(opened, "RegisteredOrganization", 0, 1, ByVal place_to_write, Len(place_to_write))
    '*********** error1

End Sub

Public Sub set_menu_graphics()
    
'******************** WRITING THE SPEED *******************
    Dim speed_to_write As String
    Dim create_open As Long
    speed_to_write = Trim(Str(Form1.menu_speed))
        
    subkey = "Control Panel\Desktop"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    retval = RegSetValueEx(opened, "MenuShowDelay", 0, 1, ByVal speed_to_write, Len(speed_to_write))
    '************* error1
'*********************** FINISH WRITING SPEED******************

'********************** WRITING SMART MENU ********************
    
    subkey = "Software\Microsoft\Office\9.0\Common\Toolbars"
    
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    
    If (Form1.ms_mnu(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "AdaptiveMenus", 0, 4, 0, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "AdaptiveMenus", 0, 4, 1, 4)
        '************* error1
    End If
'********************* FINISH WRITING SMART MENU *****************

'********************* WRITING HIDE DESKTOP ********************
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    If (Form1.ms_mnu(1).Value = 1) Then
        retval = RegSetValueEx(opened, "NoDesktop", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoDesktop", 0, 4, 0, 4)
        '************* error1
    End If

'********************* FINISH WRITING DESKTOP ****************
    
'******************** WRITING AUTO REFRESH *****************

    subkey = "System\CurrentControlSet\Control\Update"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error

    If (Form1.ms_mnu(2).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "UpdateMode", 0, 4, 7, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "UpdateMode", 0, 4, 1, 4)
        '************* error1
    End If

'******************** FINISHED AUTO REFRESH **************

'********************* WRITING THUMBNAIL *******************

    subkey = "Paint.Picture\DefaultIcon"
    
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    retval = RegDeleteValue(opened, "(Default)")
    If (Form1.ms_mnu(3).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "%1", 2)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "shimgvw.dll,1", 13)
        '************* error1
    End If
    
'********************** FINISHED WRITING THUMBNAIL **************
    
'********************* WRITING ARROW **********************
    
    subkey = "lnkfile"
    
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mnu(4).Value = vbChecked) Then
        retval = RegDeleteValue(opened, "IsShortcut")
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, opened)
        retval = RegDeleteValue(opened, "IsShortcut")
        '************* error1
    Else
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "lnkfile", 0, KEY_ALL_ACCESS, opened)
        retval = RegSetValueEx(opened, "IsShortcut", 0, 1, ByVal "", 1)
        retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, "piffile", 0, KEY_ALL_ACCESS, opened)
        retval = RegSetValueEx(opened, "IsShortcut", 0, 1, ByVal "", 1)
        '************* error1
    End If

'****************** FINISHED WRITING ARROW **************

'****************** WRITING FULL WINDOW **************
    subkey = "Control Panel\Desktop"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mnu(5).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "DragFullWindows", 0, 1, ByVal "1", 2)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "DragFullWindows", 0, 1, ByVal "0", 1)
        '************* error1
    End If

'*********************** FINISHED WRITING FULL WINDOW ****************

'*********************** WRITING CLICK HERE ************************
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mnu(6).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoStartBanner", 0, 3, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoStartBanner", 0, 3, 0, 4)
        '************* error1
    End If

'*********************** FINISHED CLICK HERE *********************

'*************** WRITING NEW MENU ********************
    subkey = "Software\CLASSES\Directory\Background\shellex\ContextMenuHandlers\New"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mnu(7).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "{Disabled}", 10)
        '************* error1
    Else
        
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "{D969A300-E7FF-11d0-A93B-00A0C90F2719}", 38)
        '************* error1
    End If
'****************** FINISHED WRITING NEW MENU ***************

'****************** WRITING SEND TO MENU ********************
    subkey = "Software\CLASSES\AllFilesystemObjects\shellex\ContextMenuHandlers\Send To"
    
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mnu(8).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "{Disabled}", 10)
        '************* error1
    Else
        
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "{7BA4C740-9E81-11CF-99D3-00AA004AE837}", 38)
        '************* error1
    End If
'****************** FINISHED WRITING SENT TO MENU *****************


End Sub

Public Sub set_active_desktop()
    Dim create_open As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
    
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    
    If (Form1.ms_ad(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoChangingWallpaper", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoChangingWallpaper", 0, 4, 0, 4)
        '************* error1
    End If
    
    If (Form1.ms_ad(1).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoComponents", 0, 4, 1, 4)
        '************* error1
    Else
            retval = RegSetValueEx(opened, "NoComponents", 0, 4, 0, 4)
        '************* error1
    End If
    
    If (Form1.ms_ad(2).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoAddingComponents", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoAddingComponents", 0, 4, 0, 4)
        '************* error1
    End If
    
    If (Form1.ms_ad(3).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoDeletingComponents", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoDeletingComponents", 0, 4, 0, 4)
        '************* error1
    End If
    
    If (Form1.ms_ad(4).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoEditingComponents", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoEditingComponents", 0, 4, 0, 4)
        '************* error1
    End If
    
    
    If (Form1.ms_ad(5).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoCloseDragDropBands", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoCloseDragDropBands", 0, 4, 0, 4)
        '************* error1
    End If
    
    
    If (Form1.ms_ad(6).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoMovingBands", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoMovingBands", 0, 4, 0, 4)
        '************* error1
    End If
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_ad(7).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "NoActiveDesktopChanges", 0, 4, 1, 4)
        '************* error1
    Else
        retval = RegSetValueEx(opened, "NoActiveDesktopChanges", 0, 4, 0, 4)
        '************* error1
    End If

End Sub

Public Sub get_drives()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        
    retval = RegQueryValueEx(opened, "NoDrives", 0, datatype, Lvalue_read, 4)
    '*********** error1
    If (Lvalue_read = 1) Then
        Form1.cu_drives(0).Value = True
        Lvalue_read = 0
    Else
        If (Lvalue_read = 67108863) Then
            Form1.cu_drives(1).Value = True
            Lvalue_read = 0
        Else
            Form1.cu_drives(2).Value = True
            Lvalue_read = 0
        End If
    End If

End Sub

Public Sub set_drives()
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        
    retval = RegQueryValueEx(opened, "NoDrives", 0, datatype, Lvalue_read, 4)
    '*********** error1
    
    If (Form1.cu_drives(0).Value = True) Then
       retval = RegSetValueEx(opened, "NoDrives", 0, 4, 1, 4)
       '*********** error1
    End If
    
    If (Form1.cu_drives(1).Value = True) Then
       retval = RegSetValueEx(opened, "NoDrives", 0, 4, 67108863, 4)
       '*********** error1
    End If
    
    If (Form1.cu_drives(2).Value = True) Then
       retval = RegSetValueEx(opened, "NoDrives", 0, 4, 0, 4)
       '*********** error1
    End If
   
End Sub

Public Sub set_mouse()
    
    Dim create_open As Long
    subkey = "Control Panel\Mouse"
    
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************* error
    
    If (Form1.ms_mouse(0).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "ActiveWindowTracking", 0, 4, 1, 4)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "ActiveWindowTracking", 0, 4, 0, 4)
        '*********** error1
    End If
    
    subkey = "Software\Microsoft\Office\8.0\Word\Options"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    '************* error
    
    If (Form1.ms_mouse(1).Value = vbChecked) Then
        retval = RegSetValueEx(opened, "LiveScrolling", 0, 1, ByVal "1", 1)
        '*********** error1
    Else
        retval = RegSetValueEx(opened, "LiveScrolling", 0, 1, ByVal "0", 1)
        '*********** error1
    End If
    
    
    
End Sub

