Attribute VB_Name = "ie"

Public Sub get_ie()
    Dim create_open As Long
    
    subkey = "Software\Microsoft\Internet Explorer\Toolbar"
    '****************** GETTING THE IE TOOLBAR PICTURE
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "BackBitmapIE5", 0, datatype, ByVal Svalue_read, 255)
        
        If (retval <> 0) Then
            Form1.ie_toolpic = "<Default>"
        Else
            Form1.ie_enable_explorer.Value = vbChecked
            Form1.ie_toolpic = Trim(Svalue_read)
        End If
    
    '************************* GETTING THE SYSTEM TOOLBAR
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "BackBitmapShell", 0, datatype, ByVal Svalue_read, 255)
        
        If (retval <> 0) Then
            Form1.ie_systempic = "<Default>"
        Else
            Form1.ie_enable_windows.Value = vbChecked
            Form1.ie_systempic = Trim(Svalue_read)
        End If
    
        
    
    subkey = "Software\Microsoft\Internet Explorer\Main"
    
    '**************** GETTING THE IE CAPION
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "Window Title", 0, datatype, ByVal Svalue_read, 255)
        If (retval <> 0) Then
            Form1.ie_caption = "Microsoft Internet Explorer"
        Else
            Form1.ie_caption = Trim(Svalue_read)
        End If
        
    '************ GETTING THE SCRIPT DEBUGGER ***************
        retval = RegQueryValueEx(opened, "Disable Script Debugger", 0, datatype, ByVal Svalue_read, 255)
        If (retval = 0 And Left$(Svalue_read, 3) = "Yes") Then
            Form1.ie_options(2).Value = vbChecked
        End If
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

    '************* GETTING HIDE ICON IN DESKTOP
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "NoInternetIcon", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(0).Value = vbChecked
        End If
        
    '*************** GETTING THE EXPANEDED NEW MENU ******************
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "NoExpandedNewMenu", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(3).Value = vbChecked
        End If
        
    
    subkey = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    '************** GETTING THE BROWSER CLOSE ****************
        retval = RegQueryValueEx(opened, "NoBrowserClose", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(4).Value = vbChecked
        End If
    
    '************** GETTING THE RIGHT CLICK MENU ****************
        retval = RegQueryValueEx(opened, "NoBrowserContextMenu", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(5).Value = vbChecked
        End If
        
    '************** GETTING THE INTERNET OPTIONS ****************
        retval = RegQueryValueEx(opened, "NoBrowserOptions", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(6).Value = vbChecked
        End If
    
    '************** GETTING THE SAVE AS ****************
        retval = RegQueryValueEx(opened, "NoBrowserSaveAs", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(7).Value = vbChecked
        End If
        
    '************** GETTING FAVORITES MENU ****************
        retval = RegQueryValueEx(opened, "NoFavorites", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(8).Value = vbChecked
        End If
        
    '************** GETTING NEW MENU ****************
        retval = RegQueryValueEx(opened, "NoFileNew", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(9).Value = vbChecked
        End If
    
    '************** GETTING OPEN COMMAND ****************
        retval = RegQueryValueEx(opened, "NoFileOpen", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(10).Value = vbChecked
        End If
        
    '************** GETTING FULL SCREEN MODE ***************
        retval = RegQueryValueEx(opened, "NoTheaterMode", 0, datatype, Lvalue_read, 4)
        If (retval = 0 And Lvalue_read = 1) Then
            Form1.ie_options(14).Value = vbChecked
        End If
        
    subkey = "Software\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete"
    '********** GETTING AUTO COMPLETE
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegQueryValueEx(opened, "Append Completion", 0, datatype, ByVal Svalue_read, 255)
        
        If (retval = 0 And Left$(Svalue_read, 2) = "no") Then
            Form1.ie_options(1).Value = vbChecked
        End If

End Sub

Public Sub set_ie()
    
    Dim create_open As Long
    Dim temp_string As String
    
    subkey = "Software\Microsoft\Internet Explorer\Toolbar"
    
    '****************** SETTING THE IE TOOLBAR PICTURE
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
                
        If (Form1.ie_enable_explorer.Value = vbChecked) Then
            temp_string = Trim(Form1.ie_toolpic)
            retval = RegSetValueEx(opened, "BackBitmapIE5", 0, 1, ByVal temp_string, 255)
        Else
            retval = RegDeleteValue(opened, "BackBitmapIE5")
        End If
    
    '****************** SETTING THE SYSTEM TOOLBAR PICTURE
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
                
        If (Form1.ie_enable_windows.Value = vbChecked) Then
            temp_string = Trim(Form1.ie_systempic)
            retval = RegSetValueEx(opened, "BackBitmapShell", 0, 1, ByVal temp_string, 255)
        Else
            retval = RegDeleteValue(opened, "BackBitmapShell")
        End If
        
    
    
    subkey = "Software\Microsoft\Internet Explorer\Main"
    '***************** SETTING THE IE CAPTION
        temp_string = Trim(Form1.ie_caption)
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        retval = RegSetValueEx(opened, "Window Title", 0, 1, ByVal temp_string, Len(temp_string))

    '************** SETTING THE SCRIPT DEBUGGER ***********
        If (Form1.ie_options(2).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "Disable Script Debugger", 0, 1, ByVal "yes", 3)
        Else
            retval = RegSetValueEx(opened, "Disable Script Debugger", 0, 1, ByVal "no", 2)
        End If
    
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    '************* SETTING HIDE ICON IN DESKTOP
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
        If (Form1.ie_options(0).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoInternetIcon", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoInternetIcon", 0, 4, 0, 4)
        End If
        
    '**************** SETTING THE EXPANSION OF NEW MENU
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, 0)
        If (Form1.ie_options(3).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoExpandedNewMenu", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoExpandedNewMenu", 0, 4, 0, 4)
        End If
    
    subkey = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    
    '************** SETTING THE BROWSER CLOSE ****************
        If (Form1.ie_options(4).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoBrowserClose", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoBrowserClose", 0, 4, 0, 4)
        End If
    
    '************** SETTING THE RIGHT CLICK MENU ****************
        If (Form1.ie_options(5).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoBrowserContextMenu", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoBrowserContextMenu", 0, 4, 0, 4)
        End If
    
    '************** SETTING THE INTERNET OPTIOS ****************
        If (Form1.ie_options(6).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoBrowserOptions", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoBrowserOptions", 0, 4, 0, 4)
        End If
        
    '************** SETTING THE SAVE AS ****************
        If (Form1.ie_options(7).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoBrowserSaveAs", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoBrowserSaveAs", 0, 4, 0, 4)
        End If
    
    '************** SETTING FAVORITES MENU ****************
        If (Form1.ie_options(8).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoFavorites", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoFavorites", 0, 4, 0, 4)
        End If
        
    '************** SETTING NEW MENU ****************
        If (Form1.ie_options(9).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoFileNew", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoFileNew", 0, 4, 0, 4)
        End If
    
    '************** SETTING OPEN COMMAND ****************
        If (Form1.ie_options(10).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoFileOpen", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoFileOpen", 0, 4, 0, 4)
        End If
        
    '************** SETTING FULL SCREEN MODE ****************
        If (Form1.ie_options(14).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "NoTheaterMode", 0, 4, 1, 4)
        Else
            retval = RegSetValueEx(opened, "NoTheaterMode", 0, 4, 0, 4)
        End If
        
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete"
    '************ SETTING AUTO COMPLETE *****************
        retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
        If (Form1.ie_options(1).Value = vbChecked) Then
            retval = RegSetValueEx(opened, "Append Completion", 0, 1, ByVal "no", 3)
        Else
            retval = RegSetValueEx(opened, "Append Completion", 0, 1, ByVal "yes", 2)
        End If
            
            
End Sub
