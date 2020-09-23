Attribute VB_Name = "ms1"

Public Sub clearrun()
    Dim valuelen As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    While retval = 0
        valuelen = 255
        retval = RegEnumValue(opened, 0, Svalue_read, valuelen, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then
            retval = RegDeleteValue(opened, Left(Svalue_read, valuelen))
        End If
    Wend
    Form1.ms1_clear_run.Enabled = False
End Sub


Public Sub get_run_history()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    If (retval = 0) Then
        retval = RegEnumValue(opened, 0, Svalue_read, 255, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then Form1.ms1_clear_run.Enabled = True
    End If
End Sub

Public Sub clearie()
    Dim valuelen As Long
    subkey = "Software\Microsoft\Internet Explorer\TypedURLs"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    While retval = 0
        valuelen = 255
        retval = RegEnumValue(opened, 0, Svalue_read, valuelen, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then
            retval = RegDeleteValue(opened, Left(Svalue_read, valuelen))
        End If
    Wend
    Form1.ms1_clear_ie.Enabled = False
End Sub

Public Sub get_ie_history()
    subkey = "Software\Microsoft\Internet Explorer\TypedURLs"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    If (retval = 0) Then
        retval = RegEnumValue(opened, 0, Svalue_read, 255, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then Form1.ms1_clear_ie.Enabled = True
    End If
End Sub


Public Sub set_auto_logon()
    Dim temp_string As String
    subkey = "Network\Logon"
    temp_string = Trim(Form1.ms1_user_name)
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCECC, opened)
    If retval <> 0 Then Exit Sub
    retval = RegSetValueEx(opened, "Username", 0, 1, ByVal temp_string, Len(temp_string))
    
    subkey = "Software\Microsoft\Windows\CurrentVersion\Winlogon"
    temp_string = Trim(Form1.ms1_password)
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    If (Form1.ms1_auto_logon.Value = vbChecked And retval = 0) Then
        retval = RegSetValueEx(opened, "DefaultPassword", 0, 1, ByVal temp_string, Len(temp_string))
    Else
        retval = RegDeleteValue(opened, "DefaultPassword")
    End If
End Sub

Public Sub get_auto_logon()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Winlogon"
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "DefaultPassword", 0, datatype, ByVal Svalue_read, 255)
    If (retval = 0) Then
        Form1.ms1_auto_logon.Value = vbChecked
        Form1.ms1_password = Trim(Svalue_read)
    End If
    subkey = "Network\Logon"
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCECC, opened)
    retval = RegQueryValueEx(opened, "Username", 0, datatype, ByVal Svalue_read, 255)
    If retval = 0 Then Form1.ms1_user_name = Trim(Svalue_read)
    
End Sub

Public Sub get_disable_cmd()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "Disabled", 0, datatype, Lvalue_read, 4)
    If (retval = 0 And Lvalue_read = 1) Then
        Form1.ms1_disable_cmd.Value = vbChecked
    End If
End Sub

Public Sub set_disable_cmd()
    Dim create_open As Long
    subkey = "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    If (Form1.ms1_disable_cmd.Value = vbChecked) Then
        retval = RegSetValueEx(opened, "Disabled", 0, 4, 1, 4)
    Else
        retval = RegSetValueEx(opened, "Disabled", 0, 4, 0, 4)
    End If
End Sub

Public Sub get_add_cmd()
    Dim vallen As Long
    vallen = 255
    subkey = "Directory\shell\CommandPrompt\command"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    If (retval = 0) Then
        retval = RegQueryValueEx(opened, "", 0, datatype, ByVal Svalue_read, vallen)
        If (Left(Svalue_read, 17) = "command.com /k cd") Then
            Form1.ms1_add_cmd.Value = vbChecked
        End If
    End If
    
End Sub

Public Sub set_add_cmd()
    Dim create_open As Long
    Dim mystring As String
    mystring = "command.com /k cd " + Chr(34) + "%1" + Chr(34)
    subkey = "Directory\shell\CommandPrompt"
    retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    If (retval = 0) Then
        retval = RegSetValueEx(opened, "", 0, 1, ByVal "Command prompt from here", 24)
        subkey = "Directory\shell\CommandPrompt\command"
        retval = RegCreateKeyEx(HKEY_CLASSES_ROOT, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
        If (retval = 0) Then
            If (Form1.ms1_add_cmd.Value = vbChecked) Then
                retval = RegSetValueEx(opened, "", 0, 1, ByVal mystring, Len(mystring))
            Else
                retval = RegDeleteKey(HKEY_CLASSES_ROOT, "Directory\shell\CommandPrompt")
            End If
        End If
    End If
        
End Sub


Public Sub get_source_path()
    subkey = "Software\Microsoft\Windows\CurrentVersion\Setup"
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    If retval <> 0 Then Exit Sub
    retval = RegQueryValueEx(opened, "SourcePath", 0, datatype, ByVal Svalue_read, 255)
    If retval <> 0 Then Exit Sub
    Form1.ms1_install_location = Trim(Svalue_read)
End Sub

Public Sub set_source_path()
    Dim create_open As Long
    Dim tempstring As String
    tempstring = Trim(Form1.ms1_install_location)
    subkey = "Software\Microsoft\Windows\CurrentVersion\Setup"
    retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    If retval <> 0 Then Exit Sub
    retval = RegSetValueEx(opened, "SourcePath", 0, 1, ByVal tempstring, Len(tempstring))
End Sub

Public Sub set_beep()
    Dim create_open As Long
    subkey = "Control Panel\Sound"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    If retval <> 0 Then Exit Sub
    If (Form1.ms1_diable_beep.Value = vbChecked) Then
        retval = RegSetValueEx(opened, "Beep", 0, 1, ByVal "No", 2)
    Else
        retval = RegSetValueEx(opened, "Beep", 0, 1, ByVal "Yes", 3)
    End If
End Sub

Public Sub get_beep()
    subkey = "Control Panel\Sound"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    If retval <> 0 Then Exit Sub
    retval = RegQueryValueEx(opened, "Beep", 0, datatype, ByVal Svalue_read, 255)
    If (retval = 0 And Left(Svalue_read, 2) = "No") Then
        Form1.ms1_diable_beep.Value = vbChecked
    End If
    
End Sub

Public Sub get_min_animation()
    subkey = "Control Panel\Desktop\WindowMetrics"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_ALL_ACCESS, opened)
    If retval <> 0 Then Exit Sub
    retval = RegQueryValueEx(opened, "MinAnimate", 0, datatype, ByVal Svalue_read, 255)
    If (retval = 0 And Left(Svalue_read, 1) = "1") Then
        Form1.ms1_min_animation.Value = vbChecked
    End If
End Sub

Public Sub set_min_animation()
    Dim create_open As Long
    subkey = "Control Panel\Desktop\WindowMetrics"
    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_ALL_ACCESS, secattr, opened, create_open)
    If retval <> o Then Exit Sub
    If (Form1.ms1_min_animation.Value = vbChecked) Then
        retval = RegSetValueEx(opened, "MinAnimate", 0, 1, ByVal "1", 1)
    Else
        retval = RegSetValueEx(opened, "MinAnimate", 0, 1, ByVal "0", 1)
    End If
End Sub

Public Sub get_logon()
    subkey = "Software\Microsoft\Windows\Currentversion\Winlogon"
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    If retval <> 0 Then Exit Sub
    retval = RegQueryValueEx(opened, "LegalNoticeCaption", 0, datatype, ByVal Svalue_read, 255)
    If retval <> 0 Then Exit Sub
    Form1.ms1_legal_caption = Trim(Svalue_read)
    retval = RegQueryValueEx(opened, "LegalNoticeText", 0, datatype, ByVal Svalue_read, 255)
    If (retval <> 0) Then
        Form1.ms1_legal_caption = ""
        Exit Sub
    Else
        Form1.ms1_legal_text = Trim(Svalue_read)
    End If
    Form1.ms1_enable_logon.Value = vbChecked
End Sub

Public Sub set_logon()
    Dim tempstring As String
    subkey = "Software\Microsoft\Windows\Currentversion\Winlogon"
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_ALL_ACCESS, opened)
    If retval <> 0 Then Exit Sub
    If (Form1.ms1_enable_logon.Value = vbChecked And Len(Trim(Form1.ms1_legal_caption)) <> 0 And Len(Trim(Form1.ms1_legal_text) <> 0)) Then
        tempstring = Trim(Form1.ms1_legal_caption)
        retval = RegSetValueEx(opened, "LegalNoticeCaption", 0, 1, ByVal tempstring, Len(tempstring))
        
        If retval <> 0 Then Exit Sub
        
        tempstring = Trim(Form1.ms1_legal_text)
        retval = RegSetValueEx(opened, "LegalNoticeText", 0, 1, ByVal tempstring, Len(tempstring))
    Else
        retval = RegDeleteValue(opened, "LegalNoticeCaption")
        retval = RegDeleteValue(opened, "LegalNoticeText")
    End If
End Sub

Public Sub get_find_history()
    subkey = "Software\Microsoft\Windows\Currentversion\Explorer\Doc Find Spec MRU"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    If (retval = 0) Then
        retval = RegEnumValue(opened, 0, Svalue_read, 255, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then Form1.ms1_clear_find.Enabled = True
    Else
        subkey = "Software\Microsoft\Internet Explorer\Explorer Bars\{C4EE31F3-4768-11D2-BE5C-00A0C9A83DA1}\FilesNamedMRU"
        retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
        If (retval = 0) Then
            retval = RegEnumValue(opened, 0, Svalue_read, 255, 0, datatype, ByVal 0, 0)
            If (retval = 0) Then Form1.ms1_clear_find.Enabled = True
        End If
    End If
End Sub

Public Sub clearfind()
    Dim valuelen As Long
    subkey = "Software\Microsoft\Windows\Currentversion\Explorer\Doc Find Spec MRU"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    While retval = 0
        valuelen = 255
        retval = RegEnumValue(opened, 0, Svalue_read, valuelen, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then
            retval = RegDeleteValue(opened, Left(Svalue_read, valuelen))
        End If
    Wend
    subkey = "Software\Microsoft\Internet Explorer\Explorer Bars\{C4EE31F3-4768-11D2-BE5C-00A0C9A83DA1}\FilesNamedMRU"
    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_QUERY_VALUE, opened)
    While retval = 0
        valuelen = 255
        retval = RegEnumValue(opened, 0, Svalue_read, valuelen, 0, datatype, ByVal 0, 0)
        If (retval = 0) Then
            retval = RegDeleteValue(opened, Left(Svalue_read, valuelen))
        End If
    Wend
    Form1.ms1_clear_find.Enabled = False
End Sub
