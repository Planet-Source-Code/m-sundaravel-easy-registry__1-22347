Attribute VB_Name = "infotip"

Public Sub get_infotip()
    subkey = "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    
    '****** My computer
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_mycomp = Trim(Svalue_read)
    
    '****** My computer
    subkey = "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_bin = Trim(Svalue_read)
    
    '****** My documents
    subkey = "CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_mydocuments = Trim(Svalue_read)
    
    '****** Network Neighbourhood
    subkey = "CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_network = Trim(Svalue_read)
    
    '****** Control Panel
    subkey = "CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_panel = Trim(Svalue_read)
    
    '****** Taskbar & StartMenu
    subkey = "CLSID\{0DF44EAA-FF21-4412-828E-260A8728E7F1}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_smenu = Trim(Svalue_read)
    
    '****** Printers
    subkey = "CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_printers = Trim(Svalue_read)
        
    '****** Folder Options
    subkey = "CLSID\{6DFD7C5C-2451-11d3-A299-00C04F8EF6AF}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_folder_options = Trim(Svalue_read)
    
    '******scheduled tasks
    subkey = "CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_schedule = Trim(Svalue_read)
    
    '******Scanners & Cameras
    subkey = "CLSID\{E211B736-43FD-11D1-9EFB-0000F8757FCD}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    retval = RegQueryValueEx(opened, "InfoTip", 0, datatype, ByVal Svalue_read, 255)
    Form1.it_scanner = Trim(Svalue_read)
End Sub

Public Sub set_infotip()

    Dim temp_string As String
    
    '****** My computer
    subkey = "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_mycomp)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** Recycle Bin
    subkey = "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_bin)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** My documents
    subkey = "CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_mydocuments)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** Network Neighbourhood
    subkey = "CLSID\{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_network)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** Control Panel
    subkey = "CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_panel)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** Taskbar & StartMenu
    subkey = "CLSID\{0DF44EAA-FF21-4412-828E-260A8728E7F1}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_smenu)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '****** Printers
    subkey = "CLSID\{2227A280-3AEA-1069-A2DE-08002B30309D}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_printers)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
        
    '****** Folder Options
    subkey = "CLSID\{6DFD7C5C-2451-11d3-A299-00C04F8EF6AF}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_folder_options)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '******scheduled tasks
    subkey = "CLSID\{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_schedule)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
    
    '******Scanners & Cameras
    subkey = "CLSID\{E211B736-43FD-11D1-9EFB-0000F8757FCD}"
    retval = RegOpenKeyEx(HKEY_CLASSES_ROOT, subkey, 0, KEY_ALL_ACCESS, opened)
    temp_string = Trim(Form1.it_scanner)
    retval = RegSetValueEx(opened, "InfoTip", 0, 1, ByVal temp_string, Len(temp_string))
End Sub
