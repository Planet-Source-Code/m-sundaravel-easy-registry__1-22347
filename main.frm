VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EasyReg"
   ClientHeight    =   7515
   ClientLeft      =   855
   ClientTop       =   780
   ClientWidth     =   10410
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   105
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmap Files(*.bmp)|*.bmp"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabHeight       =   520
      TabCaption(0)   =   "Control Panel"
      TabPicture(0)   =   "main.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame_cu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Information Tips"
      TabPicture(1)   =   "main.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame_infotip"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Look and Feel"
      TabPicture(2)   =   "main.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frame_ms"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Password Management"
      TabPicture(3)   =   "main.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "frame_security"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Internet Explorer (ver 5.0)"
      TabPicture(4)   =   "main.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame12"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "browse2"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "ie_systempic"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Browse1"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "ie_toolpic"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "ie_caption"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label10(2)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label10(1)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label10(0)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Miscellaneous"
      TabPicture(5)   =   "main.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame20"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame19"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame17"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame16"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Frame15"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Frame13"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      Begin VB.Frame Frame20 
         Caption         =   "Logon Message"
         ForeColor       =   &H000000FF&
         Height          =   1665
         Left            =   -70050
         TabIndex        =   158
         Top             =   4725
         Width           =   4725
         Begin VB.CheckBox ms1_enable_logon 
            Caption         =   "Enable the logon screen"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   390
            TabIndex        =   163
            Top             =   285
            Width           =   2025
         End
         Begin VB.TextBox ms1_legal_text 
            Enabled         =   0   'False
            Height          =   345
            Left            =   990
            TabIndex        =   162
            ToolTipText     =   "Reboot Required"
            Top             =   1110
            Width           =   3465
         End
         Begin VB.TextBox ms1_legal_caption 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1005
            TabIndex        =   161
            ToolTipText     =   "Reboot Required"
            Top             =   660
            Width           =   3465
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Text:"
            Height          =   195
            Left            =   570
            TabIndex        =   160
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Caption:"
            Height          =   195
            Left            =   375
            TabIndex        =   159
            Top             =   735
            Width           =   585
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Installation path"
         ForeColor       =   &H000000FF&
         Height          =   1755
         Left            =   -74850
         TabIndex        =   154
         Top             =   4635
         Width           =   4665
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            TabIndex        =   155
            Top             =   975
            Width           =   375
         End
         Begin VB.TextBox ms1_install_location 
            Enabled         =   0   'False
            Height          =   345
            Left            =   195
            TabIndex        =   156
            ToolTipText     =   "No Reboot Required"
            Top             =   945
            Width           =   4290
         End
         Begin VB.Label Label13 
            Caption         =   "Location of windows Installation files:"
            Height          =   225
            Left            =   180
            TabIndex        =   157
            Top             =   615
            Width           =   2640
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Miscellaneous"
         ForeColor       =   &H000000FF&
         Height          =   1785
         Left            =   -70065
         TabIndex        =   139
         Top             =   2760
         Width           =   4800
         Begin VB.CheckBox ms1_min_animation 
            Caption         =   "Minimize the animation to improve performance"
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   435
            TabIndex        =   141
            ToolTipText     =   "Reboot Required"
            Top             =   915
            Width           =   3660
         End
         Begin VB.CheckBox ms1_diable_beep 
            Caption         =   "Diable the Beep Sound on errors"
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   435
            TabIndex        =   140
            ToolTipText     =   "Reboot Required"
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "History Clearence"
         ForeColor       =   &H000000FF&
         Height          =   1770
         Left            =   -74835
         TabIndex        =   136
         Top             =   2760
         Width           =   4665
         Begin VB.CommandButton ms1_clear_find 
            Caption         =   "Clear Find History "
            Enabled         =   0   'False
            Height          =   390
            Left            =   1080
            TabIndex        =   142
            ToolTipText     =   "No Reboot Required"
            Top             =   750
            Width           =   2940
         End
         Begin VB.CommandButton ms1_clear_ie 
            Caption         =   "Clear IE History"
            Enabled         =   0   'False
            Height          =   390
            Left            =   1080
            TabIndex        =   138
            ToolTipText     =   "No Reboot Required"
            Top             =   1215
            Width           =   2940
         End
         Begin VB.CommandButton ms1_clear_run 
            Caption         =   "Clear the Run menu History"
            Enabled         =   0   'False
            Height          =   390
            Left            =   1080
            TabIndex        =   137
            ToolTipText     =   "No Reboot Required"
            Top             =   300
            Width           =   2940
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Command Prompt"
         ForeColor       =   &H000000FF&
         Height          =   1875
         Left            =   -74820
         TabIndex        =   133
         Top             =   810
         Width           =   4635
         Begin VB.CheckBox ms1_add_cmd 
            Caption         =   "Add command promt to context menu"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   285
            TabIndex        =   135
            ToolTipText     =   "No Reboot Required"
            Top             =   1200
            Width           =   2985
         End
         Begin VB.CheckBox ms1_disable_cmd 
            Caption         =   "Disable Command Prompt"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   285
            TabIndex        =   134
            ToolTipText     =   "No Reboot Required"
            Top             =   615
            Width           =   2175
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Automatic Logon"
         ForeColor       =   &H000000FF&
         Height          =   1860
         Left            =   -70065
         TabIndex        =   127
         Top             =   825
         Width           =   4800
         Begin VB.TextBox ms1_password 
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1395
            PasswordChar    =   "*"
            TabIndex        =   132
            ToolTipText     =   "Reboot Required"
            Top             =   1260
            Width           =   3225
         End
         Begin VB.TextBox ms1_user_name 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1395
            TabIndex        =   131
            ToolTipText     =   "Reboot Required"
            Top             =   780
            Width           =   3225
         End
         Begin VB.CheckBox ms1_auto_logon 
            Caption         =   "Enable Automatic Logon"
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   450
            TabIndex        =   128
            Top             =   285
            Width           =   2100
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Height          =   195
            Left            =   495
            TabIndex        =   130
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "User Name:"
            Height          =   195
            Left            =   435
            TabIndex        =   129
            Top             =   855
            Width           =   840
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Toolbar Picture"
         ForeColor       =   &H000000FF&
         Height          =   1680
         Left            =   -66945
         TabIndex        =   123
         Top             =   1905
         Width           =   1725
         Begin VB.CheckBox ie_enable_windows 
            Caption         =   "ChangeWindows"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   60
            TabIndex        =   125
            ToolTipText     =   "No Reboot Required"
            Top             =   1350
            Width           =   1530
         End
         Begin VB.CheckBox ie_enable_explorer 
            Caption         =   "Change Explorer"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   90
            TabIndex        =   124
            ToolTipText     =   "No Reboot Required"
            Top             =   465
            Width           =   1530
         End
      End
      Begin VB.Frame Frame12 
         ForeColor       =   &H8000000E&
         Height          =   2655
         Left            =   -74655
         TabIndex        =   110
         Top             =   3690
         Width           =   9225
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable full screen Mode"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   14
            Left            =   6480
            TabIndex        =   122
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1980
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable open in file menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   10
            Left            =   6480
            TabIndex        =   121
            ToolTipText     =   "Restarting of IE Required"
            Top             =   225
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable New in file menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   9
            Left            =   6480
            TabIndex        =   120
            ToolTipText     =   "Restarting of IE Required"
            Top             =   810
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the Favourites menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   8
            Left            =   3600
            TabIndex        =   119
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1980
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the Save as option"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   7
            Left            =   3600
            TabIndex        =   118
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1395
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the Internet Options menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   6
            Left            =   3600
            TabIndex        =   117
            ToolTipText     =   "Restarting of IE Required"
            Top             =   810
            Width           =   2745
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the right click menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   5
            Left            =   3600
            TabIndex        =   116
            ToolTipText     =   "Restarting of IE Required"
            Top             =   225
            Width           =   2325
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable closing of Browser "
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   4
            Left            =   6480
            TabIndex        =   115
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1395
            Width           =   2235
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the expansion of New menu"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   3
            Left            =   210
            TabIndex        =   114
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1980
            Width           =   2985
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable Script Debugger in IE"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   2
            Left            =   210
            TabIndex        =   113
            ToolTipText     =   "Restarting of IE Required"
            Top             =   1395
            Width           =   2445
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Disable the auto complete feature in IE"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   1
            Left            =   210
            TabIndex        =   112
            ToolTipText     =   "Restarting of IE Required"
            Top             =   810
            Width           =   3045
         End
         Begin VB.CheckBox ie_options 
            Caption         =   "Hide IE icon in the desktop"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   0
            Left            =   225
            TabIndex        =   111
            ToolTipText     =   "Restarting of IE Required"
            Top             =   225
            Width           =   2265
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   3
            X1              =   6390
            X2              =   6390
            Y1              =   90
            Y2              =   2640
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   2
            X1              =   6405
            X2              =   6405
            Y1              =   105
            Y2              =   2625
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   1
            X1              =   3390
            X2              =   3390
            Y1              =   90
            Y2              =   2640
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   0
            X1              =   3405
            X2              =   3405
            Y1              =   105
            Y2              =   2625
         End
      End
      Begin VB.CommandButton browse2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67455
         TabIndex        =   109
         Top             =   3225
         Width           =   375
      End
      Begin VB.TextBox ie_systempic 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74640
         TabIndex        =   107
         ToolTipText     =   "Restarting of IE Required"
         Top             =   3195
         Width           =   7575
      End
      Begin VB.CommandButton Browse1 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67470
         TabIndex        =   106
         Top             =   2325
         Width           =   375
      End
      Begin VB.TextBox ie_toolpic 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74640
         TabIndex        =   105
         ToolTipText     =   "Restarting of IE Required"
         Top             =   2295
         Width           =   7575
      End
      Begin VB.TextBox ie_caption 
         Height          =   375
         Left            =   -74655
         TabIndex        =   103
         ToolTipText     =   "Restarting of IE Required"
         Top             =   1440
         Width           =   9405
      End
      Begin VB.Frame frame_infotip 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   5355
         Left            =   -74730
         TabIndex        =   80
         Top             =   1110
         Width           =   9255
         Begin VB.TextBox it_mycomp 
            Height          =   345
            Left            =   120
            TabIndex        =   90
            ToolTipText     =   "No Reboot Required"
            Top             =   1065
            Width           =   3855
         End
         Begin VB.TextBox it_bin 
            Height          =   345
            Left            =   150
            TabIndex        =   89
            ToolTipText     =   "No Reboot Required"
            Top             =   1995
            Width           =   3855
         End
         Begin VB.TextBox it_mydocuments 
            Height          =   345
            Left            =   120
            TabIndex        =   88
            ToolTipText     =   "No Reboot Required"
            Top             =   2985
            Width           =   3855
         End
         Begin VB.TextBox it_network 
            Height          =   345
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "No Reboot Required"
            Top             =   3915
            Width           =   3855
         End
         Begin VB.TextBox it_panel 
            Height          =   345
            Left            =   135
            TabIndex        =   86
            ToolTipText     =   "No Reboot Required"
            Top             =   4815
            Width           =   3855
         End
         Begin VB.TextBox it_smenu 
            Height          =   345
            Left            =   4860
            TabIndex        =   85
            ToolTipText     =   "No Reboot Required"
            Top             =   1065
            Width           =   3855
         End
         Begin VB.TextBox it_printers 
            Height          =   345
            Left            =   4860
            TabIndex        =   84
            ToolTipText     =   "No Reboot Required"
            Top             =   1995
            Width           =   3855
         End
         Begin VB.TextBox it_folder_options 
            Height          =   345
            Left            =   4860
            TabIndex        =   83
            ToolTipText     =   "No Reboot Required"
            Top             =   2985
            Width           =   3855
         End
         Begin VB.TextBox it_schedule 
            Height          =   345
            Left            =   4860
            TabIndex        =   82
            ToolTipText     =   "No Reboot Required"
            Top             =   3915
            Width           =   3855
         End
         Begin VB.TextBox it_scanner 
            Height          =   345
            Left            =   4860
            TabIndex        =   81
            ToolTipText     =   "No Reboot Required"
            Top             =   4815
            Width           =   3855
         End
         Begin VB.Label Label6 
            Caption         =   $"main.frx":04EA
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   435
            Left            =   105
            TabIndex        =   101
            Top             =   45
            Width           =   9000
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "My Computer:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   100
            Top             =   825
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Recycle Bin:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   99
            Top             =   1725
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "My Documents:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   98
            Top             =   2685
            Width           =   1110
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Network Neighbourhood:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   97
            Top             =   3645
            Width           =   1785
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Taskbar && Start Meu:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   4860
            TabIndex        =   96
            Top             =   825
            Width           =   1500
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Printers:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   4860
            TabIndex        =   95
            Top             =   1725
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Folder Options:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   4860
            TabIndex        =   94
            Top             =   2745
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Scheduled Tasks:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   7
            Left            =   4860
            TabIndex        =   93
            Top             =   3645
            Width           =   1290
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Scanners && Cameras:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   4860
            TabIndex        =   92
            Top             =   4545
            Width           =   1515
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Control Panel:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   91
            Top             =   4545
            Width           =   990
         End
      End
      Begin VB.Frame frame_cu 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74715
         TabIndex        =   51
         Top             =   870
         Width           =   9255
         Begin VB.Frame Frame1 
            Caption         =   "Display Settings"
            ForeColor       =   &H000000FF&
            Height          =   2085
            Left            =   90
            TabIndex        =   75
            Top             =   855
            Width           =   2505
            Begin VB.CheckBox cu_display 
               Caption         =   "Hide the Settings page"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   3
               Left            =   45
               TabIndex        =   79
               ToolTipText     =   "No Reboot Required"
               Top             =   1680
               Width           =   2325
            End
            Begin VB.CheckBox cu_display 
               Caption         =   "Hide the Screen saver page"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   75
               TabIndex        =   78
               ToolTipText     =   "No Reboot Required"
               Top             =   1230
               Width           =   2325
            End
            Begin VB.CheckBox cu_display 
               Caption         =   "Hide the Background page"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   75
               TabIndex        =   77
               ToolTipText     =   "No Reboot Required"
               Top             =   795
               Width           =   2325
            End
            Begin VB.CheckBox cu_display 
               Caption         =   "Hide the appearence page"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   90
               TabIndex        =   76
               ToolTipText     =   "No Reboot Required"
               Top             =   285
               Width           =   2325
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Network"
            ForeColor       =   &H000000FF&
            Height          =   2085
            Left            =   2640
            TabIndex        =   71
            Top             =   855
            Width           =   3375
            Begin VB.CheckBox cu_network 
               Caption         =   "Hide File and Printer Sharing Controls "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   90
               TabIndex        =   74
               ToolTipText     =   "No Reboot Required"
               Top             =   285
               Width           =   2955
            End
            Begin VB.CheckBox cu_network 
               Caption         =   "Hide Network Identification Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   90
               TabIndex        =   73
               ToolTipText     =   "No Reboot Required"
               Top             =   750
               Width           =   2685
            End
            Begin VB.CheckBox cu_network 
               Caption         =   "Hide Network Access Control Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   72
               ToolTipText     =   "No Reboot Required"
               Top             =   1260
               Width           =   2805
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Passwords"
            ForeColor       =   &H000000FF&
            Height          =   2085
            Left            =   6060
            TabIndex        =   67
            Top             =   855
            Width           =   3105
            Begin VB.CheckBox cu_passwords 
               Caption         =   "Hide the Remote Administration Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   90
               TabIndex        =   70
               ToolTipText     =   "No Reboot Required"
               Top             =   285
               Width           =   2955
            End
            Begin VB.CheckBox cu_passwords 
               Caption         =   "Hide the User Profiles Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   105
               TabIndex        =   69
               ToolTipText     =   "No Reboot Required"
               Top             =   765
               Width           =   2955
            End
            Begin VB.CheckBox cu_passwords 
               Caption         =   "Hide the Change Passwords Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   105
               TabIndex        =   68
               ToolTipText     =   "No Reboot Required"
               Top             =   1260
               Width           =   2955
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Printers"
            ForeColor       =   &H000000FF&
            Height          =   2535
            Left            =   2640
            TabIndex        =   63
            Top             =   3000
            Width           =   3375
            Begin VB.CheckBox cu_printers 
               Caption         =   "Disable the Addition of Printers "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   90
               TabIndex        =   66
               ToolTipText     =   "No Reboot Required"
               Top             =   285
               Width           =   2955
            End
            Begin VB.CheckBox cu_printers 
               Caption         =   "Disable the Deletion of Printers "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   90
               TabIndex        =   65
               ToolTipText     =   "No Reboot Required"
               Top             =   757
               Width           =   2955
            End
            Begin VB.CheckBox cu_printers 
               Caption         =   "Hide General and Details Printer Pages "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   64
               ToolTipText     =   "No Reboot Required"
               Top             =   1230
               Width           =   3060
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "System"
            ForeColor       =   &H000000FF&
            Height          =   2535
            Left            =   90
            TabIndex        =   58
            Top             =   3000
            Width           =   2505
            Begin VB.CheckBox cu_system 
               Caption         =   "Hide Hardware Profiles Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   120
               TabIndex        =   62
               ToolTipText     =   "No Reboot Required"
               Top             =   285
               Width           =   2340
            End
            Begin VB.CheckBox cu_system 
               Caption         =   "Hide  Device Manager Page "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   135
               TabIndex        =   61
               ToolTipText     =   "No Reboot Required"
               Top             =   745
               Width           =   2355
            End
            Begin VB.CheckBox cu_system 
               Caption         =   "Hide the File System Button "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   120
               TabIndex        =   60
               ToolTipText     =   "No Reboot Required"
               Top             =   1205
               Width           =   2325
            End
            Begin VB.CheckBox cu_system 
               Caption         =   "Hide Virtual Memory Button "
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   3
               Left            =   120
               TabIndex        =   59
               ToolTipText     =   "No Reboot Required"
               Top             =   1665
               Width           =   2325
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Start Menu"
            ForeColor       =   &H000000FF&
            Height          =   2535
            Left            =   6075
            TabIndex        =   54
            Top             =   3000
            Width           =   3075
            Begin VB.CheckBox cu_start_menu 
               Caption         =   "Dont add files to Recent files"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   4
               Left            =   165
               TabIndex        =   126
               ToolTipText     =   "Reboot Required"
               Top             =   1680
               Width           =   2385
            End
            Begin VB.CheckBox cu_start_menu 
               Caption         =   "Hide the Run command"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   165
               TabIndex        =   57
               ToolTipText     =   "Reboot Required"
               Top             =   240
               Width           =   2325
            End
            Begin VB.CheckBox cu_start_menu 
               Caption         =   "Hide the Settings Menu"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   165
               TabIndex        =   56
               ToolTipText     =   "Reboot Required"
               Top             =   720
               Width           =   2325
            End
            Begin VB.CheckBox cu_start_menu 
               Caption         =   "Hide the Find Command"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   3
               Left            =   165
               TabIndex        =   55
               ToolTipText     =   "Reboot Required"
               Top             =   1185
               Width           =   2325
            End
         End
         Begin VB.CheckBox cu_cp 
            Caption         =   "Hide Control Panel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   2385
            TabIndex        =   53
            ToolTipText     =   "Reboot Required"
            Top             =   270
            Width           =   1965
         End
         Begin VB.CheckBox disable_registry 
            Caption         =   "Disable Registry Editing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   5048
            TabIndex        =   52
            ToolTipText     =   "No Reboot Required"
            Top             =   270
            Width           =   2355
         End
      End
      Begin VB.Frame frame_security 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   120
         TabIndex        =   44
         Top             =   915
         Width           =   9735
         Begin VB.Frame Frame18 
            Caption         =   "Password Recovery"
            ForeColor       =   &H000000FF&
            Height          =   1920
            Left            =   60
            TabIndex        =   143
            Top             =   3420
            Width           =   9525
            Begin VB.TextBox txtoutput 
               Enabled         =   0   'False
               Height          =   345
               Left            =   570
               TabIndex        =   146
               Top             =   1305
               Width           =   4665
            End
            Begin VB.PictureBox Picture1 
               Height          =   1530
               Left            =   5835
               ScaleHeight     =   1470
               ScaleWidth      =   1740
               TabIndex        =   144
               Top             =   240
               Width           =   1800
               Begin VB.Image imgTarget 
                  Height          =   480
                  Left            =   585
                  Picture         =   "main.frx":0579
                  Stretch         =   -1  'True
                  Top             =   630
                  Width           =   540
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Drag This Icon"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   330
                  TabIndex        =   145
                  Top             =   330
                  Width           =   1050
               End
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Output:"
               Height          =   195
               Left            =   570
               TabIndex        =   153
               Top             =   1020
               Width           =   525
            End
            Begin VB.Label lblRelease 
               AutoSize        =   -1  'True
               Caption         =   "Target Found! Release the Mouse"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   525
               TabIndex        =   152
               Top             =   495
               Visible         =   0   'False
               Width           =   4800
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Targeting:"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7800
               TabIndex        =   151
               Top             =   360
               Width           =   720
            End
            Begin VB.Label lblXCap 
               AutoSize        =   -1  'True
               Caption         =   "X:"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7905
               TabIndex        =   150
               Top             =   750
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label lblYCap 
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7890
               TabIndex        =   149
               Top             =   1050
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               Caption         =   "12"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   8190
               TabIndex        =   148
               Top             =   750
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Label lblY 
               AutoSize        =   -1  'True
               Caption         =   "12"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   8190
               TabIndex        =   147
               Top             =   1050
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Image imgcross 
               Height          =   480
               Left            =   8880
               Picture         =   "main.frx":0883
               Top             =   1170
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image imgNull 
               Height          =   225
               Left            =   8820
               Top             =   315
               Width           =   285
            End
         End
         Begin VB.TextBox txt_cpass 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1995
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   1245
            Width           =   4365
         End
         Begin VB.TextBox txt_npass 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1995
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1800
            Width           =   4365
         End
         Begin VB.CommandButton chg_pass 
            Caption         =   "&Change Password"
            Enabled         =   0   'False
            Height          =   435
            Left            =   7080
            TabIndex        =   46
            Top             =   1845
            Width           =   1995
         End
         Begin VB.TextBox txt_rpass 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1995
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   2340
            Width           =   4365
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Password Protect this application"
            Height          =   405
            Left            =   6735
            TabIndex        =   45
            Top             =   1350
            Value           =   1  'Checked
            Width           =   2685
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Current Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   50
            Top             =   1335
            Width           =   1560
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "New Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   49
            Top             =   1875
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Retype Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   48
            Top             =   2430
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   $"main.frx":0B8D
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   390
            Left            =   180
            TabIndex        =   47
            Top             =   75
            Width           =   8955
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frame_ms 
         BorderStyle     =   0  'None
         Height          =   5805
         Left            =   -74910
         TabIndex        =   8
         Top             =   660
         Width           =   9255
         Begin VB.Frame Frame2 
            Caption         =   "Active Desktop"
            ForeColor       =   &H000000FF&
            Height          =   5325
            Left            =   105
            TabIndex        =   35
            Top             =   435
            Width           =   3240
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable Wallpaper changing"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   0
               Left            =   240
               TabIndex        =   43
               ToolTipText     =   "No Reboot Required"
               Top             =   615
               Width           =   2295
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable Components"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   1
               Left            =   240
               TabIndex        =   42
               ToolTipText     =   "No Reboot Required"
               Top             =   1170
               Width           =   2295
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable Ability to add components"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   2
               Left            =   240
               TabIndex        =   41
               ToolTipText     =   "No Reboot Required"
               Top             =   1710
               Width           =   2805
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable Ability to delete components"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   3
               Left            =   240
               TabIndex        =   40
               ToolTipText     =   "No Reboot Required"
               Top             =   2265
               Width           =   2895
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable Ability to edit components"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   4
               Left            =   240
               TabIndex        =   39
               ToolTipText     =   "No Reboot Required"
               Top             =   2805
               Width           =   2715
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable closing dragdrop bands"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   5
               Left            =   240
               TabIndex        =   38
               ToolTipText     =   "No Reboot Required"
               Top             =   3360
               Width           =   2595
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Disable moving bands"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   6
               Left            =   240
               TabIndex        =   37
               ToolTipText     =   "No Reboot Required"
               Top             =   3900
               Width           =   2295
            End
            Begin VB.CheckBox ms_ad 
               Caption         =   "Restrict changes to Active Desktop"
               ForeColor       =   &H00FF0000&
               Height          =   345
               Index           =   7
               Left            =   240
               TabIndex        =   36
               ToolTipText     =   "No Reboot Required"
               Top             =   4455
               Width           =   2835
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "My computer"
            ForeColor       =   &H000000FF&
            Height          =   1755
            Left            =   6915
            TabIndex        =   31
            Top             =   450
            Width           =   2250
            Begin VB.OptionButton cu_drives 
               Caption         =   "Show all Drives"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   150
               TabIndex        =   34
               ToolTipText     =   "Reboot Required"
               Top             =   1365
               Width           =   1425
            End
            Begin VB.OptionButton cu_drives 
               Caption         =   "Hide all drives"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   33
               ToolTipText     =   "Reboot Required"
               Top             =   915
               Width           =   1305
            End
            Begin VB.OptionButton cu_drives 
               Caption         =   "Hide floppy"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   32
               ToolTipText     =   "Reboot Required"
               Top             =   420
               Width           =   1125
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Mouse"
            ForeColor       =   &H000000FF&
            Height          =   1320
            Left            =   6900
            TabIndex        =   28
            Top             =   2280
            Width           =   2265
            Begin VB.CheckBox ms_mouse 
               Caption         =   "Active window tracking "
               ForeColor       =   &H00FF0000&
               Height          =   405
               Index           =   0
               Left            =   180
               TabIndex        =   30
               ToolTipText     =   "Reboot Required"
               Top             =   300
               Width           =   1980
            End
            Begin VB.CheckBox ms_mouse 
               Caption         =   "Live Scrolling in Word"
               ForeColor       =   &H00FF0000&
               Height          =   405
               Index           =   1
               Left            =   180
               TabIndex        =   29
               ToolTipText     =   "Restarting of Office applications required"
               Top             =   810
               Width           =   1920
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Menus and graphics"
            ForeColor       =   &H000000FF&
            Height          =   5355
            Left            =   3420
            TabIndex        =   14
            Top             =   435
            Width           =   3405
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Disable smart menu in MS Office 2000"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   24
               ToolTipText     =   "Restarting of Office applications required"
               Top             =   495
               Width           =   3015
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Hide all items in the Desktop"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   105
               TabIndex        =   23
               ToolTipText     =   "Reboot Required"
               Top             =   928
               Width           =   2325
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Auto Refresh windows"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   22
               ToolTipText     =   "Reboot Required"
               Top             =   1361
               Width           =   3015
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Automatically view thumbnails of bitmaps"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   90
               TabIndex        =   21
               ToolTipText     =   "Reboot Required"
               Top             =   1794
               Width           =   3165
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Remove arrow in shortcuts"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   90
               TabIndex        =   20
               ToolTipText     =   "Reboot Required"
               Top             =   2227
               Width           =   2265
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Drag full window"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   90
               TabIndex        =   19
               ToolTipText     =   "Reboot Required"
               Top             =   2660
               Width           =   1515
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Disable ""Click here to start"" banner"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   6
               Left            =   90
               TabIndex        =   18
               ToolTipText     =   "Reboot Required"
               Top             =   3093
               Width           =   3045
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Hide the ""New"" Menu"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   7
               Left            =   90
               TabIndex        =   16
               ToolTipText     =   "No Reboot Required"
               Top             =   3526
               Width           =   3045
            End
            Begin VB.CheckBox ms_mnu 
               Caption         =   "Hide the ""Send to"" Menu"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   8
               Left            =   90
               TabIndex        =   15
               ToolTipText     =   "No Reboot Required"
               Top             =   3960
               Width           =   3045
            End
            Begin MSComctlLib.Slider menu_speed 
               Height          =   360
               Left            =   315
               TabIndex        =   17
               ToolTipText     =   "Reboot Required"
               Top             =   4740
               Width           =   2775
               _ExtentX        =   4921
               _ExtentY        =   635
               _Version        =   393216
               Max             =   300
               TickStyle       =   3
            End
            Begin VB.Label Label5 
               Caption         =   "Menu display speed in milli seconds"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   405
               TabIndex        =   27
               Top             =   4500
               Width           =   2655
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Speeder"
               Height          =   195
               Left            =   300
               TabIndex        =   26
               Top             =   5115
               Width           =   600
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Slower"
               Height          =   195
               Left            =   2550
               TabIndex        =   25
               Top             =   5115
               Width           =   480
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Computer Owner Information"
            ForeColor       =   &H000000FF&
            Height          =   2055
            Left            =   6885
            TabIndex        =   9
            Top             =   3705
            Width           =   2265
            Begin VB.TextBox txt_owner 
               Height          =   345
               Left            =   150
               TabIndex        =   11
               ToolTipText     =   "No Reboot Required"
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txt_place 
               Height          =   315
               Left            =   150
               TabIndex        =   10
               ToolTipText     =   "No Reboot Required"
               Top             =   1500
               Width           =   1995
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Left            =   150
               TabIndex        =   13
               Top             =   360
               Width           =   465
            End
            Begin VB.Label Label2 
               Caption         =   "Organisation:"
               Height          =   255
               Left            =   135
               TabIndex        =   12
               Top             =   1230
               Width           =   975
            End
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Windows Explorer toolbar Background Bitmap:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   108
         Top             =   2925
         Width           =   3285
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Internet Explorer toolbar Background Bitmap:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -74580
         TabIndex        =   104
         Top             =   2025
         Width           =   3165
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Internet Explorer Window Caption:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -74580
         TabIndex        =   102
         Top             =   1170
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmd_about 
      Caption         =   "&About"
      Height          =   405
      Left            =   210
      TabIndex        =   6
      Top             =   7020
      Width           =   1305
   End
   Begin VB.CommandButton cmd_apply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   405
      Left            =   8355
      TabIndex        =   5
      Top             =   7005
      Width           =   1305
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   7005
      TabIndex        =   4
      Top             =   7020
      Width           =   1305
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   5625
      TabIndex        =   0
      Top             =   7020
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim setting As New cp
Dim setting_ms As New ms

Private Sub browse2_Click()
    Dim temp_string As String
    temp_string = ie_systempic
    cd.ShowOpen
    If (cd.FileName <> temp_string) Then
        ie_systempic = cd.FileName
    End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        txt_cpass.Enabled = True
        txt_npass.Enabled = True
        txt_rpass.Enabled = True
        'chg_pass.Enabled = True
    Else
        txt_cpass.Enabled = False
        txt_npass.Enabled = False
        txt_rpass.Enabled = False
        'chg_pass.Enabled = False
    End If
    cmd_apply.Enabled = True

End Sub

Private Sub chg_pass_Click()
    Dim temp_read As String
    Dim slength As Long
    Dim to_write As String
    Dim to_check As String
    temp_read = Space(255)
    slength = GetProfileString("PhotoOptics", "sclassname", "N", temp_read, 255)
    temp_read = Left(temp_read, slength)
    If (temp_read = "N") Then
        'MsgBox "reded is n"
        If (txt_cpass <> "PleaseLogin") Then
            MsgBox "Please enter the current password correctly", vbCritical, "Error"
            Exit Sub
        Else
            If (Trim(txt_npass) <> Trim(txt_rpass)) Then
                MsgBox "New and Retyped passwords doesn't match", vbCritical, "Error"
                Exit Sub
            End If
            
            to_write = encrypt(Trim(txt_npass))
            Call WriteProfileString("PhotoOptics", "sclassname", to_write)
            MsgBox "Password successfully changed", vbInformation, "Password changed"
            Exit Sub
        End If
    Else
        to_read = decrypt(temp_read)
        If (to_read <> txt_cpass) Then
            MsgBox "Please enter the current password correctly", vbCritical, "Error"
            Exit Sub
        Else
            If (Trim(txt_npass) <> Trim(txt_rpass)) Then
                MsgBox "New and Retyped passwords doesn't match", vbCritical, "Error"
                Exit Sub
            End If
            
            to_write = encrypt(Trim(txt_npass))
            Call WriteProfileString("PhotoOptics", "sclassname", to_write)
            MsgBox "Password successfully changed", vbInformation, "Password changed"
            Exit Sub
        End If
    End If
End Sub

Private Sub cmd_about_Click()
    frmAbout.Show 1
End Sub

Private Sub cmd_apply_Click()

Call setting.set_cp
Call setting.set_display
Call setting.set_network
Call setting.set_passwords
Call setting.set_printers
Call setting.set_start_menu
Call setting.set_system
Call setting.set_disable

Call setting_ms.set_drives
Call setting_ms.set_owner
Call setting_ms.set_mouse
Call setting_ms.set_menu_graphics
Call setting_ms.set_active_desktop

Call set_infotip
Call set_ie
Call set_auto_logon
Call set_disable_cmd
Call set_add_cmd
Call set_source_path
Call set_beep
Call set_min_animation
Call set_logon

If Check1.Value = vbChecked Then
    Call WriteProfileString("intl", "sType", "1")
Else
    Call WriteProfileString("intl", "sType", "0")
End If


cmd_apply.Enabled = False

End Sub

Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_ok_Click()
Call setting.set_cp
Call setting.set_display
Call setting.set_network
Call setting.set_passwords
Call setting.set_printers
Call setting.set_start_menu
Call setting.set_system
Call setting.set_disable

Call setting_ms.set_drives
Call setting_ms.set_owner
Call setting_ms.set_mouse
Call setting_ms.set_menu_graphics
Call setting_ms.set_active_desktop

Call set_infotip
Call set_ie
Call set_auto_logon
Call set_disable_cmd
Call set_add_cmd
Call set_source_path
Call set_beep
Call set_min_animation
Call set_logon

'********** DETERMINING ENABLING PASSWORD
If Check1.Value = vbChecked Then
    Call WriteProfileString("intl", "sType", "1")
Else
    Call WriteProfileString("intl", "sType", "0")
End If

If (vbOK = MsgBox("Some of the settings may require the system to reboot.Do you want to restart now?", vbInformation + vbOKCancel, "Confirmation")) Then
    Call ExitWindowsEx(2, 0)
End If
End
End Sub

Private Sub cp_show_Click()
    
End Sub


Private Sub Browse1_Click()
    Dim temp_string As String
    temp_string = ie_toolpic
    cd.ShowOpen
    If (cd.FileName <> temp_string) Then
        ie_toolpic = cd.FileName
    End If
End Sub

Private Sub Command1_Click()
    dir_select.Show 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cu_cp_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub cu_display_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub cu_drives_Click(index As Integer)
    cmd_apply.Enabled = True
End Sub

Private Sub cu_network_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub cu_passwords_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub cu_printers_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub cu_start_menu_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub cu_system_Click(index As Integer)
cmd_apply.Enabled = True
End Sub

Private Sub disable_registry_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub Form_Activate()
    If (flag = 0) Then
        cmd_apply.Enabled = False
        flag = 1
    End If
    End Sub



Private Sub Form_Load()
       
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = True
    
    Call setting.get_cp
    Call setting.get_display
    Call setting.get_network
    Call setting.get_passwords
    Call setting.get_printers
    Call setting.get_start_menu
    Call setting.get_system
    Call setting.get_disable
    
    Call setting_ms.get_drives
    Call setting_ms.get_owner
    Call setting_ms.get_mouse
    Call setting_ms.get_menu_graphics
    Call setting_ms.get_active_desktop

    Call get_infotip
    Call get_ie
    Call get_run_history
    Call get_ie_history
    Call get_auto_logon
    Call get_disable_cmd
    Call get_add_cmd
    Call get_source_path
    Call get_beep
    Call get_min_animation
    Call get_logon
    Call get_find_history
    
    cmd_apply.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub ie_caption_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub ie_enable_explorer_Click()
    If (ie_enable_explorer.Value = vbChecked) Then
        Browse1.Enabled = True
    Else
        ie_toolpic = "<Default>"
        Browse1.Enabled = False
    End If
    cmd_apply.Enabled = True
End Sub

Private Sub ie_enable_windows_Click()
    If (ie_enable_windows.Value = vbChecked) Then
        browse2.Enabled = True
    Else
        ie_systempic = "<Default>"
        browse2.Enabled = False
    End If
    cmd_apply.Enabled = True
End Sub

Private Sub ie_options_Click(index As Integer)
    cmd_apply.Enabled = True
End Sub

Private Sub ie_scope_Click(index As Integer)
    cmd_apply.Enabled = True
    If (index = 0) Then
        IE_KEY = HKEY_CURRENT_USER
    Else
        IE_KEY = HKEY_LOCAL_MACHINE
    End If
End Sub

Private Sub ie_systempic_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub ie_toolpic_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Targeting = True
    imgTarget.Picture = imgNull.Picture
    lblXCap.Visible = True
    lblYCap.Visible = True
    lblX.Visible = True
    lblY.Visible = True
    
    Me.MousePointer = 99
    Me.MouseIcon = imgcross.Picture
    txtoutput.Text = ""
End Sub

Private Sub imgTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sName As String, sClassName As String * 255, TempHwnd As Long
    
    If Targeting = False Then Exit Sub
    Call GetCursorPos(CursorPosition)

    ' Display mouse's location
    lblX.Caption = CursorPosition.X
    lblY.Caption = CursorPosition.Y

    ' Check whether the cursor is pointing to a TextBox or not
    TempHwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y)
    Call GetClassName(TempHwnd, sClassName, 255)
    sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))
    
    If sName = "Edit" Or InStr(sName, "TextBox") > 0 Then
        lblRelease.Visible = True
    Else
        lblRelease.Visible = False
    End If
End Sub

Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TargetLen As Long, tempstring As String, hwnd As Long
' Visual effects
    Targeting = False
    lblXCap.Visible = False
    lblYCap.Visible = False
    lblX.Visible = False
    lblY.Visible = False
    lblRelease.Visible = False
    
    imgTarget.Picture = imgcross.Picture
    Me.MousePointer = 0

'
Call GetCursorPos(CursorPosition)
    hwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y) ' Get target window's handle
    hwnd = GetTopLevelParent(hwnd)          ' Get target window's parent's handle
    TargetLen& = SendMessage(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
    tempstring$ = String(TargetLen&, 0&)
    Call sendmessagebystring(hwnd&, WM_GETTEXT, TargetLen& + 1, tempstring$)
    txtoutput.Text = tempstring$
End Sub

Private Sub it_bin_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub it_folder_options_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub it_mycomp_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub it_mydocuments_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub it_network_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub it_panel_Change()
cmd_apply.Enabled = True
End Sub

Private Sub it_printers_Change()
cmd_apply.Enabled = True
End Sub

Private Sub it_scanner_Change()
cmd_apply.Enabled = True
End Sub

Private Sub it_schedule_Change()
cmd_apply.Enabled = True
End Sub

Private Sub it_smenu_Change()
cmd_apply.Enabled = True
End Sub



Private Sub menu_speed_Change()
 cmd_apply.Enabled = True
End Sub

Private Sub ms_ad_Click(index As Integer)
    cmd_apply.Enabled = True
End Sub

Private Sub ms_mnu_Click(index As Integer)
    cmd_apply.Enabled = True
End Sub

Private Sub pframe_Click()
 
frame_ms.Left = 290
frame_ms.Top = 690
 
frame_cu.Visible = False
frame_ms.Visible = False
frame_infotip.Visible = False
frame_security.Visible = False
 
Select Case pframe.SelectedItem.index
    Case 4
        frame_security.Visible = True
    Case 3
        frame_ms.Visible = True
    Case 2
        frame_infotip.Visible = True
    Case 1
        frame_cu.Visible = True
End Select


End Sub

Private Sub ms1_add_cmd_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_auto_logon_Click()
    If (ms1_auto_logon.Value = vbChecked) Then
        ms1_user_name.Enabled = True
        ms1_password.Enabled = True
    Else
        ms1_user_name.Enabled = False
        ms1_password.Enabled = False
    End If
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_clear_find_Click()
    Call clearfind
End Sub

Private Sub ms1_clear_ie_Click()
    Call clearie
End Sub

Private Sub ms1_clear_run_Click()
    Call clearrun
End Sub

Private Sub ms1_diable_beep_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_disable_cmd_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_enable_logon_Click()
    If (ms1_enable_logon.Value = 0) Then
        Form1.ms1_legal_caption.Enabled = False
        Form1.ms1_legal_text.Enabled = False
    Else
        Form1.ms1_legal_caption.Enabled = True
        Form1.ms1_legal_text.Enabled = True
    End If
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_install_location_Change()
   ' cmd_apply.Enabled = True
End Sub

Private Sub ms1_legal_caption_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_legal_text_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_min_animation_Click()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_password_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub ms1_user_name_Change()
    cmd_apply.Enabled = True
End Sub

Private Sub txt_cpass_KeyPress(KeyAscii As Integer)
    If ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90)) Then
        KeyAscii = KeyAscii
    Else
        If (KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txt_npass_Change()
If Trim(txt_npass.Text) <> "" Then
        chg_pass.Enabled = True
    Else
        chg_pass.Enabled = False
End If
    
End Sub

Private Sub txt_npass_KeyPress(KeyAscii As Integer)
    If ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90)) Then
        KeyAscii = KeyAscii
    Else
        If (KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
    Text1 = Len(txt_npass)
    
End Sub

Private Sub txt_owner_GotFocus()
    txt_owner.SelStart = 0
    txt_owner.SelLength = Len(txt_owner)
End Sub

Private Sub txt_owner_LostFocus()
    If (Len(Trim(txt_owner)) = 0) Then
        MsgBox "Name field cannot be empty", vbCritical, "Attention"
        txt_owner.SetFocus
    End If
End Sub
Private Sub txt_place_GotFocus()
    txt_place.SelStart = 0
    txt_place.SelLength = Len(txt_place)
End Sub

Private Sub txt_rpass_KeyPress(KeyAscii As Integer)
    If ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90)) Then
        KeyAscii = KeyAscii
    Else
        If (KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

