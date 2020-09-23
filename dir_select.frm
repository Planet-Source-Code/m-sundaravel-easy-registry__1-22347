VERSION 5.00
Begin VB.Form dir_select 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory Select"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   1928
      TabIndex        =   3
      Top             =   3915
      Width           =   1155
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   390
      Left            =   578
      TabIndex        =   2
      Top             =   3900
      Width           =   1155
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   75
      TabIndex        =   1
      Top             =   945
      Width           =   3480
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   510
      Width           =   3510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select the directory you want:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   165
      Width           =   2565
   End
End
Attribute VB_Name = "dir_select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_ok_Click()
    Form1.ms1_install_location = Dir1.Path
    Unload Me
    Form1.cmd_apply.Enabled = True
End Sub

Private Sub Form_Load()
    On Error GoTo skip
        Dir1.Path = Form1.ms1_install_location
skip:
End Sub
