VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2625
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6150
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1550.937
   ScaleMode       =   0  'User
   ScaleWidth      =   5774.517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   405
      Left            =   3750
      TabIndex        =   4
      Top             =   2010
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2422
      TabIndex        =   3
      Top             =   2010
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1095
      TabIndex        =   2
      Top             =   2010
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1290
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "This Program is protected by the password. Please provide the correct password to log on to the program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   375
      TabIndex        =   1
      Top             =   390
      Width           =   5535
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private string_got As String * 255
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    End
End Sub

Private Sub cmdOK_Click()
    Dim cpass As String
    Dim temp_read As String
    Dim slength As Long
    Dim to_check As String
    
    temp_read = Space(255)
    slength = GetProfileString("PhotoOptics", "sclassname", "N", temp_read, 255)
    temp_read = Left(temp_read, slength)
    If (temp_read = "N") Then
        cpass = "PleaseLogin"
    Else
        to_check = decrypt(temp_read)
        cpass = to_check
    End If
    
    If txtPassword = cpass Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        Form1.Show
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", vbInformation, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command1_Click()
frmAbout.Show 1
End Sub
