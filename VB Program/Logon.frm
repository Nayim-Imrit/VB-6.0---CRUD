VERSION 5.00
Begin VB.Form Logon 
   Caption         =   "Login Form"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   Picture         =   "Logon.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtpw 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtuid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Security form to access the system by entering user ID and  password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   12135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sofap Paint LTD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcancel_Click() 'program to exit logon form
End 'exit login form

End Sub

Private Sub CmdOk_Click() 'program to validate password and user ID
If txtpw = "Sofap1995" And txtuid = "Anil" Then 'if password is Sofap1995 and user id is Anil then
    Unload Me 'unload the current logon form
     modatabasename = App.Path & "\Psales.mdb"
   frmSplash.Show 'load the splash form
Else 'else if password and user ID wrong
   MsgBox ("Access Denied ,Please Try again") 'display message "Access Denied, Try again"
End If 'end if

End Sub






