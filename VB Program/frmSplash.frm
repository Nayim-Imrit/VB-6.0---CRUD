VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1800
      Top             =   3840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address: Sofap_paint@gmail.com "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax: (+230) 4321234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel: 4225521"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Royal-Road, Port-louis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sofap Paint LTD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
frmSplash.Enabled = False
Unload Me 'unload form splash
mmenu.Show 'load menu form
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False '
frmSplash.Enabled = False
Unload Me 'unload form splash
mmenu.Show 'load menu form
End Sub
