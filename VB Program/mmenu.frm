VERSION 5.00
Begin VB.Form mmenu 
   Caption         =   "Main menu form"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   Picture         =   "mmenu.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdemm 
      Caption         =   "Exit Main Menu Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   6720
      Width           =   3375
   End
   Begin VB.CommandButton Cmdrf 
      Caption         =   "Reports Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton Cmdsof 
      Caption         =   "Sales Order Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Cmdsf 
      Caption         =   "Sales Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Cmdpf 
      Caption         =   "Paint Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Main menu form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sofap Paint LTD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "mmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdemm_Click() 'program to exit main menu form'
Unload Me 'unload main menu form'
logoutfrm.Show 'load log out form'
End Sub

Private Sub Cmdpf_Click() 'program to enter paint form'
Unload Me 'unload main menu form'
paintform.Show 'load paint form'
End Sub

Private Sub Cmdrf_Click() 'program to enter report form'
Unload Me 'unload main menu form
Rep.Show 'load report form'
End Sub

Private Sub Cmdsf_Click() 'program to enter sales form'
Unload Me 'unload main menu form'
sales.Show 'load Sales form'
End Sub

Private Sub Cmdsof_Click() 'program to enter sales order form'
Unload Me 'unload main menu form'
sofrm.Show 'load Sales Order form'
End Sub


