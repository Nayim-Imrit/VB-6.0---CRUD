VERSION 5.00
Begin VB.Form Rep 
   Caption         =   "Report Menu"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   Picture         =   "Rep.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdE 
      Caption         =   "Back To Main Menu"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Preview/Print List of Paints"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview/Print Sales Order"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Cmdfrmpaintpr 
      Caption         =   "Run Sales report"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Form to allow printing of reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdE_Click() 'program to exit form
Unload Me 'unload report menu form
mmenu.Show 'load menu form

End Sub

Private Sub Cmdfrmpaintpr_Click() 'program to run sales report'
Unload Me 'unload report form'
paintpr.Show 'load report on paint sold'
End Sub

Private Sub Command2_Click() 'Program to preview and print sales orders'
DataRepsalesorder.Refresh 'update Data report sales order'
DataRepsalesorder.Show ' load data data report sales order'
End Sub

Private Sub Command3_Click() 'program to view and print the list of paints'
DataReppaint.Refresh 'update data report paint'
DataReppaint.Show 'load data report paint'
End Sub


