VERSION 5.00
Begin VB.Form paintpr 
   Caption         =   "Form to run report on Paint Sold"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   Picture         =   "paintpr.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpaintno 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data DSal 
      Caption         =   "DataSalesQ"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Amal\Documents\HP61 Docs\Projects 2014\Permoglaze paints\Psales.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Qsal"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Cmdexit 
      Caption         =   "Back to Report Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Cmdpr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox Cmbpaint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Form to run report on the most selling paint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   5655
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
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "paintpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fillcmbpaint()
DSal.Recordset.MoveFirst
While Not DSal.Recordset.EOF
     mpaintno = DSal.Recordset.Fields!paintno
     msid = DSal.Recordset.Fields!sid
     mson = DSal.Recordset.Fields!son
     
     mpaint = mpaintno + " " + msid + " " + mson
     Cmbpaint.AddItem mpaint
     DSal.Recordset.MoveNext
 
Wend
End Sub

Private Sub Cmbpaint_Click()
selname = Cmbpaint.Text
splitname = Split(selname, " ")
Dim sqld As String
sqld = "select * from Qsal where paintno='" & splitname(0) & "'"
DSal.RecordSource = sqld
DSal.Refresh
If DSal.Recordset.EOF = False Then
txtpaintno.Text = DSal.Recordset.Fields!paintno

DSal.Refresh
Else
MsgBox ("Sorry")
End If



End Sub

Private Sub cmdexit_Click()
Unload Me
Rep.Show
End Sub

Private Sub Cmdpr_Click()
Dim Datarep As New DataEnvironment1
Datarep.sales (txtpaintno.Text)
DataRepsales.Refresh
DataRepsales.Show
End Sub

Private Sub Form_Load()
DSal.DatabaseName = modatabasename
sql6 = "select * from Qsal"
DSal.RecordSource = sql6
DSal.Refresh
If DSal.Recordset.RecordCount > 0 Then
    Call Fillcmbpaint
    Else
    MsgBox "Sorry"
    End If
End Sub


