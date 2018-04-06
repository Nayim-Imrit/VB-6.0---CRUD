VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form paintform 
   Caption         =   "Paint form"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   Picture         =   "paint.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dtapaint 
      Caption         =   "Paint"
      Connect         =   "Access"
      DatabaseName    =   "D:\Project\Sofap paints\Psales.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Paint"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search for specific paint details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   25
      Top             =   6600
      Width           =   3135
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Back to main menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   24
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdmod 
      Caption         =   "modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   23
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   5400
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtped 
      DataField       =   "expiry"
      DataSource      =   "Dtapaint"
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777088
      CalendarTrailingForeColor=   12632256
      Format          =   76808193
      CurrentDate     =   41635
   End
   Begin MSMask.MaskEdBox mpid 
      DataField       =   "paintno"
      DataSource      =   "Dtapaint"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   16776960
      Enabled         =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "P####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtupv 
      BackColor       =   &H00FFFF80&
      DataField       =   "unitpv"
      DataSource      =   "Dtapaint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9600
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtup 
      BackColor       =   &H00FFFF80&
      DataField       =   "unitp"
      DataSource      =   "Dtapaint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txttype 
      BackColor       =   &H00FFFF80&
      DataField       =   "ptype"
      DataSource      =   "Dtapaint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3960
      Width           =   4575
   End
   Begin VB.TextBox txtdesc 
      BackColor       =   &H00FFFF80&
      DataField       =   "desc"
      DataSource      =   "Dtapaint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFF80&
      DataField       =   "pname"
      DataSource      =   "Dtapaint"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Navigation"
      Height          =   975
      Left            =   360
      TabIndex        =   26
      Top             =   5040
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Data processing"
      Height          =   1095
      Left            =   360
      TabIndex        =   27
      Top             =   6240
      Width           =   7695
   End
   Begin VB.Label Label9 
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
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Form to input paint details"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price & Vat:"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price:"
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
      Left            =   7560
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
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
      Left            =   7560
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Paint Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Paint Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Paint Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paint ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "paintform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdAdd_Click() 'program to add a new record in table paint
Dim T As Integer 'declare T as integer
T = DtaPaint.Recordset.RecordCount 'initialise T to the record count of table paint
If T <> 0 Then 'if T is not zero,that is,table not empty
DtaPaint.Recordset.MoveLast 'move to last record in table paint
DtaPaint.Recordset.MoveNext 'move to the next record after last record
End If 'end if
mpid.Enabled = True 'enable mask edit box Mpid to enable input
mpid.Mask = "P____" 'display P____ in mpid
mpid.Mask = "P####" 'create mask to accept four numbers after P
txtname.Locked = False 'unlock paint name textbox
txtdesc.Locked = False 'unlock description textbox
txttype.Locked = False 'unlock type textbox
dtped.Enabled = True 'enable expiry date and time picker
txtup.Locked = False 'unlock unit price textbox
txtname = "" 'make paint name textbox blank
txtdesc = "" 'make description textbox blank
txttype = "" 'make type textbox blank
txtup = "" 'make unit price textbox blank

mpid.SetFocus 'move cursor to mpid mask edit box
CmdSave.Enabled = True 'enable command button save to save new record
End Sub

Private Sub Cmddel_Click() 'program to delete current record in table paint'
Dim D As Integer  'declare variable D'
D = MsgBox("Confirm Delete", vbQuestion + vbYesNo) 'provides dialog box with message to confirm delete and two buttons Yes/No which are stored in variable D as variant data type'
    If D = vbYes Then 'if the user selects yes then'
        If DtaPaint.Recordset.RecordCount = 0 Then 'if number of records in table=0'
            MsgBox ("No Record to Delete") 'provides message no record to delete as table is empty'
            Else
            DtaPaint.Recordset.Delete 'delete current record in table'
            MsgBox ("Record has been deleted") 'provides message to show that record has been deleted'
            DtaPaint.Recordset.MoveFirst 'place pointer back on first record'
        End If 'end if'
    End If 'end if'

End Sub

Private Sub cmdexit_Click() 'program to exit form
Unload Me 'unload drug form
mmenu.Show 'load menu form
End Sub

Private Sub CmdFirst_Click() 'program to move pointer to first record in table paint
DtaPaint.Recordset.MoveFirst 'move to first record in table Paint
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next
End Sub

Private Sub CmdLast_Click() 'program to move pointer to last record in table Paint
DtaPaint.Recordset.MoveLast 'move pointer to last record in table
CmdFirst.Enabled = True 'enable command button first
cmdprevious.Enabled = True 'enable command button previous
CmdLast.Enabled = False 'disable command button last
Cmdnext.Enabled = False 'disable command button next

End Sub

Private Sub Cmdmod_Click()
DtaPaint.Recordset.Edit 'Allow edit mode to Paint Table
txtname.SetFocus
CmdSave.Enabled = True 'Enable the Command Button Save to allow change to the record
txtname.Locked = False 'Unlock textbox txtpaintno to allow input of Paint number
txtdesc.Locked = False 'unlock description textbox
txttype.Locked = False 'unlock paint type textbox
txtup.Locked = False 'unlock unit price textbox
dtped.Enabled = True 'enable expiry date and time picker
End Sub

Private Sub Cmdnext_Click() 'program to move pointer to next record from current record in table Paint
DtaPaint.Recordset.MoveNext 'move pointer to next record in table
If DtaPaint.Recordset.EOF Then 'if pointer is at end of table then
MsgBox ("End of file") 'message to inform that it is end of file
Cmdnext.Enabled = False 'disable command button next
CmdLast.Enabled = False 'disable command button last
End If 'end if
CmdFirst.Enabled = True 'enable command button first
cmdprevious.Enabled = True 'enable command button previous
End Sub

Private Sub cmdprevious_Click() 'program to move pointer to previous record from current record in table Paint
DtaPaint.Recordset.MovePrevious 'move pointer to previous record in table
If DtaPaint.Recordset.BOF Then 'if pointer is at begining of table then
MsgBox ("Begining of file") 'provide message to inform that it is begining of file
CmdFirst.Enabled = False 'disable command button fisrt
cmdprevious.Enabled = False 'disable command button previous
End If 'end if
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next

End Sub

Private Sub CmdSave_Click()
mpid.Enabled = False 'lock textbox txtpaintno to prevent further input of Paint number
txtname.Locked = True
txtdesc.Locked = True
dtped.Enabled = False
txtup.Locked = True
txttype.Locked = True
DtaPaint.Recordset.Update 'Copy contents of c opy buffer back to table
DtaPaint.Recordset.Requery 'Updates the data in table
MsgBox ("Record has been saved") 'provide message to confirm record has been saved
CmdSave.Enabled = False 'Disable the Command Button Save to prevent saving of the record more than once
End Sub

Private Sub Cmdsearch_Click() 'program to search for specific record in table paint
Dim searchvar As String 'declare variable searchvar as string
Dim sbookmark As String 'declare variable sbookmark as string
searchvar = InputBox("Enter The Paint Number To Find") 'store the input box with the string
searchvar = Trim$(searchvar) ' removes surplus spaces
If searchvar <> "" Then 'if searchvar not empty then
With DtaPaint.Recordset 'start loop using table paint to search
sbookmark = .Bookmark 'mark place in reordset to return to it later by storing it in the sbookmark variable
.FindFirst " paintno like '" + searchvar + "*'" 'find the first record where searchvar content=name in table and display in paint main form
If .NoMatch Then 'if no match found then
MsgBox "No Matching Found" 'provide message box informoing that no record was found
.Bookmark = sbookmark 'return pointer to initial position before starting search
End If 'end if
End With 'end loop for the search
End If 'end if

End Sub

Private Sub Form_Load()
DtaPaint.DatabaseName = modatabasename
End Sub

Private Sub mpid_LostFocus() 'program to validate mask on lost focus event in form paint
Dim Y As String 'declare Y as string
Y = mpid.Text 'copy content of mask in Y
mpid.Text = "P____" 'make mask blank
If Y <> "P____" Then 'if Y is not blank
DtaPaint.Refresh 'move paint table to memory i.e open table
DtaPaint.Recordset.FindFirst "paintno = '" + Y + "'" 'find first record in table where paintno match Y
If DtaPaint.Recordset.NoMatch = True Then 'if search not successful
DtaPaint.Recordset.AddNew 'add a new record to end of table
mpid.Text = Y 'copy content of Y in mask
txtname.SetFocus 'move cursor to drug name textbox
Else
DtaPaint.Recordset.Close 'close the table
MsgBox ("Paint Number already exist") 'provide message that paint number already exist
mpid.Mask = "P####" 'recreate the mask for a new input
End If 'end if
Else
mpid.SetFocus 'block cursor i=on mask edit box
MsgBox ("Please enter a Paint Number!"), vbOKOnly + vbInformation 'provide message signaling to enter paint number
End If 'end if
End Sub



Private Sub txtup_LostFocus() 'program to calculate the unit price and vat'
txtupv = 1.15 * txtup   'add vat to the unit price'

End Sub



