VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form sales 
   BackColor       =   &H8000000A&
   Caption         =   "Form to input sales Details"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   Picture         =   "sales.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.Data DtaPaint 
      Caption         =   "Paint"
      Connect         =   "Access"
      DatabaseName    =   "D:\Project\Sofap paints\Psales.mdb"
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
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Paint"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DtaSal 
      Caption         =   "Sales"
      Connect         =   "Access"
      DatabaseName    =   "D:\Project\Sofap paints\Psales.mdb"
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
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Back to Main Menu Form"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Cmdsearch 
      Caption         =   "Search for Sales Details"
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
      Left            =   4560
      TabIndex        =   26
      Top             =   4920
      Width           =   3015
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
      Left            =   3600
      TabIndex        =   25
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Cmdmod 
      Caption         =   "Modify"
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
      TabIndex        =   24
      Top             =   4920
      Width           =   975
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
      Left            =   2640
      TabIndex        =   23
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Cmddel 
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
      Left            =   2400
      TabIndex        =   22
      Top             =   4920
      Width           =   975
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
      Left            =   1560
      TabIndex        =   21
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
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
      Left            =   1440
      TabIndex        =   20
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Cmdfirst 
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
      Left            =   360
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtprice 
      BackColor       =   &H00FFFF80&
      DataField       =   "price"
      DataSource      =   "DtaSal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtqty 
      BackColor       =   &H00FFFF80&
      DataField       =   "qty"
      DataSource      =   "DtaSal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtupv 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtpaintno 
      BackColor       =   &H00FFFF80&
      DataField       =   "paintno"
      DataSource      =   "DtaSal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox Cmbpaint 
      BackColor       =   &H00FFFF80&
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
      Left            =   1320
      TabIndex        =   13
      Top             =   3600
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker dtpdos 
      DataField       =   "dos"
      DataSource      =   "DtaSal"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
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
      CalendarForeColor=   16711680
      CalendarTitleBackColor=   255
      CalendarTitleForeColor=   16711680
      Format          =   90832897
      CurrentDate     =   41636
   End
   Begin VB.TextBox txtson 
      BackColor       =   &H00FFFF80&
      DataField       =   "son"
      DataSource      =   "DtaSal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   2295
   End
   Begin MSMask.MaskEdBox msid 
      DataField       =   "sid"
      DataSource      =   "DtaSal"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Mask            =   "S####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Navigation"
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   5760
      Width           =   7215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "Data Processing"
      Height          =   975
      Left            =   0
      TabIndex        =   29
      Top             =   4680
      Width           =   7815
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
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
      Left            =   4080
      TabIndex        =   9
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Form to input Sales Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "Price:"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Quantity:"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Unit Price and Vat:"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Paint Number:"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Paint:"
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
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Date Of Sales:"
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Sales Order Number:"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Sales ID:"
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
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Fillcmbpaint() 'program to fill combo box paint
DtaPaint.Recordset.MoveFirst 'move pointer to first record in table
While Not DtaPaint.Recordset.EOF 'while not end of file paint start loop
     mname = DtaPaint.Recordset.Fields!pname 'assign content from name in table to variant mname
     mdesc = DtaPaint.Recordset.Fields!Desc 'assign content from desc in table to variant mdesc
     
     mpaint = mname + " " + mdesc 'store the strings mname and mdesc in mpaint after joining them
     Cmbpaint.AddItem mpaint 'add the content of mpaint to combo box paint
     DtaPaint.Recordset.MoveNext 'move to next record in table to repeat adding item
 
Wend 'end the while loop at the end of table after filling the combo box with records
End Sub

Private Sub Cmbpaint_Click() 'program to select a record from the combo box on click event
selname = Cmbpaint.Text 'store the content of combo box in variant selname
splitname = Split(selname, " ") 'split selname into separate strings between spaces
Dim sql2 As String 'declare sql2 variant as string
sql2 = "select * from Paint where pname='" & splitname(0) & "'" 'search the table where name=splitpaint(0) and store it in sql2
DtaPaint.RecordSource = sql2 'make recordsource to be sql2
DtaPaint.Refresh 'open paint table
If DtaPaint.Recordset.EOF = False Then 'if end of file drug not reached then
txtpaintno.Text = DtaPaint.Recordset.Fields!paintno 'copy content of field paintno from table to textbox in form
txtupv.Text = DtaPaint.Recordset.Fields!unitpv 'copy content of field unotpv from table to textbox in form
DtaPaint.Refresh 'open table paint for further processing

Else 'else if end of table reached
MsgBox ("No match found") 'message displayed no matching record
End If 'end if

End Sub

Private Sub CmdAdd_Click() 'program to add a new record in table sales
Dim T As Integer 'declare T as integer
T = DtaSal.Recordset.RecordCount 'initialise T to the record count of table sales
If T <> 0 Then 'if T is not zero ,that is, table not empty
DtaSal.Recordset.MoveLast 'move to last record in table sales
DtaSal.Recordset.MoveNext 'move to the next record after last record
End If 'end if

msid.Enabled = True 'enable mask edit box to enable input
msid.Mask = "S____" 'display S____ in MSID
msid.Mask = "S####" 'create mask to accept four numbers after S
Cmbpaint.Locked = False 'unlock combo box paint
dtpdos.Enabled = True 'enable expiry date and time picker
txtqty.Locked = False 'unlock quantity textbox
txtson.Locked = False 'unlock sales order number textbox
Cmbpaint = "" 'make combo box paint blank
txtson = "" 'make sales order number textbox blank
txtqty = "" 'make quantity textbox blank
txtupv = "" 'make unit price&vat textbox blank
txtprice = "" 'make price textbox blank
txtpaintno = "" 'make paint number textbox blank
msid.SetFocus 'move cursor to MSID mask edit box
CmdSave.Enabled = True 'enable command button save to save new record


End Sub

Private Sub Cmddel_Click() 'program to delete current record in table sales
B = MsgBox("Confirm Delete", vbQuestion + vbYesNo) 'provides dialog box with message to confirm delete and two buttons Yes/No which are stored in variable B as variant data type
    If B = vbYes Then 'if  the user selects yes then
        If DtaSal.Recordset.RecordCount = 0 Then 'if number of records in table=0 then
            MsgBox ("No Record to Delete") 'provides message no record to delete as table is empty
            Else
            DtaSal.Recordset.Delete 'delete current record in table
            MsgBox ("Record has been deleted") 'provides message to show that record has been deleted'
            DtaSal.Recordset.MoveFirst 'place pointer back on first record
        End If 'end if
    End If 'end if

End Sub

Private Sub cmdexit_Click() 'program to exit form
Unload Me 'unload  sales form
mmenu.Show 'load menu form

End Sub

Private Sub CmdFirst_Click() 'program to move pointer to first record in table sales
DtaSal.Recordset.MoveFirst 'move to first record in table sales
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next

End Sub

Private Sub CmdLast_Click() 'program to move pointer to last record in table sales
DtaSal.Recordset.MoveLast 'move to first record in table sales
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next

End Sub

Private Sub Cmdmod_Click() 'program to modify current record in table sales
Cmbpaint.Locked = False 'unlock combo box paint
dtpdos.Enabled = True 'enable expiry date and time picker
txtqty.Locked = False 'unlock quantity textbox
txtson.Locked = False 'unlock sales order number textbox
txtson.SetFocus 'move cursor to sales order number textbox
DtaSal.Recordset.Edit 'set edit mode of the table to allow modification
CmdSave.Enabled = True 'enable command button save


End Sub

Private Sub Cmdnext_Click() 'program to move pointer to next record from current record in table sales
DtaSal.Recordset.MoveNext 'move pointer to next record in table
If DtaSal.Recordset.EOF Then 'if pointer is at end of table then
MsgBox ("End of file") 'provide message to inform that it is end of file
Cmdnext.Enabled = False 'disable command button next
CmdLast.Enabled = False 'disable command button last
End If 'end if
CmdFirst.Enabled = True 'enable command button first
cmdprevious.Enabled = True 'enable command button previous


End Sub

Private Sub cmdprev_Click() 'program to move pointer to previous record from current record in table sales
DtaSal.Recordset.MovePrevious 'move pointer to previous record in table
If DtaSal.Recordset.BOF Then 'if pointer is at begining of table then
MsgBox ("Begining of file") 'provide message to inform that it is begining of file
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
End If 'end if
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next
End Sub

Private Sub cmdprevious_Click()
DtaSal.Recordset.MovePrevious 'move pointer to previous record in table
If DtaSal.Recordset.BOF Then 'if pointer is at begining of table then
MsgBox ("Begining of file") 'provide message to inform that it is begining of file
CmdFirst.Enabled = False 'disable command button fisrt
cmdprevious.Enabled = False 'disable command button previous
End If 'end if
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next

End Sub

Private Sub CmdSave_Click() 'program to save a new record in table sales
msid.Enabled = False 'move to last record in table sales
Cmbpaint.Locked = True 'lock combo box paint
dtpdos.Enabled = False 'disable date of sales picker
txtqty.Locked = True 'lock quantity textbox
txtson.Locked = True 'lock sales order number textbox
DtaSal.Recordset.Update 'move content from form to buffer to new record
DtaSal.Recordset.Requery 'move new record from buffer to table
MsgBox ("Record has been saved") 'provide message to confirm record has been saved
CmdSave.Enabled = False 'disable command button save so that record is not save more than once causing duplication of file
End Sub

Private Sub Cmdsearch_Click() 'program to search for specific record in table sales
Dim searchvar As String 'declare variable searchvar as string
Dim sbookmark As String 'declare variable as sbookmark as string
searchvar = InputBox("Enter The Sales ID To Find") 'store the input box with the string
searchvar = Trim$(searchvar) ' removes surplus spaces
If searchvar <> "" Then 'if searchvar not empty then
With DtaSal.Recordset 'start loop using table sales to search
sbookmark = .Bookmark 'mark place in recordset to return to it later by storing it in the sbookmark variable
.FindFirst " sid like '" + searchvar + "*'" 'find the fisrt record where searchvar content=sid in table and display in paint main form
If .NoMatch Then 'if no match found then
MsgBox "No Matching Found" 'provide message box informing that no record was found
.Bookmark = sbookmark 'return pointer to initial position before starting search
End If 'end if
End With 'end loop for the search
End If 'end if

End Sub

Private Sub Form_Load() 'program to fill combo box paint on form load event
DtaPaint.DatabaseName = modatabasename
DtaSal.DatabaseName = modatabasename
sql1 = "select * from Paint" 'store all records from table paint in variant sql1
DtaPaint.RecordSource = sql1 'sql1 is the source where records will be used to fill combo
DtaPaint.Refresh 'open paint table
If DtaPaint.Recordset.RecordCount > 0 Then 'if table paint is not empty then
    Call Fillcmbpaint 'call the procedure fillcombopaint
    Else 'else if table is empty
    MsgBox "Sorry file empty" 'provide message to inform that file is empty
End If 'end if

End Sub



Private Sub msid_LostFocus() 'program to validate mask on lost focus event in form sales
Dim X As String 'declare X as string
X = msid.Text 'copy content of mask in X
If X <> "S____" Then 'if X is not blank, then
DtaSal.Refresh 'open sales table
DtaSal.Recordset.FindFirst "SID ='" + X + "'" 'find first record in table where sid match X
     If DtaSal.Recordset.NoMatch = True Then 'if search not successful
         DtaSal.Recordset.AddNew 'add a new record to end of table
         msid.Text = X 'copy content of X in mask
         txtson.SetFocus 'move cursor to sales order number textbox
         Else 'else if there is a match
         DtaSal.Recordset.Close 'close the table
         MsgBox ("Sales ID already Exists!"), vbOKOnly + vbInformation 'provide message that Sales ID exists
         msid.SetFocus 'block cursor on mask edit box
         msid.Mask = "S####" 'recreate the mask for a new input
     End If 'end if
Else 'else if X is blank
msid.SetFocus 'block cursor on mask edit box
MsgBox "Please enter a Sales ID!", vbOKOnly + vbInformation 'provide message to enter sales Id
End If 'end if
End Sub





Private Sub txtqty_LostFocus() 'program to calculate the price'
txtprice = txtupv * txtqty     'multiply the unit price and vat by the quantity'

End Sub
