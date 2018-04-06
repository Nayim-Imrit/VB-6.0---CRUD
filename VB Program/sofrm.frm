VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form sofrm 
   Caption         =   "Form to input Sales Order Details"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   Picture         =   "sofrm.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dtaso 
      Caption         =   "Sales Order"
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
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "So"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2775
   End
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
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Paint"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtprice 
      BackColor       =   &H00FFFF80&
      DataField       =   "price"
      DataSource      =   "Dtaso"
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
      Left            =   6600
      TabIndex        =   27
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtqty 
      BackColor       =   &H00FFFF80&
      DataField       =   "qty"
      DataSource      =   "Dtaso"
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
      Left            =   7560
      TabIndex        =   26
      Top             =   3120
      Width           =   855
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
      Left            =   8040
      TabIndex        =   25
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtpaintno 
      BackColor       =   &H00FFFF80&
      DataField       =   "paintno"
      DataSource      =   "Dtaso"
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
      Left            =   7680
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox cmbpaint 
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
      Left            =   1080
      TabIndex        =   23
      Top             =   3600
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker dtpdos 
      DataField       =   "dos"
      DataSource      =   "Dtaso"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   90308609
      CurrentDate     =   41636
   End
   Begin VB.TextBox txtitnu 
      BackColor       =   &H00FFFF80&
      DataField       =   "inu"
      DataSource      =   "Dtaso"
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
      Left            =   2040
      TabIndex        =   21
      Top             =   2520
      Width           =   975
   End
   Begin MSMask.MaskEdBox msoid 
      DataField       =   "soid"
      DataSource      =   "Dtaso"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
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
      Mask            =   "I####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Cmdexit 
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
      Left            =   4440
      TabIndex        =   17
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Cmdpack 
      Caption         =   "PACK all records"
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
      Left            =   5280
      TabIndex        =   16
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton Cmdmod 
      Caption         =   "Modify"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   6240
      Width           =   1095
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
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Cmddel 
      Caption         =   "Delete"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Cmdnext 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   735
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
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton CmdLast 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton CmdFirst 
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Navigation"
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   4560
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Data Processing"
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Label Label10 
      Caption         =   "Form to input Sales Order Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   29
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000B&
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
      Left            =   4080
      TabIndex        =   28
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Sold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Sales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales order ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "sofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Fillcmbpaint() 'program to fill combo box paint
DtaPaint.Recordset.MoveFirst 'move pointer to first record in table
While Not DtaPaint.Recordset.EOF 'while not end of file paint start loop
     mname = DtaPaint.Recordset.Fields!pname 'assign content from name in table to variant mname
     mdesc = DtaPaint.Recordset.Fields!Desc 'assign content from desc in table to variant mdesc
     
     mpaint = mname + " " + mdesc 'store the strings mname and mdesc in mdrug after joining them
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
If DtaPaint.Recordset.EOF = False Then 'if end of file paint not reached then
txtpaintno.Text = DtaPaint.Recordset.Fields!paintno 'copy content of field paintno from table to textbox in form
txtupv.Text = DtaPaint.Recordset.Fields!unitpv 'copy content of field unotpv from table to textbox in form
DtaPaint.Refresh 'open table paint for further processing

Else 'else if end of table reached
MsgBox ("No match found") 'message displayed no matching record
End If 'end if

End Sub

Private Sub CmdAdd_Click() 'program to add a new record in table sales order
Dim T As Integer 'declare T as integer
T = Dtaso.Recordset.RecordCount 'initialise T to the record count of table sales order
If T <> 0 Then 'if T is not zero, that is, table not empty
Dtaso.Recordset.MoveLast 'move to last record in table sales order
Dtaso.Recordset.MoveNext 'move to the next record after last record
End If 'end if

msoid.Mask = "I____" 'display I____ in MSOID
msoid.Mask = "I####" 'create mask to accept four numbers after I
msoid.Enabled = True 'enable mask edit box MSOID to enable input

Cmbpaint.Locked = False 'unlock combo box paint
dtpdos.Enabled = True 'enable expiry date and time picker
txtqty.Locked = False 'unlock quantity textbox
txtitnu.Locked = False 'unlock item number textbox
Cmbpaint = "" 'make combo box paint blank
txtitnu = "" 'make item number textbox blank
txtpaintno = "" 'make paint number textbox blank
txtqty = "" 'make quantity textbox blank
txtupv = "" 'make unit price&vat textbox blank
txtprice = "" 'make price textbox blank
msoid.SetFocus 'move cursor to MSOID mask edit box
CmdSave.Enabled = True 'enable command button save to save new record

End Sub

Private Sub Cmddel_Click() 'program to delete current record in table sales
B = MsgBox("Confirm Delete", vbQuestion + vbYesNo) 'provides dialog box with message to confirm delete and two buttons Yes/No which are stored in variable B as variant data type
    If B = vbYes Then 'if  the user selects yes then
        If Dtaso.Recordset.RecordCount = 0 Then 'if number of records in table=0 then
            MsgBox ("No Record to Delete") 'provides message no record to delete as table is empty
            Else
            Dtaso.Recordset.Delete 'delete current record in table
            MsgBox ("Record has been deleted") 'provides message to show that record has been deleted'
            Dtaso.Recordset.MoveFirst 'place pointer back on first record
        End If 'end if
    End If 'end if

End Sub

Private Sub cmdexit_Click() 'program to exit form
Unload Me 'unload sales order form
mmenu.Show 'load menu form

End Sub

Private Sub CmdFirst_Click() 'program to move pointer to first record in table sales order
Dtaso.Recordset.MoveFirst 'move to first record in table slaes order
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next
End Sub

Private Sub CmdLast_Click() 'program to move pointer to last record in table sales order
Dtaso.Recordset.MoveLast 'move pointer to last record in table
CmdFirst.Enabled = True 'enable command button first
cmdprevious.Enabled = True 'enable command button previous
CmdLast.Enabled = False 'disable command button last
Cmdnext.Enabled = False 'disable command button next

End Sub

Private Sub Cmdmod_Click() 'program to modify current record in table sales order
Cmbpaint.Locked = False 'unlock combo box paint
dtpdos.Enabled = True 'enable expiry date and time picker
txtqty.Locked = False 'unlock quantity textbox
txtitnu.Locked = False 'unlock item number textbox
txtitnu.SetFocus ',ove cursor to item number textbox
Dtaso.Recordset.Edit 'set edit mode of the table to allow modification
CmdSave.Enabled = True 'enable command button save
End Sub

Private Sub Cmdnext_Click() 'program to move pointer to next record from current record in table sales order
Dtaso.Recordset.MoveNext 'move pointer to next record in table
If Dtaso.Recordset.EOF Then 'if pointer is at end of table then
MsgBox ("End of file") 'message to inform that it is end of file
Cmdnext.Enabled = False 'disable command button next
CmdLast.Enabled = False 'disable command button last
End If 'end if
CmdFirst.Enabled = True 'enbable command button first
cmdprevious.Enabled = True 'enable command button

End Sub

Private Sub Cmdpack_Click() 'program to delete all records in table sales order

If Dtaso.Recordset.RecordCount <> 0 Then 'if table soo is not empty then
Dtaso.Recordset.MoveFirst 'move pointer to first record in table
Do While Dtaso.Recordset.EOF = False 'do instructions below till end of table not reached
Dtaso.Recordset.Delete 'delete current record
Dtaso.Recordset.MoveFirst 'move pointer back to first record in table
Loop 'repeat above instructions
CmdAdd.Enabled = True 'enable command button add
Cmddel.Enabled = True 'enable command button delete
Cmdmod.Enabled = True 'enable command button modify
Cmdpack.Enabled = False 'disable command button pack
Else 'else if table so is empty
CmdAdd.Enabled = True 'enable command button first
Cmddel.Enabled = True 'enable command button delete
Cmdmod.Enabled = True 'enable command button modify
Cmdpack.Enabled = False 'disable command button pack
End If 'end if
msoid.Mask = "I____" 'make mask edit box MSOID blank "I____"

End Sub

Private Sub cmdprevious_Click() 'program to move pointer to previous record from current record in table sales order
Dtaso.Recordset.MovePrevious 'move pointer to previous record in table
If Dtaso.Recordset.BOF Then 'if pointer is at begining of table then
MsgBox ("Begining of file") 'provide message to inform that it is begining of file
CmdFirst.Enabled = False 'disable command button first
cmdprevious.Enabled = False 'disable command button previous
End If 'end if
CmdLast.Enabled = True 'enable command button last
Cmdnext.Enabled = True 'enable command button next

End Sub

Private Sub CmdSave_Click() 'program to save a new record in table sales order
msoid.Enabled = False 'move to last record in table sales order
Cmbpaint.Locked = True 'lock combo box paint
dtpdos.Enabled = False 'disable expiry date and time picker
txtqty.Locked = True 'lock quantity textbox
txtitnu.Locked = True 'lock item number textbox
Dtaso.Recordset.Update 'move content from form to buffer to new record
Dtaso.Recordset.Requery 'move new record from buffer to table
MsgBox ("Record has been saved") 'provide message to confirm record has been saved
CmdSave.Enabled = False 'disable command button save so that record is not save more than once
End Sub

Private Sub Form_Load() 'program to fill combo box paint on form load event
Dtaso.DatabaseName = modatabasename
DtaPaint.DatabaseName = modatabasename
sql1 = "select * from Paint" 'store all records from table paint in variant sql1
DtaPaint.RecordSource = sql1 'sql1 is the source where records will be used to fill combo
DtaPaint.Refresh 'open paint table
If DtaPaint.Recordset.RecordCount > 0 Then 'if table paint is not empty then
    Call Fillcmbpaint 'call the procedure fillcombodrug
    Else 'else if table is empty
    MsgBox "Sorry file empty" 'provide message to inform that file is empty
End If 'end if

End Sub

Private Sub msoid_LostFocus() 'program to validate mask on last focus event in form sales order
Dim X As String 'declare X as string
X = msoid.Text 'copy content of mask in X
If X <> "I____" Then 'if X is not blank, then
Dtaso.Refresh 'open sales table
Dtaso.Recordset.FindFirst "SOID ='" + X + "'" 'find first record in table where SOID match X
     If Dtaso.Recordset.NoMatch = True Then 'if search not successful
         Dtaso.Recordset.AddNew 'add a new record to end table
         msoid.Text = X 'copy content of X in mask
         txtitnu.SetFocus 'move cursor to item number textbox
         Else 'else if there is a match
         Dtaso.Recordset.Close 'close the table
         MsgBox ("Sales Order ID already Exists!"), vbOKOnly + vbInformation 'provide message that sales order ID exists
         msoid.SetFocus 'block cursor on mask edit box
         msoid.Mask = "I####" ''recreate the mask for a new input
     End If 'end if
Else 'else if X is blank
msoid.SetFocus 'block cursor on mask edit box
MsgBox ("Please enter a Sales Order ID!"), vbOKOnly + vbInformation 'provide message to tell the user to enter Sales order ID
End If 'end if

End Sub



Private Sub txtqty_LostFocus() 'program to calculate the price'
txtprice = txtupv * txtqty     'multiply the unit price and vat by the quantity'

End Sub


