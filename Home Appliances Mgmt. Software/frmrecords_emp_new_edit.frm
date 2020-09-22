VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmrecords_emp_new_edit 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "New/Edit Employee Records "
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrecords_emp_new_edit.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton cmdprevious 
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLite    =   -1  'True
      IconHighLiteColor=   16777215
      CaptionHighLiteColor=   0
      Picture         =   "frmrecords_emp_new_edit.frx":32D68
      PictureAlignment=   2
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin VB.TextBox txtgotoid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin Candy.CandyButton cmdgo 
      Height          =   390
      Left            =   3840
      TabIndex        =   27
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Go to Id"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      ForeColor       =   16777215
      PictureAlignment=   2
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin MSComCtl2.DTPicker DTPickerdob 
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16761024
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   16515073
      CurrentDate     =   40948
   End
   Begin VB.ComboBox cmbgender 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmrecords_emp_new_edit.frx":331BA
      Left            =   2880
      List            =   "frmrecords_emp_new_edit.frx":331C4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtsalary 
      DataField       =   "empsalary"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      MaxLength       =   8
      TabIndex        =   10
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtposition 
      DataField       =   "emppostion"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      MaxLength       =   100
      TabIndex        =   8
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtcity 
      DataField       =   "empcity"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      MaxLength       =   100
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "empaddress"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtcontactno 
      DataField       =   "empcontactno"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtname 
      DataField       =   "empname"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      MaxLength       =   100
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtid 
      DataField       =   "empid"
      DataSource      =   "adodcrecords_emp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc adodcrecords_emp 
      Height          =   330
      Left            =   4920
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbl_emp_record"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Candy.CandyButton cmdsave 
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Save"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   65535
      ForeColor       =   16777215
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin Candy.CandyButton cmddelete 
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   65535
      ForeColor       =   16777215
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin Candy.CandyButton cmdcancel 
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   65535
      ForeColor       =   16777215
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin MSComCtl2.DTPicker DTPickerdoj 
      Height          =   390
      Left            =   2880
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16761024
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   16515073
      CurrentDate     =   40948
   End
   Begin Candy.CandyButton cmdnext 
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ">"
      IconHighLite    =   -1  'True
      IconHighLiteColor=   16777215
      CaptionHighLiteColor=   0
      Picture         =   "frmrecords_emp_new_edit.frx":331D6
      PictureAlignment=   3
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   26
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date of Joining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   1080
      Picture         =   "frmrecords_emp_new_edit.frx":33628
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblheader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3960
      TabIndex        =   16
      Top             =   1320
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmrecords_emp_new_edit.frx":3499E
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmrecords_emp_new_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CandyButton1_Click()

End Sub

Private Sub cmdcancel_Click()
Dim a As Integer
a = MsgBox("Are you sure to cancel ?", vbYesNo, "AADIK Technologies")
If a = vbYes Then
adodcrecords_emp.Recordset.CancelUpdate
mdimain.mnuemprec_Click
End If
End Sub

Private Sub cmddelete_Click()
Dim a As Integer
a = MsgBox("Are you sure to delete this record", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    adodcrecords_emp.Recordset.Delete
    adodcrecords_emp.Recordset.MoveFirst
    MsgBox "Your record have been successfully deleted", vbOKOnly, "AADIK Technologies"
    End If
If adodcrecords_pro.Recordset.BOF = True Then
mdimain.mnuemprec_Click
End If
End Sub

Private Sub cmdgo_Click()
adodcrecords_emp.Recordset.MoveFirst
adodcrecords_emp.Recordset.Find "empid='" & txtgotoid.Text & "'"
    If adodcrecords_emp.Recordset.EOF = True Then
    MsgBox "No Record Found", vbOKOnly, "AADIK Technologies"
    adodcrecords_emp.Recordset.MoveFirst
    Exit Sub
    Else
    cmbgender.Text = adodcrecords_emp.Recordset.Fields("empgender")
    DTPickerdob.Value = adodcrecords_emp.Recordset.Fields("empdob")
    DTPickerdoj.Value = adodcrecords_emp.Recordset.Fields("empdoj")
    End If
End Sub





Private Sub cmdnext_Click()

adodcrecords_emp.Recordset.MoveNext
    If adodcrecords_emp.Recordset.EOF = True Then
    adodcrecords_emp.Recordset.MoveLast
    End If
cmbgender.Text = adodcrecords_emp.Recordset.Fields("empgender")
DTPickerdob.Value = adodcrecords_emp.Recordset.Fields("empdob")
DTPickerdoj.Value = adodcrecords_emp.Recordset.Fields("empdoj")




End Sub

Private Sub cmdprevious_Click()
adodcrecords_emp.Recordset.MovePrevious
    If adodcrecords_emp.Recordset.BOF = True Then
    adodcrecords_emp.Recordset.MoveFirst
    End If
cmbgender.Text = adodcrecords_emp.Recordset.Fields("empgender")
DTPickerdob.Value = adodcrecords_emp.Recordset.Fields("empdob")
DTPickerdoj.Value = adodcrecords_emp.Recordset.Fields("empdoj")

End Sub

Private Sub cmdsave_Click()
If txtid.Text = "" Or txtname = "" Or cmbgender.Text = "" Or DTPickerdob.Value = "" Or DTPickerdoj.Value = "" Or txtcontactno.Text = "" Or txtaddress.Text = "" Or txtcity.Text = "" Or txtposition.Text = "" Or txtsalary.Text = "" Then
MsgBox "Please fill all fields", vbOKOnly, "AADIK Technologies"
Exit Sub
Else
Dim a As Integer
adodcrecords_emp.Recordset.Fields("empgender") = cmbgender
adodcrecords_emp.Recordset.Fields("empdob") = DTPickerdob.Value
adodcrecords_emp.Recordset.Fields("empdoj") = DTPickerdoj.Value
adodcrecords_emp.Recordset.Save
MsgBox "Your record have been successfully saved", vbOKOnly, "AADIK Technologies"
a = MsgBox("Are you want to Adding records continue ?", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    lblheader.Caption = "Add Employee Records"
    adodcrecords_emp.Recordset.AddNew
    Else
    mdimain.mnuemprec_Click
    End If
    
End If
End Sub



Private Sub Form_Load()
DTPickerdob.Value = Now()
DTPickerdoj.Value = Now()
DTPickerdoj.MaxDate = Now()
If emprecclick = "New" Then
lblheader.Caption = "Add Employee Records"
cmddelete.Enabled = False
txtgotoid.Visible = False
cmdgo.Visible = False
cmdprevious.Visible = False
cmdnext.Visible = False
adodcrecords_emp.Recordset.AddNew
DTPickerdoj.Value = Now()

ElseIf emprecclick = "Edit" Then
lblheader.Caption = "Edit Employee Records"
cmddelete.Enabled = True
txtgotoid.Visible = True
cmdgo.Visible = True
cmdprevious.Visible = True
cmdnext.Visible = True
cmbgender.Text = adodcrecords_emp.Recordset.Fields("empgender")
DTPickerdob.Value = adodcrecords_emp.Recordset.Fields("empdob")
DTPickerdoj.Value = adodcrecords_emp.Recordset.Fields("empdoj")
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
Call onlyaddress(KeyAscii)
End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
Call onlyalpha(KeyAscii)
End Sub

Private Sub txtcontactno_KeyPress(KeyAscii As Integer)
Call onlynumeric(KeyAscii)
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
Call onlyalphanum(KeyAscii)
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Call onlyalpha(KeyAscii)
End Sub

Private Sub txtposition_KeyPress(KeyAscii As Integer)
Call onlyalpha(KeyAscii)
End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
Call onlynumeric(KeyAscii)
End Sub
