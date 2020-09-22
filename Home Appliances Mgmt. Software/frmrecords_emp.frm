VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmrecords_emp 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Employee Records"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrecords_emp.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmrecords_emp.frx":32D68
      Height          =   3855
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   10
      _Band(0)._MapCol(0)._Name=   "empid"
      _Band(0)._MapCol(0)._Caption=   "Id"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "empname"
      _Band(0)._MapCol(1)._Caption=   "Name"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "empgender"
      _Band(0)._MapCol(2)._Caption=   "Gender"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "empdob"
      _Band(0)._MapCol(3)._Caption=   "Date of Birth"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "empcontactno"
      _Band(0)._MapCol(4)._Caption=   "Contact No."
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Alignment=   7
      _Band(0)._MapCol(5)._Name=   "empaddress"
      _Band(0)._MapCol(5)._Caption=   "Address"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "empcity"
      _Band(0)._MapCol(6)._Caption=   "City"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "empdoj"
      _Band(0)._MapCol(7)._Caption=   "Date of Joining"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "emppostion"
      _Band(0)._MapCol(8)._Caption=   "Position"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "empsalary"
      _Band(0)._MapCol(9)._Caption=   "Salary"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(9)._Alignment=   7
   End
   Begin MSAdodcLib.Adodc adodcrecords_emp 
      Height          =   330
      Left            =   5040
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
   Begin Candy.CandyButton cmdnew 
      Height          =   375
      Left            =   4440
      TabIndex        =   0
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
      Caption         =   "New"
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
   Begin Candy.CandyButton cmdedit 
      Height          =   375
      Left            =   5520
      TabIndex        =   1
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
      Caption         =   "Edit"
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
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmrecords_emp.frx":32D87
      MousePointer    =   99  'Custom
      Picture         =   "frmrecords_emp.frx":33091
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmrecords_emp.frx":33B05
      MousePointer    =   99  'Custom
      Picture         =   "frmrecords_emp.frx":33E0F
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   1080
      Picture         =   "frmrecords_emp.frx":348D9
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      TabIndex        =   2
      Top             =   1320
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmrecords_emp.frx":35C4F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmrecords_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmain.funcmdrecords
End Sub

Private Sub cmdedit_Click()
emprecclick = "Edit"
UnloadAllForms ("mdimain")
frmrecords_emp_new_edit.Left = 0
frmrecords_emp_new_edit.Top = 0
End Sub





Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub



Private Sub cmdnew_Click()
emprecclick = "New"
UnloadAllForms ("mdimain")
frmrecords_emp_new_edit.Left = 0
frmrecords_emp_new_edit.Top = 0
End Sub


Private Sub Form_Load()
If adodcrecords_emp.Recordset.EOF = True Then
    cmdedit.Enabled = False
    Exit Sub
Else
cmdedit.Enabled = True
End If
End Sub
