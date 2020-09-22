VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmrecords_sup 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Supplier Records"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrecords_sup.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmrecords_sup.frx":32D68
      Height          =   3855
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   7
      _Band(0)._MapCol(0)._Name=   "supid"
      _Band(0)._MapCol(0)._Caption=   "Id"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "supname"
      _Band(0)._MapCol(1)._Caption=   "Name"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "supcontactno"
      _Band(0)._MapCol(2)._Caption=   "Contact No."
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "supemailid"
      _Band(0)._MapCol(3)._Caption=   "Email Id"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "supaddress"
      _Band(0)._MapCol(4)._Caption=   "Address"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "supcity"
      _Band(0)._MapCol(5)._Caption=   "City"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "supstate"
      _Band(0)._MapCol(6)._Caption=   "State"
      _Band(0)._MapCol(6)._RSIndex=   6
   End
   Begin MSAdodcLib.Adodc adodcrecords_sup 
      Height          =   330
      Left            =   4920
      Top             =   120
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
      RecordSource    =   "tbl_sup_record"
      Caption         =   ""
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
      TabIndex        =   1
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
      TabIndex        =   2
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
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
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   960
      Picture         =   "frmrecords_sup.frx":32D87
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Records"
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
      Left            =   4110
      TabIndex        =   3
      Top             =   1320
      Width           =   3450
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmrecords_sup.frx":343CA
      MousePointer    =   99  'Custom
      Picture         =   "frmrecords_sup.frx":346D4
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmrecords_sup.frx":3519E
      MousePointer    =   99  'Custom
      Picture         =   "frmrecords_sup.frx":354A8
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmrecords_sup.frx":35F1C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmrecords_sup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdback_Click()
frmmain.funcmdrecords
End Sub

Private Sub cmdedit_Click()
suprecclick = "Edit"
UnloadAllForms ("mdimain")
frmrecords_sup_new_edit.Left = 0
frmrecords_sup_new_edit.Top = 0
End Sub





Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub



Private Sub cmdnew_Click()
suprecclick = "New"
UnloadAllForms ("mdimain")
frmrecords_sup_new_edit.Left = 0
frmrecords_sup_new_edit.Top = 0
End Sub






Private Sub Form_Load()
If adodcrecords_sup.Recordset.EOF = True Then
    cmdedit.Enabled = False
    Exit Sub
Else
cmdedit.Enabled = True
End If
End Sub
