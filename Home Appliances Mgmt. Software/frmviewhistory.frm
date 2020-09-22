VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmviewhistory 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "frmmain.funcmdrecords"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmviewhistory.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adodcviewhistory 
      Height          =   375
      Left            =   4080
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "tbl_viewhistory"
      Caption         =   "Adodc View History"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgviewhistory 
      Bindings        =   "frmviewhistory.frx":32D68
      Height          =   4095
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "username"
      _Band(0)._MapCol(0)._Caption=   "User Name"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "date"
      _Band(0)._MapCol(1)._Caption=   "Date"
      _Band(0)._MapCol(1)._RSIndex=   3
      _Band(0)._MapCol(2)._Name=   "login"
      _Band(0)._MapCol(2)._Caption=   "Login Time"
      _Band(0)._MapCol(2)._RSIndex=   1
      _Band(0)._MapCol(3)._Name=   "logout"
      _Band(0)._MapCol(3)._Caption=   "Logout Time"
      _Band(0)._MapCol(3)._RSIndex=   2
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmviewhistory.frx":32D87
      MousePointer    =   99  'Custom
      Picture         =   "frmviewhistory.frx":33091
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmviewhistory.frx":33B5B
      MousePointer    =   99  'Custom
      Picture         =   "frmviewhistory.frx":33E65
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1350
      Left            =   1200
      Picture         =   "frmviewhistory.frx":348D9
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "View History"
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
      Height          =   435
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmviewhistory.frx":3638A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmviewhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmain.funcmdfile
End Sub

Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub

