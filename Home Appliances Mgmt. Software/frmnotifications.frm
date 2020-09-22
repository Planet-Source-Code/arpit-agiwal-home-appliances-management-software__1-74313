VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmnotifications 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Notifications"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnotifications.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adodcnotifications 
      Height          =   330
      Left            =   2520
      Top             =   600
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
      RecordSource    =   "qry_notifications"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgnotifications 
      Bindings        =   "frmnotifications.frx":32D68
      Height          =   4455
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7858
      _Version        =   393216
      BackColorBkg    =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmnotifications.frx":32D89
      MousePointer    =   99  'Custom
      Picture         =   "frmnotifications.frx":33093
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmnotifications.frx":33B07
      MousePointer    =   99  'Custom
      Picture         =   "frmnotifications.frx":33E11
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Notifications"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmnotifications.frx":348DB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmnotifications"
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



