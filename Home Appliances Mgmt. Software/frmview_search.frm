VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmview_search 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Search"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmview_search.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtsearch 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   34
      Top             =   2040
      Width           =   9135
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   5535
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Employees  "
      TabPicture(0)   =   "frmview_search.frx":32D68
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "adodcemployees"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "hfgemployees"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Suppliers   "
      TabPicture(1)   =   "frmview_search.frx":32D84
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "adodcsuppliers"
      Tab(1).Control(1)=   "hfgsuppliers"
      Tab(1).Control(2)=   "Option5"
      Tab(1).Control(3)=   "Option6"
      Tab(1).Control(4)=   "Option7"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Customers   "
      TabPicture(2)   =   "frmview_search.frx":32DA0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "adodccustomers"
      Tab(2).Control(1)=   "hfgcustomers"
      Tab(2).Control(2)=   "Option8"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Products    "
      TabPicture(3)   =   "frmview_search.frx":32DBC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "adodcproducts"
      Tab(3).Control(1)=   "hfgproducts"
      Tab(3).Control(2)=   "Option9"
      Tab(3).Control(3)=   "Option10"
      Tab(3).Control(4)=   "Option11"
      Tab(3).Control(5)=   "Option12"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Stocks     "
      TabPicture(4)   =   "frmview_search.frx":32DD8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "adodcstocks"
      Tab(4).Control(1)=   "hfgstocks"
      Tab(4).Control(2)=   "Option13"
      Tab(4).Control(3)=   "Option14"
      Tab(4).Control(4)=   "Option15"
      Tab(4).Control(5)=   "Option16"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Sales      "
      TabPicture(5)   =   "frmview_search.frx":32DF4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "adodcsales"
      Tab(5).Control(1)=   "hfgsales"
      Tab(5).Control(2)=   "Option17"
      Tab(5).Control(3)=   "Option18"
      Tab(5).Control(4)=   "Option19"
      Tab(5).Control(5)=   "Option20"
      Tab(5).Control(6)=   "Option21"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "Purchases      "
      TabPicture(6)   =   "frmview_search.frx":32E10
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "adodcpurchases"
      Tab(6).Control(1)=   "hfgpurchases"
      Tab(6).Control(2)=   "Option22"
      Tab(6).Control(3)=   "Option23"
      Tab(6).Control(4)=   "Option24"
      Tab(6).Control(5)=   "Option25"
      Tab(6).Control(6)=   "Option26"
      Tab(6).ControlCount=   7
      Begin VB.OptionButton Option16 
         Caption         =   "Product Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68400
         TabIndex        =   15
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Product Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70605
         TabIndex        =   33
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option26 
         Caption         =   "Purchases Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68160
         TabIndex        =   25
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option25 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69930
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option24 
         Caption         =   "Product Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71460
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option23 
         Caption         =   "Supplier Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73470
         TabIndex        =   22
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option22 
         Caption         =   "Bill No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option21 
         Caption         =   "Sales Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67680
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option20 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69600
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Product Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71280
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Bill No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72795
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Product Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69480
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70680
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72960
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69960
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71280
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72240
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5730
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4605
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3375
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2490
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgemployees 
         Bindings        =   "frmview_search.frx":32E2C
         Height          =   3855
         Left            =   0
         TabIndex        =   26
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
      Begin MSAdodcLib.Adodc adodcemployees 
         Height          =   330
         Left            =   5760
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsuppliers 
         Bindings        =   "frmview_search.frx":32E49
         Height          =   3855
         Left            =   -75000
         TabIndex        =   27
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
      Begin MSAdodcLib.Adodc adodcsuppliers 
         Height          =   330
         Left            =   -69960
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
         RecordSource    =   "tbl_sup_record"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgcustomers 
         Bindings        =   "frmview_search.frx":32E66
         Height          =   3855
         Left            =   -75000
         TabIndex        =   28
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
         _Band(0)._NumMapCols=   3
         _Band(0)._MapCol(0)._Name=   "custname"
         _Band(0)._MapCol(0)._Caption=   "Name"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "custcontactno"
         _Band(0)._MapCol(1)._Caption=   "Contact No."
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "custaddress"
         _Band(0)._MapCol(2)._Caption=   "Address"
         _Band(0)._MapCol(2)._RSIndex=   2
      End
      Begin MSAdodcLib.Adodc adodccustomers 
         Height          =   330
         Left            =   -72120
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
         RecordSource    =   "tbl_cust_record"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgproducts 
         Bindings        =   "frmview_search.frx":32E83
         Height          =   3855
         Left            =   -75000
         TabIndex        =   29
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
         _Band(0)._NumMapCols=   9
         _Band(0)._MapCol(0)._Name=   "proid"
         _Band(0)._MapCol(0)._Caption=   "Id"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "proname"
         _Band(0)._MapCol(1)._Caption=   "Name"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "protype"
         _Band(0)._MapCol(2)._Caption=   "Type"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "procompany"
         _Band(0)._MapCol(3)._Caption=   "Company"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(4)._Name=   "prounitprice"
         _Band(0)._MapCol(4)._Caption=   "Unit Price"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(0)._MapCol(4)._Alignment=   7
         _Band(0)._MapCol(5)._Name=   "prounitdiscount"
         _Band(0)._MapCol(5)._Caption=   "Unit Discount"
         _Band(0)._MapCol(5)._RSIndex=   5
         _Band(0)._MapCol(5)._Alignment=   7
         _Band(0)._MapCol(6)._Name=   "progrossamt"
         _Band(0)._MapCol(6)._Caption=   "Gross Amount"
         _Band(0)._MapCol(6)._RSIndex=   6
         _Band(0)._MapCol(6)._Alignment=   7
         _Band(0)._MapCol(7)._Name=   "blt"
         _Band(0)._MapCol(7)._Caption=   "Best Lead Time"
         _Band(0)._MapCol(7)._RSIndex=   7
         _Band(0)._MapCol(8)._Name=   "vat"
         _Band(0)._MapCol(8)._Caption=   "VAT"
         _Band(0)._MapCol(8)._RSIndex=   8
         _Band(0)._MapCol(8)._Alignment=   7
      End
      Begin MSAdodcLib.Adodc adodcproducts 
         Height          =   330
         Left            =   -69360
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
         RecordSource    =   "tbl_pro_record"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgstocks 
         Bindings        =   "frmview_search.frx":32E9F
         Height          =   3855
         Left            =   -75000
         TabIndex        =   30
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
         _Band(0)._NumMapCols=   5
         _Band(0)._MapCol(0)._Name=   "proid"
         _Band(0)._MapCol(0)._Caption=   "Id"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "proname"
         _Band(0)._MapCol(1)._Caption=   "Name"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "protype"
         _Band(0)._MapCol(2)._Caption=   "Type"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "procompany"
         _Band(0)._MapCol(3)._Caption=   "Company"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(4)._Name=   "stkqty"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(0)._MapCol(4)._Alignment=   7
      End
      Begin MSAdodcLib.Adodc adodcstocks 
         Height          =   330
         Left            =   -68040
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
         RecordSource    =   "tbl_stk_record"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgsales 
         Bindings        =   "frmview_search.frx":32EB9
         Height          =   3855
         Left            =   -75000
         TabIndex        =   31
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   3
         Cols            =   18
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   18
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
         _Band(0)._NumMapCols=   18
         _Band(0)._MapCol(0)._Name=   "billno"
         _Band(0)._MapCol(0)._Caption=   "Bill No."
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "custname"
         _Band(0)._MapCol(1)._Caption=   "Name"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "custcontactno"
         _Band(0)._MapCol(2)._Caption=   "Contact No."
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "custaddress"
         _Band(0)._MapCol(3)._Caption=   "Address"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(4)._Name=   "proid"
         _Band(0)._MapCol(4)._Caption=   "Product Id"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(0)._MapCol(5)._Name=   "proname"
         _Band(0)._MapCol(5)._Caption=   "Product Name"
         _Band(0)._MapCol(5)._RSIndex=   5
         _Band(0)._MapCol(6)._Name=   "protype"
         _Band(0)._MapCol(6)._Caption=   "Product Type"
         _Band(0)._MapCol(6)._RSIndex=   6
         _Band(0)._MapCol(7)._Name=   "procompany"
         _Band(0)._MapCol(7)._Caption=   "Product Company"
         _Band(0)._MapCol(7)._RSIndex=   7
         _Band(0)._MapCol(8)._Name=   "qty"
         _Band(0)._MapCol(8)._Caption=   "Quantity"
         _Band(0)._MapCol(8)._RSIndex=   8
         _Band(0)._MapCol(8)._Alignment=   7
         _Band(0)._MapCol(9)._Name=   "unitprice"
         _Band(0)._MapCol(9)._Caption=   "Unit Price"
         _Band(0)._MapCol(9)._RSIndex=   9
         _Band(0)._MapCol(9)._Alignment=   7
         _Band(0)._MapCol(10)._Name=   "unitdiscount"
         _Band(0)._MapCol(10)._Caption=   "Unit Discount"
         _Band(0)._MapCol(10)._RSIndex=   10
         _Band(0)._MapCol(10)._Alignment=   7
         _Band(0)._MapCol(11)._Name=   "grossamt"
         _Band(0)._MapCol(11)._Caption=   "Gross Amount"
         _Band(0)._MapCol(11)._RSIndex=   11
         _Band(0)._MapCol(11)._Alignment=   7
         _Band(0)._MapCol(12)._Name=   "vat"
         _Band(0)._MapCol(12)._Caption=   "VAT"
         _Band(0)._MapCol(12)._RSIndex=   12
         _Band(0)._MapCol(12)._Alignment=   7
         _Band(0)._MapCol(13)._Name=   "netamt"
         _Band(0)._MapCol(13)._Caption=   "Net Amount"
         _Band(0)._MapCol(13)._RSIndex=   13
         _Band(0)._MapCol(13)._Alignment=   7
         _Band(0)._MapCol(14)._Name=   "receivedamt"
         _Band(0)._MapCol(14)._Caption=   "Received Amount"
         _Band(0)._MapCol(14)._RSIndex=   14
         _Band(0)._MapCol(14)._Alignment=   7
         _Band(0)._MapCol(15)._Name=   "dueamt"
         _Band(0)._MapCol(15)._Caption=   "Due Amount"
         _Band(0)._MapCol(15)._RSIndex=   15
         _Band(0)._MapCol(15)._Alignment=   7
         _Band(0)._MapCol(16)._Name=   "date"
         _Band(0)._MapCol(16)._Caption=   "Date"
         _Band(0)._MapCol(16)._RSIndex=   16
         _Band(0)._MapCol(17)._Name=   "salestype"
         _Band(0)._MapCol(17)._Caption=   "Sales Type"
         _Band(0)._MapCol(17)._RSIndex=   17
      End
      Begin MSAdodcLib.Adodc adodcsales 
         Height          =   330
         Left            =   -67440
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
         RecordSource    =   "tbl_sales_record"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgpurchases 
         Bindings        =   "frmview_search.frx":32ED2
         Height          =   3855
         Left            =   -75000
         TabIndex        =   32
         Top             =   1680
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         _Version        =   393216
         ForeColor       =   4210752
         Rows            =   3
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
         _Band(0)._NumMapCols=   16
         _Band(0)._MapCol(0)._Name=   "billno"
         _Band(0)._MapCol(0)._Caption=   "Bill No."
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "supid"
         _Band(0)._MapCol(1)._Caption=   "Supplier Id"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "proid"
         _Band(0)._MapCol(2)._Caption=   "Product Id"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "proname"
         _Band(0)._MapCol(3)._Caption=   "Product Name"
         _Band(0)._MapCol(3)._RSIndex=   3
         _Band(0)._MapCol(4)._Name=   "protype"
         _Band(0)._MapCol(4)._Caption=   "Product Type"
         _Band(0)._MapCol(4)._RSIndex=   4
         _Band(0)._MapCol(5)._Name=   "procompany"
         _Band(0)._MapCol(5)._Caption=   "Product Company"
         _Band(0)._MapCol(5)._RSIndex=   5
         _Band(0)._MapCol(6)._Name=   "qty"
         _Band(0)._MapCol(6)._Caption=   "Quantity"
         _Band(0)._MapCol(6)._RSIndex=   6
         _Band(0)._MapCol(6)._Alignment=   7
         _Band(0)._MapCol(7)._Name=   "unitprice"
         _Band(0)._MapCol(7)._Caption=   "Unit Price"
         _Band(0)._MapCol(7)._RSIndex=   7
         _Band(0)._MapCol(7)._Alignment=   7
         _Band(0)._MapCol(8)._Name=   "unitdiscount"
         _Band(0)._MapCol(8)._Caption=   "Unit Discount"
         _Band(0)._MapCol(8)._RSIndex=   8
         _Band(0)._MapCol(8)._Alignment=   7
         _Band(0)._MapCol(9)._Name=   "grossamt"
         _Band(0)._MapCol(9)._Caption=   "Gross Amount"
         _Band(0)._MapCol(9)._RSIndex=   9
         _Band(0)._MapCol(9)._Alignment=   7
         _Band(0)._MapCol(10)._Name=   "vat"
         _Band(0)._MapCol(10)._Caption=   "VAT"
         _Band(0)._MapCol(10)._RSIndex=   10
         _Band(0)._MapCol(10)._Alignment=   7
         _Band(0)._MapCol(11)._Name=   "netamt"
         _Band(0)._MapCol(11)._Caption=   "Net Amount"
         _Band(0)._MapCol(11)._RSIndex=   11
         _Band(0)._MapCol(11)._Alignment=   7
         _Band(0)._MapCol(12)._Name=   "paidamt"
         _Band(0)._MapCol(12)._Caption=   "Paid Amount"
         _Band(0)._MapCol(12)._RSIndex=   12
         _Band(0)._MapCol(12)._Alignment=   7
         _Band(0)._MapCol(13)._Name=   "dueamt"
         _Band(0)._MapCol(13)._Caption=   "Due Amount"
         _Band(0)._MapCol(13)._RSIndex=   13
         _Band(0)._MapCol(13)._Alignment=   7
         _Band(0)._MapCol(14)._Name=   "date"
         _Band(0)._MapCol(14)._Caption=   "Date"
         _Band(0)._MapCol(14)._RSIndex=   14
         _Band(0)._MapCol(15)._Name=   "purchasestype"
         _Band(0)._MapCol(15)._Caption=   "Purchases Type"
         _Band(0)._MapCol(15)._RSIndex=   15
      End
      Begin MSAdodcLib.Adodc adodcpurchases 
         Height          =   330
         Left            =   -67800
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
         RecordSource    =   "tbl_purchases_record"
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
   End
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmview_search.frx":32EEF
      MousePointer    =   99  'Custom
      Picture         =   "frmview_search.frx":331F9
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmview_search.frx":33C6D
      MousePointer    =   99  'Custom
      Picture         =   "frmview_search.frx":33F77
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
End
Attribute VB_Name = "frmview_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmmain.funcmdview
End Sub

Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub




Private Sub Option1_Click()
If Option1.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option11_Click()
If Option11.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option12_Click()
If Option12.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option13_Click()
If Option13.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option14_Click()
If Option14.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option15_Click()
If Option15.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option16_Click()
If Option16.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option17_Click()
If Option17.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option18_Click()
If Option18.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option19_Click()
If Option19.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option20_Click()
If Option20.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option21_Click()
If Option21.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option22_Click()
If Option22.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option23_Click()
If Option23.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option24_Click()
If Option24.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option25_Click()
If Option25.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option26_Click()
If Option26.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option6_Click()
If Option6.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option7_Click()
If Option7.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
txtsearch.Enabled = True
txtsearch.Text = ""
txtsearch_Change
Else
txtsearch.Enabled = False
End If
End Sub



Private Sub SStab1_Click(PreviousTab As Integer)
If SStab1.Caption = "Employees  " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Suppliers   " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Products    " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Customers   " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Stocks     " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Sales      " Then
txtsearch.Enabled = False
ElseIf SStab1.Caption = "Purchases      " Then
txtsearch.Enabled = False
End If
End Sub

Private Sub txtsearch_Change()
If adodcemployees.Recordset.State = adStateOpen Then adodcemployees.Recordset.Close
If adodcsuppliers.Recordset.State = adStateOpen Then adodcsuppliers.Recordset.Close
If adodccustomers.Recordset.State = adStateOpen Then adodccustomers.Recordset.Close
If adodcproducts.Recordset.State = adStateOpen Then adodcproducts.Recordset.Close
If adodcstocks.Recordset.State = adStateOpen Then adodcstocks.Recordset.Close
If adodcsales.Recordset.State = adStateOpen Then adodcsales.Recordset.Close
If adodcpurchases.Recordset.State = adStateOpen Then adodcpurchases.Recordset.Close

If SStab1.Caption = "Employees  " Then
    If Option1.Value = True Then
    adodcemployees.Recordset.Open "select * from tbl_emp_record where empid like '" & txtsearch.Text & "%'", adodcemployees.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option2.Value = True Then
    adodcemployees.Recordset.Open "select * from tbl_emp_record where empname like '" & txtsearch.Text & "%'", adodcemployees.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option3.Value = True Then
    adodcemployees.Recordset.Open "select * from tbl_emp_record where empcity like '" & txtsearch.Text & "%'", adodcemployees.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option4.Value = True Then
    adodcemployees.Recordset.Open "select * from tbl_emp_record where emppostion like '" & txtsearch.Text & "%'", adodcemployees.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcemployees.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgemployees.DataSource = adodcemployees
        End If
ElseIf SStab1.Caption = "Suppliers   " Then
    If Option5.Value = True Then
    adodcsuppliers.Recordset.Open "select * from tbl_sup_record where supid like '" & txtsearch.Text & "%'", adodcsuppliers.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option6.Value = True Then
    adodcsuppliers.Recordset.Open "select * from tbl_sup_record where supname like '" & txtsearch.Text & "%'", adodcsuppliers.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option7.Value = True Then
    adodcsuppliers.Recordset.Open "select * from tbl_sup_record where supcity like '" & txtsearch.Text & "%'", adodcsuppliers.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcsuppliers.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgsuppliers.DataSource = adodcsuppliers
        End If
    
ElseIf SStab1.Caption = "Customers   " Then
    If Option8.Value = True Then
    adodccustomers.Recordset.Open "select * from tbl_cust_record where custname like '" & txtsearch.Text & "%'", adodccustomers.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodccustomers.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgcustomers.DataSource = adodccustomers
        End If
    
ElseIf SStab1.Caption = "Products    " Then
    If Option9.Value = True Then
    adodcproducts.Recordset.Open "select * from tbl_pro_record where proid like '" & txtsearch.Text & "%'", adodcproducts.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option10.Value = True Then
    adodcproducts.Recordset.Open "select * from tbl_pro_record where proname like '" & txtsearch.Text & "%'", adodcproducts.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option11.Value = True Then
    adodcproducts.Recordset.Open "select * from tbl_pro_record where protype like '" & txtsearch.Text & "%'", adodcproducts.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option12.Value = True Then
    adodcproducts.Recordset.Open "select * from tbl_pro_record where procompany like '" & txtsearch.Text & "%'", adodcproducts.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcproducts.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgproducts.DataSource = adodcproducts
        End If
    
ElseIf SStab1.Caption = "Stocks     " Then
    If Option13.Value = True Then
    adodcstocks.Recordset.Open "select * from tbl_stk_record where proid like '" & txtsearch.Text & "%'", adodcstocks.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option14.Value = True Then
    adodcstocks.Recordset.Open "select * from tbl_stk_record where proname like '" & txtsearch.Text & "%'", adodcstocks.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option15.Value = True Then
    adodcstocks.Recordset.Open "select * from tbl_stk_record where protype like '" & txtsearch.Text & "%'", adodcstocks.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option16.Value = True Then
    adodcstocks.Recordset.Open "select * from tbl_stk_record where procompany like '" & txtsearch.Text & "%'", adodcstocks.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcstocks.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgstocks.DataSource = adodcstocks
        End If
   
ElseIf SStab1.Caption = "Sales      " Then
    If Option17.Value = True Then
    adodcsales.Recordset.Open "select * from tbl_sales_record where billno like '" & txtsearch.Text & "%'", adodcsales.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option18.Value = True Then
    adodcsales.Recordset.Open "select * from tbl_sales_record where custname like '" & txtsearch.Text & "%'", adodcsales.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option19.Value = True Then
    adodcsales.Recordset.Open "select * from tbl_sales_record where proid like '" & txtsearch.Text & "%'", adodcsales.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option20.Value = True Then
    adodcsales.Recordset.Open "select * from tbl_sales_record where proname like '" & txtsearch.Text & "%'", adodcsales.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option21.Value = True Then
    adodcsales.Recordset.Open "select * from tbl_sales_record where salestype like '" & txtsearch.Text & "%'", adodcsales.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcsales.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgsales.DataSource = adodcsales
        End If
    
ElseIf SStab1.Caption = "Purchases      " Then
    If Option22.Value = True Then
    adodcpurchases.Recordset.Open "select * from tbl_purchases_record where billno like '" & txtsearch.Text & "%'", adodcpurchases.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option23.Value = True Then
    adodcpurchases.Recordset.Open "select * from tbl_purchases_record where supid like '" & txtsearch.Text & "%'", adodcpurchases.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option24.Value = True Then
    adodcpurchases.Recordset.Open "select * from tbl_purchases_record where proid like '" & txtsearch.Text & "%'", adodcpurchases.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option25.Value = True Then
    adodcpurchases.Recordset.Open "select * from tbl_purchases_record where proname like '" & txtsearch.Text & "%'", adodcpurchases.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    ElseIf Option26.Value = True Then
    adodcpurchases.Recordset.Open "select * from tbl_purchases_record where purchasestype like '" & txtsearch.Text & "%'", adodcpurchases.Recordset.ActiveConnection, adOpenKeyset, adLockOptimistic
    End If
        If adodcpurchases.Recordset.RecordCount = 0 Then
        MsgBox "No Record Found", vbInformation, "AADIK Technologies"
        txtsearch.Text = ""
        Exit Sub
        Else
        Set hfgpurchases.DataSource = adodcpurchases
        End If
End If

End Sub
