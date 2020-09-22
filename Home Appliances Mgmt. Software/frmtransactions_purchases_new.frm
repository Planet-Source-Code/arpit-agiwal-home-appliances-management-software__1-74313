VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtransactions_purchases_new 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmtransactions_purchases_new.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmtransactions_purchases_new.frx":32D68
      Height          =   135
      Left            =   5160
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc adodcsupplier 
      Height          =   330
      Left            =   2040
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
      RecordSource    =   "tbl_sup_record"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin Candy.CandyButton cmdgo 
      Height          =   390
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Go to Bill no."
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
   Begin Candy.CandyButton cmdprevious 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   1680
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
      Picture         =   "frmtransactions_purchases_new.frx":32D81
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
   Begin Candy.CandyButton cmdnext 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   1680
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
      Picture         =   "frmtransactions_purchases_new.frx":331D3
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
   Begin VB.Frame frameproduct 
      BackColor       =   &H00808080&
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   960
      TabIndex        =   38
      Top             =   3240
      Width           =   9135
      Begin VB.TextBox txtqty 
         DataField       =   "qty"
         DataSource      =   "adodctransactions_purchases"
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
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbpurchasestype 
         DataSource      =   "adodctransactions_purchases"
         Enabled         =   0   'False
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
         ItemData        =   "frmtransactions_purchases_new.frx":33625
         Left            =   5040
         List            =   "frmtransactions_purchases_new.frx":3362F
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txttype 
         BackColor       =   &H00E0E0E0&
         DataField       =   "protype"
         DataSource      =   "adodcproduct"
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtcompany 
         BackColor       =   &H00E0E0E0&
         DataField       =   "procompany"
         DataSource      =   "adodcproduct"
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
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cmbproname 
         Bindings        =   "frmtransactions_purchases_new.frx":33650
         Height          =   360
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "proname"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbproid 
         Bindings        =   "frmtransactions_purchases_new.frx":3366B
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "proid"
         Text            =   ""
      End
      Begin VB.OptionButton optproname 
         BackColor       =   &H00800000&
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
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optproid 
         BackColor       =   &H00800000&
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtdueamt 
         BackColor       =   &H00E0E0E0&
         DataField       =   "dueamt"
         DataSource      =   "adodctransactions_purchases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox txtpaidamt 
         DataField       =   "paidamt"
         DataSource      =   "adodctransactions_purchases"
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
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   27
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtnetamt 
         BackColor       =   &H00E0E0E0&
         DataField       =   "netamt"
         DataSource      =   "adodctransactions_purchases"
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtvat 
         BackColor       =   &H00E0E0E0&
         DataField       =   "vat"
         DataSource      =   "adodcproduct"
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
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtgrossamt 
         BackColor       =   &H00E0E0E0&
         DataField       =   "grossamt"
         DataSource      =   "adodctransactions_purchases"
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtunitdiscount 
         BackColor       =   &H00E0E0E0&
         DataField       =   "prounitdiscount"
         DataSource      =   "adodcproduct"
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
         Left            =   7560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtunitprice 
         BackColor       =   &H00E0E0E0&
         DataField       =   "prounitprice"
         DataSource      =   "adodcproduct"
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
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin Candy.CandyButton cmdnetamtcalc 
         Height          =   375
         Left            =   8520
         TabIndex        =   26
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
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
         IconHighLiteColor=   0
         CaptionHighLite =   -1  'True
         CaptionHighLiteColor=   65535
         ForeColor       =   16777215
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
      Begin Candy.CandyButton cmddueamtcalc 
         Height          =   375
         Left            =   8520
         TabIndex        =   29
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
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
         IconHighLiteColor=   0
         CaptionHighLite =   -1  'True
         CaptionHighLiteColor=   65535
         ForeColor       =   16777215
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
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8640
         Picture         =   "frmtransactions_purchases_new.frx":33686
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   5040
         Picture         =   "frmtransactions_purchases_new.frx":34229
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   5040
         Picture         =   "frmtransactions_purchases_new.frx":35387
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   5040
         Picture         =   "frmtransactions_purchases_new.frx":364E5
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   5040
         Picture         =   "frmtransactions_purchases_new.frx":37643
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Gross Amount "
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
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   9120
         X2              =   9120
         Y1              =   120
         Y2              =   0
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Purchases Type "
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
         Left            =   120
         TabIndex        =   49
         Top             =   2760
         Width           =   4935
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Due Amount "
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
         Left            =   120
         TabIndex        =   48
         Top             =   2400
         Width           =   4935
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Paid Amount "
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
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Net Amount "
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
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "VAT "
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
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Unit Dis."
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
         Left            =   7560
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Unit Price"
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
         Left            =   6240
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Quantity"
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
         Left            =   5040
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Company"
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
         Left            =   3600
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Type"
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
         Left            =   2280
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame framesupplier 
      BackColor       =   &H00808080&
      Caption         =   "Supplier Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   960
      TabIndex        =   34
      Top             =   1920
      Width           =   9135
      Begin VB.OptionButton optsupid 
         BackColor       =   &H00800000&
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optsupname 
         BackColor       =   &H00800000&
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
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtbillno 
         BackColor       =   &H00FFFFFF&
         DataField       =   "billno"
         DataSource      =   "adodctransactions_purchases"
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
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtcontactno 
         BackColor       =   &H00E0E0E0&
         DataField       =   "supcontactno"
         DataSource      =   "adodcsupplier"
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
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtemailid 
         BackColor       =   &H00E0E0E0&
         DataField       =   "supemailid"
         DataSource      =   "adodcsupplier"
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
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtcity 
         BackColor       =   &H00E0E0E0&
         DataField       =   "supcity"
         DataSource      =   "adodcsupplier"
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
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpdate 
         Height          =   375
         Left            =   7440
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   16646145
         CurrentDate     =   40948
      End
      Begin MSDataListLib.DataCombo cmbsupname 
         Bindings        =   "frmtransactions_purchases_new.frx":387A1
         Height          =   360
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "supname"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbsupid 
         Bindings        =   "frmtransactions_purchases_new.frx":387BD
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "supid"
         Text            =   ""
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "AADIK Technologies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   52
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Date  "
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
         Left            =   6720
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Bill No."
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
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Contact no."
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
         Left            =   3120
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Email Id"
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
         Left            =   4920
         TabIndex        =   36
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         Left            =   6960
         TabIndex        =   35
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc adodctransactions_purchases 
      Height          =   330
      Left            =   6120
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
      RecordSource    =   "tbl_purchases_record"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adodcstock 
      Height          =   330
      Left            =   4680
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
      RecordSource    =   "tbl_stk_record"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adodcproduct 
      Height          =   330
      Left            =   3360
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
      RecordSource    =   "tbl_pro_record"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Candy.CandyButton cmdsave 
      Height          =   375
      Left            =   3240
      TabIndex        =   31
      Top             =   6480
      Width           =   1095
      _ExtentX        =   1931
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
      Left            =   6720
      TabIndex        =   32
      Top             =   6480
      Width           =   1095
      _ExtentX        =   1931
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
      Left            =   7920
      TabIndex        =   33
      Top             =   6480
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.Image Image2 
      Height          =   855
      Left            =   1200
      Picture         =   "frmtransactions_purchases_new.frx":387D9
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblheader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Purchases"
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
      Left            =   4185
      TabIndex        =   0
      Top             =   1200
      Width           =   3030
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmtransactions_purchases_new.frx":39540
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmtransactions_purchases_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbproid_Click(Area As Integer)
If adodcproduct.Recordset.BOF = True Then
Exit Sub
Else
adodcproduct.Recordset.MoveFirst
adodcproduct.Recordset.Find "proid='" & cmbproid & "'"
    If adodcproduct.Recordset.EOF = True Then
    Else
    cmbproname.Text = adodcproduct.Recordset.Fields("proname")
    End If
End If
End Sub

Private Sub cmbproname_Click(Area As Integer)
If adodcproduct.Recordset.BOF = True Then
Exit Sub
Else
adodcproduct.Recordset.MoveFirst
adodcproduct.Recordset.Find "proname='" & cmbproname & "'"
    If adodcproduct.Recordset.EOF = True Then
    Else
    cmbproid.Text = adodcproduct.Recordset.Fields("proid")
    End If
End If
End Sub

Private Sub cmbsupid_Click(Area As Integer)
If adodcsupplier.Recordset.BOF = True Then
Exit Sub
Else
adodcsupplier.Recordset.MoveFirst
adodcsupplier.Recordset.Find "supid='" & cmbsupid & "'"
    If adodcsupplier.Recordset.EOF = True Then
    Else
    cmbsupname.Text = adodcsupplier.Recordset.Fields("supname")
    End If
End If
End Sub

Private Sub cmbsupname_Click(Area As Integer)
If adodcsupplier.Recordset.BOF = True Then
Exit Sub
Else
adodcsupplier.Recordset.MoveFirst
adodcsupplier.Recordset.Find "supname='" & cmbsupname & "'"
    If adodcsupplier.Recordset.EOF = True Then
    Else
    cmbsupid.Text = adodcsupplier.Recordset.Fields("supid")
    End If
End If
End Sub

Private Sub cmdcancel_Click()
Dim a As Integer
a = MsgBox("Are you sure to cancel ?", vbYesNo, "AADIK Technologies")
If a = vbYes Then
adodctransactions_purchases.Recordset.CancelUpdate
mdimain.mnutrnspur_Click
End If
End Sub

Private Sub cmddelete_Click()
Dim a As Integer
a = MsgBox("Are you sure to delete this record", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    adodctransactions_purchases.Recordset.Delete
    adodctransactions_purchases.Recordset.MoveFirst
    MsgBox "Your record have been successfully deleted", vbOKOnly, "AADIK Technologies"
    End If
If adodctransactions_purchases.Recordset.BOF = True Then
mdimain.mnutrnspur_Click
End If
End Sub

Private Sub cmddueamtcalc_Click()
If Val(txtpaidamt.Text) > Val(txtnetamt.Text) Then
MsgBox "Paid Amount should not be greater than Net Amount", vbOKOnly, "AADIK Technologies"
txtpaidamt.Text = ""
txtpaidamt.SetFocus
Else
txtdueamt.Text = Val(txtnetamt.Text) - Val(txtpaidamt.Text)
End If
End Sub

Private Sub cmdgo_Click()
If adodctransactions_purchases.Recordset.BOF = True Then
Else
adodctransactions_purchases.Recordset.MoveFirst
End If
adodctransactions_purchases.Recordset.Find "billno='" & txtgotoid.Text & "'"
    If adodctransactions_purchases.Recordset.EOF = True Then
    MsgBox "No Record Found", vbOKOnly, "AADIK Technologies"
    adodctransactions_purchases.Recordset.MoveFirst
    Exit Sub
    End If
cmbsupid.Text = adodctransactions_purchases.Recordset.Fields("supid")

If adodcsupplier.Recordset.BOF = True Then
Exit Sub
Else
adodcsupplier.Recordset.MoveFirst
adodcsupplier.Recordset.Find "supid='" & cmbsupid & "'"
    If adodcsupplier.Recordset.EOF = True Then
    Else
    cmbsupname.Text = adodcsupplier.Recordset.Fields("supname")
    txtcontactno.Text = adodcsupplier.Recordset.Fields("supcontactno")
    txtemailid.Text = adodcsupplier.Recordset.Fields("supemailid")
    txtcity.Text = adodcsupplier.Recordset.Fields("supcity")
    End If
End If



cmbproid.Text = adodctransactions_purchases.Recordset.Fields("proid")
cmbproname.Text = adodctransactions_purchases.Recordset.Fields("proname")
txttype.Text = adodctransactions_purchases.Recordset.Fields("protype")
txtcompany.Text = adodctransactions_purchases.Recordset.Fields("procompany")
txtunitprice.Text = adodctransactions_purchases.Recordset.Fields("unitprice")
txtunitdiscount.Text = adodctransactions_purchases.Recordset.Fields("unitdiscount")
txtgrossamt.Text = adodctransactions_purchases.Recordset.Fields("grossamt")
txtvat.Text = adodctransactions_purchases.Recordset.Fields("vat")
dtpdate.Value = adodctransactions_purchases.Recordset.Fields("date")
cmbpurchasestype.Text = adodctransactions_purchases.Recordset.Fields("purchasestype")

If cmbpurchasestype.Text = "Purchases Return" Then
framesupplier.Enabled = False
frameproduct.Enabled = False
Else
framesupplier.Enabled = True
frameproduct.Enabled = True
End If
End Sub

Private Sub cmdnetamtcalc_Click()
txtgrossamt.Text = (Val(txtunitprice.Text) - Val(txtunitdiscount)) * Val(txtqty.Text)
txtnetamt.Text = Val(txtgrossamt.Text) + (Val(txtgrossamt.Text) * Val(txtvat)) / 100
End Sub

Private Sub cmdsave_Click()
If txtbillno.Text = "" Or dtpdate.Value = "" Or cmbsupid.Text = "" Or cmbsupname.Text = "" Or cmbproid.Text = "" Or cmbproname.Text = "" Or txtqty.Text = "" Or txtnetamt.Text = "" Or txtpaidamt.Text = "" Or txtdueamt.Text = "" Or cmbpurchasestype = "" Then
MsgBox "Please fill all fields", vbOKOnly, "AADIK Technologies"
Exit Sub
Else
Dim a As Integer, number As Integer

pbillno = txtbillno.Text
pdate = dtpdate.Value
psupid = cmbsupid.Text
psupname = cmbsupname.Text
psupcontactno = txtcontactno.Text
psupcity = txtcity.Text
psupemailid = txtemailid.Text
    
    
    Dim stockquantity As Integer
    If adodcstock.Recordset.BOF = True Then
    Else
    adodcstock.Recordset.MoveFirst
    End If
    
    adodcstock.Recordset.Find "proid='" & cmbproid & "'"
    If cmbpurchasestype.Text = "Purchases" Then
        stockquantity = adodcstock.Recordset.Fields("stkqty").Value
        stockquantity = stockquantity + Val(txtqty.Text)
        adodcstock.Recordset.Fields("stkqty").Value = stockquantity
        adodcstock.Recordset.Save
      
    End If
    

adodctransactions_purchases.Recordset.Fields("billno") = pbillno
adodctransactions_purchases.Recordset.Fields("supid") = psupid
adodctransactions_purchases.Recordset.Fields("proid") = cmbproid.Text
adodctransactions_purchases.Recordset.Fields("proname") = cmbproname.Text
adodctransactions_purchases.Recordset.Fields("protype") = txttype.Text
adodctransactions_purchases.Recordset.Fields("procompany") = txtcompany.Text
adodctransactions_purchases.Recordset.Fields("unitprice") = txtunitprice.Text
adodctransactions_purchases.Recordset.Fields("unitdiscount") = txtunitdiscount.Text
adodctransactions_purchases.Recordset.Fields("grossamt") = txtgrossamt.Text
adodctransactions_purchases.Recordset.Fields("vat") = txtvat.Text
adodctransactions_purchases.Recordset.Fields("date") = dtpdate.Value
adodctransactions_purchases.Recordset.Fields("purchasestype") = cmbpurchasestype.Text


adodctransactions_purchases.Recordset.Save




MsgBox "Your record have been successfully saved", vbOKOnly, "AADIK Technologies"
a = MsgBox("Are you want to Add more product ?", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    lblheader.Caption = "Add New Product"
    framesupplier.Enabled = False
    adodctransactions_purchases.Recordset.AddNew
    txtbillno.Text = pbillno
    dtpdate.Value = pdate
    cmbsupid.Text = psupid
    cmbsupname.Text = psupname
    txtcontactno.Text = psupcontactno
    txtcity.Text = psupcity
    txtemailid.Text = psupemailid
    Exit Sub
    End If
    
     

    a = MsgBox("Are you want to Add more purchases record ?", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    lblheader.Caption = "Add New Product"
    framesupplier.Enabled = True
    
    
    lblheader.Caption = "Add New Purchases"
    cmddelete.Enabled = False
    txtgotoid.Visible = False
    cmdgo.Visible = False
    cmdprevious.Visible = False
    cmdnext.Visible = False
    
    adodctransactions_purchases.Recordset.AddNew

    
    Else
    mdimain.mnutrnspur_Click
    End If
    
End If


End Sub

Private Sub cmdnext_Click()
txtgotoid.Text = ""
adodctransactions_purchases.Recordset.MoveNext
    If adodctransactions_purchases.Recordset.EOF = True Then
    adodctransactions_purchases.Recordset.MoveLast
    End If

cmbsupid.Text = adodctransactions_purchases.Recordset.Fields("supid")

If adodcsupplier.Recordset.BOF = True Then
Exit Sub
Else
adodcsupplier.Recordset.MoveFirst
adodcsupplier.Recordset.Find "supid='" & cmbsupid & "'"
    If adodcsupplier.Recordset.EOF = True Then
    Else
    cmbsupname.Text = adodcsupplier.Recordset.Fields("supname")
    txtcontactno.Text = adodcsupplier.Recordset.Fields("supcontactno")
    txtemailid.Text = adodcsupplier.Recordset.Fields("supemailid")
    txtcity.Text = adodcsupplier.Recordset.Fields("supcity")
    End If
End If

cmbproid.Text = adodctransactions_purchases.Recordset.Fields("proid")
cmbproname.Text = adodctransactions_purchases.Recordset.Fields("proname")
txttype.Text = adodctransactions_purchases.Recordset.Fields("protype")
txtcompany.Text = adodctransactions_purchases.Recordset.Fields("procompany")
txtunitprice.Text = adodctransactions_purchases.Recordset.Fields("unitprice")
txtunitdiscount.Text = adodctransactions_purchases.Recordset.Fields("unitdiscount")
txtgrossamt.Text = adodctransactions_purchases.Recordset.Fields("grossamt")
txtvat.Text = adodctransactions_purchases.Recordset.Fields("vat")
dtpdate.Value = adodctransactions_purchases.Recordset.Fields("date")
cmbpurchasestype.Text = adodctransactions_purchases.Recordset.Fields("purchasestype")

If cmbpurchasestype.Text = "Purchases Return" Then
framesupplier.Enabled = False
frameproduct.Enabled = False
Else
framesupplier.Enabled = True
frameproduct.Enabled = True
End If
End Sub

Private Sub cmdprevious_Click()
txtgotoid.Text = ""
adodctransactions_purchases.Recordset.MovePrevious
    If adodctransactions_purchases.Recordset.BOF = True Then
    adodctransactions_purchases.Recordset.MoveFirst
    End If
cmbsupid.Text = adodctransactions_purchases.Recordset.Fields("supid")

If adodcsupplier.Recordset.BOF = True Then
Exit Sub
Else
adodcsupplier.Recordset.MoveFirst
adodcsupplier.Recordset.Find "supid='" & cmbsupid & "'"
    If adodcsupplier.Recordset.EOF = True Then
    Else
    cmbsupname.Text = adodcsupplier.Recordset.Fields("supname")
    txtcontactno.Text = adodcsupplier.Recordset.Fields("supcontactno")
    txtemailid.Text = adodcsupplier.Recordset.Fields("supemailid")
    txtcity.Text = adodcsupplier.Recordset.Fields("supcity")
    End If
End If

cmbproid.Text = adodctransactions_purchases.Recordset.Fields("proid")
cmbproname.Text = adodctransactions_purchases.Recordset.Fields("proname")
txttype.Text = adodctransactions_purchases.Recordset.Fields("protype")
txtcompany.Text = adodctransactions_purchases.Recordset.Fields("procompany")
txtunitprice.Text = adodctransactions_purchases.Recordset.Fields("unitprice")
txtunitdiscount.Text = adodctransactions_purchases.Recordset.Fields("unitdiscount")
txtgrossamt.Text = adodctransactions_purchases.Recordset.Fields("grossamt")
txtvat.Text = adodctransactions_purchases.Recordset.Fields("vat")
dtpdate.Value = adodctransactions_purchases.Recordset.Fields("date")
cmbpurchasestype.Text = adodctransactions_purchases.Recordset.Fields("purchasestype")

If cmbpurchasestype.Text = "Purchases Return" Then
framesupplier.Enabled = False
frameproduct.Enabled = False
Else
framesupplier.Enabled = True
frameproduct.Enabled = True
End If
End Sub

Private Sub Form_Load()
dtpdate.MaxDate = Now()
dtpdate.Value = Format(Now(), "mm/dd/yy")

If purchasesclick = "New" Then
lblheader.Caption = "Add New Purchases"
cmddelete.Enabled = False
txtgotoid.Visible = False
cmdgo.Visible = False
cmdprevious.Visible = False
cmdnext.Visible = False
adodctransactions_purchases.Recordset.AddNew
cmbpurchasestype.Text = "Purchases"

ElseIf purchasesclick = "Edit" Then
lblheader.Caption = "Edit Purchases Records"
cmddelete.Enabled = True
txtgotoid.Visible = True
cmdgo.Visible = True
cmdprevious.Visible = True
cmdnext.Visible = True

End If


End Sub

Private Sub optproid_Click()
If optproid.Value = True Then
cmbproid.Enabled = True
cmbproname.Enabled = False
End If
End Sub

Private Sub optproname_Click()
If optproname.Value = True Then
cmbproid.Enabled = False
cmbproname.Enabled = True
End If
End Sub

Private Sub optsupid_Click()
If optsupid.Value = True Then
cmbsupid.Enabled = True
cmbsupname.Enabled = False
End If
End Sub

Private Sub optsupname_Click()
If optsupname.Value = True Then
cmbsupid.Enabled = False
cmbsupname.Enabled = True
End If
End Sub



Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub
