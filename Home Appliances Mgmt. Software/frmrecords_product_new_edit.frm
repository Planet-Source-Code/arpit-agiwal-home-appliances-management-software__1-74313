VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmrecords_product_new_edit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "New/Edit Product Records "
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmrecords_product_new_edit.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtvat 
      DataField       =   "vat"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   26
      Top             =   5640
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmrecords_product_new_edit.frx":32D68
      Height          =   135
      Left            =   4920
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtblt 
      DataField       =   "blt"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   7440
      MaxLength       =   100
      TabIndex        =   11
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox txtnetamt 
      DataField       =   "progrossamt"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtdiscount 
      DataField       =   "prounitdiscount"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   7800
      MaxLength       =   8
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtcompany 
      DataField       =   "procompany"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   7440
      MaxLength       =   100
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txttype 
      DataField       =   "protype"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   3000
      MaxLength       =   100
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtname 
      DataField       =   "proname"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   7440
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtid 
      DataField       =   "proid"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin Candy.CandyButton cmdprevious 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
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
      Picture         =   "frmrecords_product_new_edit.frx":32D89
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
      TabIndex        =   1
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
   Begin MSAdodcLib.Adodc adodcrecords_pro 
      Height          =   330
      Left            =   1800
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
      RecordSource    =   "tbl_pro_record"
      Caption         =   "adodcrecords_pro"
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
      Left            =   6120
      TabIndex        =   12
      Top             =   5640
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
      Left            =   7200
      TabIndex        =   13
      Top             =   5640
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
      Left            =   8280
      TabIndex        =   14
      Top             =   5640
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
   Begin Candy.CandyButton cmdnext 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
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
      Picture         =   "frmrecords_product_new_edit.frx":331DB
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
   Begin Candy.CandyButton cmdnetamtcalc 
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   5040
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
   Begin MSAdodcLib.Adodc adodcrecords_stock 
      Height          =   330
      Left            =   3120
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
      RecordSource    =   "tbl_stk_record"
      Caption         =   "adodcrecords_stock"
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
   Begin VB.TextBox txtgrossamt 
      DataField       =   "prounitprice"
      DataSource      =   "adodcrecords_pro"
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
      Left            =   3360
      MaxLength       =   8
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5040
      Picture         =   "frmrecords_product_new_edit.frx":3362D
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   3000
      Picture         =   "frmrecords_product_new_edit.frx":341D0
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   7440
      Picture         =   "frmrecords_product_new_edit.frx":3532E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3000
      Picture         =   "frmrecords_product_new_edit.frx":3648C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VAT"
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
      Left            =   1200
      TabIndex        =   27
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   1080
      Picture         =   "frmrecords_product_new_edit.frx":375EA
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Best Lead Time"
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
      Left            =   5640
      TabIndex        =   23
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gross Amount"
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
      Left            =   1200
      TabIndex        =   22
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discount"
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
      Left            =   5640
      TabIndex        =   21
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1200
      TabIndex        =   20
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1200
      TabIndex        =   18
      Top             =   3600
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
      Left            =   5640
      TabIndex        =   17
      Top             =   2880
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
      Left            =   1200
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblheader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Records"
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
      Left            =   4170
      TabIndex        =   15
      Top             =   1320
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmrecords_product_new_edit.frx":38F87
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmrecords_product_new_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdcancel_Click()
Dim a As Integer
a = MsgBox("Are you sure to cancel ?", vbYesNo, "AADIK Technologies")
If a = vbYes Then
adodcrecords_pro.Recordset.CancelUpdate
adodcrecords_stock.Recordset.CancelUpdate
mdimain.mnuprorec_Click
End If
End Sub

Private Sub cmddelete_Click()
Dim a As Integer
a = MsgBox("Are you sure to delete this record", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    adodcrecords_pro.Recordset.Delete
    adodcrecords_pro.Recordset.MoveFirst
    adodcrecords_stock.Recordset.Delete
    adodcrecords_stock.Recordset.MoveFirst
    MsgBox "Your record have been successfully deleted", vbOKOnly, "AADIK Technologies"
    End If
If adodcrecords_pro.Recordset.BOF = True Then
mdimain.mnuprorec_Click
End If
End Sub

Private Sub cmdgo_Click()
adodcrecords_pro.Recordset.MoveFirst
adodcrecords_pro.Recordset.Find "proid='" & txtgotoid.Text & "'"
adodcrecords_stock.Recordset.MoveFirst
adodcrecords_stock.Recordset.Find "proid='" & txtgotoid.Text & "'"

    If adodcrecords_pro.Recordset.EOF = True Then
    MsgBox "No Record Found", vbOKOnly, "AADIK Technologies"
    adodcrecords_pro.Recordset.MoveFirst
    adodcrecords_stock.Recordset.MoveFirst
    Exit Sub
    End If
End Sub

Private Sub cmdnetamtcalc_Click()
If Val(txtdiscount.Text) > Val(txtgrossamt.Text) Then
MsgBox "Discount should not be greater than Gross Amount", vbOKOnly, "AADIK Technologies"
Else
txtnetamt.Text = Val(txtgrossamt.Text) - Val(txtdiscount.Text)
End If
End Sub

Private Sub cmdnext_Click()
adodcrecords_pro.Recordset.MoveNext
adodcrecords_stock.Recordset.MoveNext
    If adodcrecords_pro.Recordset.EOF = True Then
    adodcrecords_pro.Recordset.MoveLast
    adodcrecords_stock.Recordset.MoveLast
    End If
End Sub

Private Sub cmdprevious_Click()
adodcrecords_pro.Recordset.MovePrevious
adodcrecords_stock.Recordset.MovePrevious
    If adodcrecords_pro.Recordset.BOF = True Then
    adodcrecords_pro.Recordset.MoveFirst
    adodcrecords_stock.Recordset.MoveFirst
    End If

End Sub

Private Sub cmdsave_Click()
If txtid.Text = "" Or txtname = "" Or txttype.Text = "" Or txtcompany.Text = "" Or txtgrossamt.Text = "" Or txtdiscount.Text = "" Or txtnetamt.Text = "" Or txtblt.Text = "" Or txtvat.Text = "" Then
MsgBox "Please fill all fields", vbOKOnly, "AADIK Technologies"
Exit Sub
Else
Dim a As Integer
adodcrecords_stock.Recordset.Fields("proid") = txtid.Text
adodcrecords_stock.Recordset.Fields("proname") = txtname.Text
adodcrecords_stock.Recordset.Fields("protype") = txttype.Text
adodcrecords_stock.Recordset.Fields("procompany") = txtcompany.Text
adodcrecords_stock.Recordset.Save
adodcrecords_pro.Recordset.Save
MsgBox "Your record have been successfully saved", vbOKOnly, "AADIK Technologies"
a = MsgBox("Are you want to Adding records continue ?", vbYesNo, "AADIK Technologies")
    If a = vbYes Then
    lblheader.Caption = "Add New Product"
    adodcrecords_pro.Recordset.AddNew
    adodcrecords_stock.Recordset.AddNew
    Else
    mdimain.mnuprorec_Click
    End If
    
End If
End Sub



Private Sub Form_Load()
If prorecclick = "New" Then
lblheader.Caption = "Add New Product"
cmddelete.Enabled = False
txtgotoid.Visible = False
cmdgo.Visible = False
cmdprevious.Visible = False
cmdnext.Visible = False
adodcrecords_pro.Recordset.AddNew
adodcrecords_stock.Recordset.AddNew

ElseIf prorecclick = "Edit" Then
lblheader.Caption = "Edit Product Records"
cmddelete.Enabled = True
txtgotoid.Visible = True
cmdgo.Visible = True
cmdprevious.Visible = True
cmdnext.Visible = True
End If
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
Call onlyalpha(KeyAscii)
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
Call onlynumeric(KeyAscii)
End Sub

Private Sub txtgrossamt_KeyPress(KeyAscii As Integer)
Call onlynumeric(KeyAscii)
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
Call onlyalphanum(KeyAscii)
End Sub

Private Sub txtvat_KeyPress(KeyAscii As Integer)
Call onlynumeric(KeyAscii)
End Sub
