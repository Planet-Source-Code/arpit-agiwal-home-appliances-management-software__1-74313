VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmtransactions_purchases_purchasesreturn 
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
   Picture         =   "frmtransactions_purchases_purchasesreturn.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmtransactions_purchases_purchasesreturn.frx":32D68
      Height          =   135
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Bindings        =   "frmtransactions_purchases_purchasesreturn.frx":32D81
      Height          =   135
      Left            =   6720
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc adodctransactions_purchases2 
      Height          =   330
      Left            =   6360
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmtransactions_purchases_purchasesreturn.frx":32DAC
      Height          =   3855
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "billno"
         Caption         =   "Bill No."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "supid"
         Caption         =   "Supplier Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "proid"
         Caption         =   "Product Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "proname"
         Caption         =   "Product Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "protype"
         Caption         =   "Product Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "procompany"
         Caption         =   "Product Company"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "qty"
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "unitprice"
         Caption         =   "Unit Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "unitdiscount"
         Caption         =   "Unit Discount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "grossamt"
         Caption         =   "Gross Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "vat"
         Caption         =   "VAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "netamt"
         Caption         =   "Net Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "paidamt"
         Caption         =   "Paid Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "dueamt"
         Caption         =   "Due Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "purchasestype"
         Caption         =   "Purchases Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtqty 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      MaxLength       =   8
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adodctransactions_purchases 
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_purchases_record where purchasestype='Purchases'"
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
   Begin Candy.CandyButton cmdsave 
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   6000
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
   Begin Candy.CandyButton cmdcancel 
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   6000
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
   Begin MSAdodcLib.Adodc adodcstock 
      Height          =   330
      Left            =   3480
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
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   1200
      Picture         =   "frmtransactions_purchases_purchasesreturn.frx":32DD6
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Change Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label lblheader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Purchases Return"
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
      Left            =   3330
      TabIndex        =   0
      Top             =   1200
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmtransactions_purchases_purchasesreturn.frx":33B3D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmtransactions_purchases_purchasesreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcancel_Click()
Dim a As Integer
a = MsgBox("Are you sure to cancel ?", vbYesNo, "AADIK Technologies")
If a = vbYes Then
adodctransactions_purchases.Recordset.CancelUpdate
mdimain.mnutrnspur_Click
End If
End Sub

Private Sub cmdsave_Click()
Dim stockquantity As Integer, proid As String, selectrec As Integer
selectrec = MsgBox("Are you sure to change the quantity of selected record", vbYesNo, "AADIK Technologies")

If selectrec = vbYes Then
pbillno = adodctransactions_purchases.Recordset.Fields("billno")
proid = adodctransactions_purchases.Recordset.Fields("proid")

If adodctransactions_purchases2.Recordset.BOF = True Then
Else
adodctransactions_purchases2.Recordset.MoveFirst
End If

Do
    If adodctransactions_purchases2.Recordset.Fields("billno") = pbillno And adodctransactions_purchases2.Recordset.Fields("proid") = proid And adodctransactions_purchases2.Recordset.Fields("purchasestype") = "Purchases" Then
    stockquantity = stockquantity + adodctransactions_purchases2.Recordset.Fields("qty").Value
    ElseIf adodctransactions_purchases2.Recordset.Fields("billno") = pbillno And adodctransactions_purchases2.Recordset.Fields("proid") = proid And adodctransactions_purchases2.Recordset.Fields("purchasestype") = "Purchases Return" Then
    stockquantity = stockquantity - adodctransactions_purchases2.Recordset.Fields("qty").Value
    End If
adodctransactions_purchases2.Recordset.MoveNext
Loop Until adodctransactions_purchases2.Recordset.EOF = True
adodctransactions_purchases2.Refresh

        
        If txtqty.Text > stockquantity Then
        MsgBox "Quantity should not be greater than purchase time quantity", vbOKOnly, "AADIK Technologies"
        txtqty.Text = ""
        txtqty.SetFocus
        Exit Sub
        Else
        pbillno = adodctransactions_purchases.Recordset.Fields("billno")
        supid = adodctransactions_purchases.Recordset.Fields("supid")
        proid = adodctransactions_purchases.Recordset.Fields("proid")
        proname = adodctransactions_purchases.Recordset.Fields("proname")
        protype = adodctransactions_purchases.Recordset.Fields("protype")
        procompany = adodctransactions_purchases.Recordset.Fields("procompany")
        qty = Val(txtqty.Text)
        unitprice = adodctransactions_purchases.Recordset.Fields("unitprice")
        unitdiscount = adodctransactions_purchases.Recordset.Fields("unitdiscount")
        grossamt = (unitprice - unitdiscount) * qty
        vat = adodctransactions_purchases.Recordset.Fields("vat")
        netamt = grossamt + (grossamt * vat) / 100
        paidamt = 0
        dueamt = 0
        prdate = Format(Now(), "mm/dd/yyyy")
        
        adodctransactions_purchases.Recordset.AddNew
        adodctransactions_purchases.Recordset.Fields("billno") = pbillno
        adodctransactions_purchases.Recordset.Fields("supid") = supid
        adodctransactions_purchases.Recordset.Fields("proid") = proid
        adodctransactions_purchases.Recordset.Fields("proname") = proname
        adodctransactions_purchases.Recordset.Fields("protype") = protype
        adodctransactions_purchases.Recordset.Fields("procompany") = procompany
        adodctransactions_purchases.Recordset.Fields("qty") = qty
        adodctransactions_purchases.Recordset.Fields("unitprice") = unitprice
        adodctransactions_purchases.Recordset.Fields("unitdiscount") = unitdiscount
        adodctransactions_purchases.Recordset.Fields("grossamt") = grossamt
        adodctransactions_purchases.Recordset.Fields("vat") = vat
        adodctransactions_purchases.Recordset.Fields("netamt") = netamt
        adodctransactions_purchases.Recordset.Fields("paidamt") = paidamt
        adodctransactions_purchases.Recordset.Fields("dueamt") = dueamt
        adodctransactions_purchases.Recordset.Fields("date") = prdate
        adodctransactions_purchases.Recordset.Fields("purchasestype") = "Purchases Return"
        DataGrid1.AllowUpdate = True
        adodctransactions_purchases.Recordset.Update
        DataGrid1.AllowUpdate = False
            If adodcstock.Recordset.BOF = True Then
            Else
            adodcstock.Recordset.MoveFirst
            End If
        adodcstock.Recordset.Find "proid='" & proid & "'"
        stockquantity = adodcstock.Recordset.Fields("stkqty")
        stockquantity = stockquantity - Val(txtqty.Text)
        adodcstock.Recordset.Fields("stkqty").Value = stockquantity
        adodcstock.Recordset.Save
        MsgBox "Record Added Successfully", vbOKOnly, "AADIK Technologies"
        adodcstock.Refresh
        adodctransactions_purchases.Refresh
        DataGrid1.Refresh
        End If
    
    
Else
txtqty.Text = ""
End If
End Sub



Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub
