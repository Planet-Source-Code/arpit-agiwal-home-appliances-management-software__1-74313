VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm mdimain 
   BackColor       =   &H00000080&
   Caption         =   "Home Appliances Management Software  -  Arpit Agiwal"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "mdimain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdimain.frx":2EA5A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "mdimain.frx":B73A8
         Height          =   135
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   238
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSAdodcLib.Adodc adodcvh 
         Height          =   330
         Left            =   0
         Top             =   0
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
         RecordSource    =   "tbl_viewhistory"
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
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunotifications 
         Caption         =   "Notifications"
      End
      Begin VB.Menu mnuuseraccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuviewhistory 
         Caption         =   "View History"
      End
      Begin VB.Menu mnulcksys 
         Caption         =   "Locked System"
      End
      Begin VB.Menu mnulogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnurecords 
      Caption         =   "Records"
      Begin VB.Menu mnuemprec 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnusuprec 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnucustrec 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuprorec 
         Caption         =   "Product"
      End
      Begin VB.Menu mnustkrec 
         Caption         =   "Stocks"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "Reports"
      Begin VB.Menu mnureportssales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnureportspurchases 
         Caption         =   "Purchases"
      End
      Begin VB.Menu mnucurstkrep 
         Caption         =   "Current Stock"
      End
   End
   Begin VB.Menu mnutrns 
      Caption         =   "Transactions"
      Begin VB.Menu mnutrnssales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnutrnspur 
         Caption         =   "Purchases"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnusrch 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools"
      Begin VB.Menu mnucalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnunotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuhtlp 
      Caption         =   "Help"
      Begin VB.Menu mnureadme 
         Caption         =   "Read Me"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "mdimain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As Integer
Public Sub MDIForm_Load()
If accounttype = "Administrator" Then
mnuviewhistory.Visible = True
Else
mnuviewhistory.Visible = False
End If
frmmain.Left = 0
frmmain.Top = 0
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

z = MsgBox("Are you sure to Exit the software", vbYesNo, "AADIK Technologies")
If z = vbYes Then
adodcvh.Recordset.AddNew
adodcvh.Recordset.Fields("username") = username
adodcvh.Recordset.Fields("date") = Date
adodcvh.Recordset.Fields("login") = logintime
adodcvh.Recordset.Fields("logout") = Time()
adodcvh.Recordset.Save
Else
Cancel = 1
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
MsgBox "Thank you for using the AADIK Technologies " & vbCrLf & vbCrLf & "                          Arpit Agiwal", vbOKOnly, "AADIK Technologies"
End Sub

Private Sub mnuabout_Click()
UnloadAllForms ("mdimain")
frmaboutus.Left = 0
frmaboutus.Top = 0
End Sub

Public Sub mnucalc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Public Sub mnucomplains_Click()
End Sub


Private Sub mnucurstkrep_Click()
currentstockreport.Show
End Sub

Public Sub mnucustrec_Click()
UnloadAllForms ("mdimain")
frmrecords_cust.Left = 0
frmrecords_cust.Top = 0
End Sub


Public Sub mnuemprec_Click()
UnloadAllForms ("mdimain")
frmrecords_emp.Left = 0
frmrecords_emp.Top = 0
End Sub


Public Sub mnuexit_Click()
Unload Me
End Sub

Public Sub mnulcksys_Click()
frmlocked.Show vbModal
End Sub

Public Sub mnulogout_Click()
Unload Me
If z = vbYes Then
frmlogin.Show
End If
End Sub

Public Sub mnunotepad_Click()
Shell "notepad.exe", vbNormalFocus
End Sub

Public Sub mnunotifications_Click()
UnloadAllForms ("mdimain")
frmnotifications.Left = 0
frmnotifications.Top = 0
End Sub


Public Sub mnuprorec_Click()
UnloadAllForms ("mdimain")
frmrecords_product.Left = 0
frmrecords_product.Top = 0
End Sub

Private Sub mnureadme_Click()
UnloadAllForms ("mdimain")
frmreadme.Left = 0
frmreadme.Top = 0
End Sub

Private Sub mnureportspurchases_Click()
purchasesreport.Show
End Sub

Private Sub mnureportssales_Click()
salesreport.Show
End Sub

Public Sub mnusrch_Click()
UnloadAllForms ("mdimain")
frmview_search.Left = 0
frmview_search.Top = 0
End Sub

Public Sub mnustkrec_Click()
UnloadAllForms ("mdimain")
frmrecords_stock.Left = 0
frmrecords_stock.Top = 0
End Sub

Public Sub mnusuprec_Click()
UnloadAllForms ("mdimain")
frmrecords_sup.Left = 0
frmrecords_sup.Top = 0
End Sub



Public Sub mnutrnspur_Click()
UnloadAllForms ("mdimain")
frmtransactions_purchases.Left = 0
frmtransactions_purchases.Top = 0
End Sub

Public Sub mnutrnssales_Click()
UnloadAllForms ("mdimain")
frmtransactions_sales.Left = 0
frmtransactions_sales.Top = 0
End Sub

Public Sub mnuuseraccounts_Click()
UnloadAllForms ("mdimain")
frmuseraccounts.Left = 0
frmuseraccounts.Top = 0
End Sub



Public Sub mnuviewhistory_Click()
UnloadAllForms ("mdimain")
frmviewhistory.Left = 0
frmviewhistory.Top = 0
End Sub
