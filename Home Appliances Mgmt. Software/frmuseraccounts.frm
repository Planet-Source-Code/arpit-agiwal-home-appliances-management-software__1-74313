VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmuseraccounts 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "User Accounts"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmuseraccounts.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton cmdrefresh 
      Height          =   855
      Left            =   8880
      TabIndex        =   12
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Refresh"
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
   Begin Candy.CandyButton cmddelete 
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin Candy.CandyButton cmdshowall 
      Height          =   855
      Left            =   5520
      TabIndex        =   11
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Show All"
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
   Begin MSAdodcLib.Adodc adodcuseraccounts 
      Height          =   330
      Left            =   2760
      Top             =   480
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
      RecordSource    =   "tbl_useraccounts"
      Caption         =   "AdodcUserAccounts"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgusraccounts 
      Bindings        =   "frmuseraccounts.frx":32D68
      Height          =   3255
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   1
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "empname"
      _Band(0)._MapCol(0)._Caption=   "Employee Name"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "username"
      _Band(0)._MapCol(1)._Caption=   "User Name"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "password"
      _Band(0)._MapCol(2)._Caption=   "Password"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Hidden=   -1  'True
      _Band(0)._MapCol(3)._Name=   "accounttype"
      _Band(0)._MapCol(3)._Caption=   "Account Type"
      _Band(0)._MapCol(3)._RSIndex=   3
   End
   Begin Candy.CandyButton cmdsave 
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Style           =   3
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
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin VB.PictureBox framedataentry 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3015
      Left            =   5400
      Picture         =   "frmuseraccounts.frx":32D88
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   13
      Top             =   2160
      Width           =   4455
      Begin VB.TextBox txtempname 
         DataField       =   "empname"
         DataSource      =   "adodcuseraccounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   19
         Top             =   120
         Width           =   2655
      End
      Begin Candy.CandyButton cmdnext 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   2520
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
         Caption         =   ">"
         IconHighLite    =   -1  'True
         IconHighLiteColor=   16777215
         CaptionHighLiteColor=   0
         Picture         =   "frmuseraccounts.frx":36345
         PictureAlignment=   3
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   15728640
         BorderBrightness=   0
         ColorBright     =   16711680
         DisplayHand     =   0   'False
         ColorScheme     =   3
      End
      Begin VB.ComboBox cmbaccounttype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmuseraccounts.frx":36797
         Left            =   1680
         List            =   "frmuseraccounts.frx":367A1
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
         DataField       =   "password"
         DataSource      =   "adodcuseraccounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtusrname 
         DataField       =   "username"
         DataSource      =   "adodcuseraccounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin Candy.CandyButton cmdprevious 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   2520
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
         Caption         =   ""
         IconHighLite    =   -1  'True
         IconHighLiteColor=   16777215
         CaptionHighLiteColor=   0
         Picture         =   "frmuseraccounts.frx":367BE
         PictureAlignment=   2
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   15728640
         BorderBrightness=   0
         ColorBright     =   16711680
         DisplayHand     =   0   'False
         ColorScheme     =   3
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   10
         X1              =   3120
         X2              =   3960
         Y1              =   2640
         Y2              =   3000
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C00000&
         BorderWidth     =   10
         X1              =   600
         X2              =   1440
         Y1              =   3000
         Y2              =   2640
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1650
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1635
      End
   End
   Begin Candy.CandyButton cmdedit 
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin Candy.CandyButton cmdnew 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   3
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmuseraccounts.frx":36C10
      MousePointer    =   99  'Custom
      Picture         =   "frmuseraccounts.frx":36F1A
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmuseraccounts.frx":379E4
      MousePointer    =   99  'Custom
      Picture         =   "frmuseraccounts.frx":37CEE
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   10
      X1              =   8520
      X2              =   9360
      Y1              =   6120
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   10
      X1              =   6120
      X2              =   6840
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Label lbltotalrecords 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   2895
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   1200
      Picture         =   "frmuseraccounts.frx":38762
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Accounts"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   1320
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmuseraccounts.frx":39FEB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmuseraccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdback_Click()
frmmain.funcmdfile
End Sub

Private Sub cmdcancel_Click()
adodcuseraccounts.Recordset.CancelUpdate
adodcuseraccounts.Refresh
hfgusraccounts.Refresh
framedataentry.Enabled = False
cmdnew.Enabled = True
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmddelete.Enabled = False
cmdrefresh.Enabled = True
End Sub

Private Sub cmddelete_Click()
Dim a As Integer
a = MsgBox("Are you sure to delete this record", vbYesNo, "Delete record")
    If a = vbYes Then
    adodcuseraccounts.Recordset.Delete
    adodcuseraccounts.Recordset.MoveFirst
    MsgBox "Your record have been successfully deleted", vbOKOnly, "Successfully Deleted "
    adodcuseraccounts.Refresh
    hfgusraccounts.Refresh
    End If
End Sub

Private Sub cmdedit_Click()
framedataentry.Enabled = True
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdedit.Enabled = False
cmddelete.Enabled = True
cmdrefresh.Enabled = False
txtempname.SetFocus
cmbaccounttype.Text = adodcuseraccounts.Recordset.Fields("accounttype")
End Sub

Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub



Private Sub cmdnew_Click()
framedataentry.Enabled = True
adodcuseraccounts.Recordset.AddNew
cmdnew.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdrefresh.Enabled = False
txtempname.SetFocus
End Sub

Private Sub cmdnext_Click()
adodcuseraccounts.Recordset.MoveNext
    If adodcuseraccounts.Recordset.EOF = True Then
    adodcuseraccounts.Recordset.MoveLast
    End If
cmbaccounttype.Text = adodcuseraccounts.Recordset.Fields("accounttype")
End Sub

Private Sub cmdprevious_Click()
adodcuseraccounts.Recordset.MovePrevious
    If adodcuseraccounts.Recordset.BOF = True Then
    adodcuseraccounts.Recordset.MoveFirst
    End If
cmbaccounttype.Text = adodcuseraccounts.Recordset.Fields("accounttype")
End Sub

Private Sub cmdrefresh_Click()
adodcuseraccounts.Refresh
hfgusraccounts.Refresh
End Sub

Private Sub cmdsave_Click()
If txtempname.Text = "" Or txtusrname.Text = "" Or txtpass.Text = "" Or cmbaccounttype.Text = "" Then
MsgBox "Please fill all fields", vbOKOnly, "AADIK Technologies"
Exit Sub
Else
adodcuseraccounts.Recordset.Fields("accounttype") = cmbaccounttype
adodcuseraccounts.Recordset.Save
MsgBox "Your record have been successfully saved", vbOKOnly, "Successfully Saved "
adodcuseraccounts.Refresh
hfgusraccounts.Refresh

framedataentry.Enabled = False
cmdnew.Enabled = True
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmddelete.Enabled = False
cmdrefresh.Enabled = True
End If
End Sub

Private Sub cmdupdate_Click()
adodcuseraccounts.Recordset.Update
MsgBox "Your record have been successfully saved", vbOKOnly, "Successfully Saved "
txtempname.Text = ""
txtusrname.Text = ""
txtpass.Text = ""
framedataentry.Enabled = False
cmdnew.Enabled = True
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmddelete.Enabled = False

End Sub

Private Sub cmdshowall_Click()
If cmdshowall.Caption = "Show All" Then
hfgusraccounts.Visible = True
lbltotalrecords.Caption = "Total Records : " & adodcuseraccounts.Recordset.RecordCount
lbltotalrecords.Visible = True
cmdshowall.Caption = "Hide All"
ElseIf cmdshowall.Caption = "Hide All" Then
hfgusraccounts.Visible = False
lbltotalrecords.Visible = False
cmdshowall.Caption = "Show All"
End If
End Sub





Private Sub Form_Load()
If accounttype = "Administrator" Then
cmdnew.Enabled = True
cmdedit.Enabled = True
Else
cmdnew.Enabled = False
cmdedit.Enabled = False
End If
cmbaccounttype.Text = adodcuseraccounts.Recordset.Fields("accounttype")
End Sub



