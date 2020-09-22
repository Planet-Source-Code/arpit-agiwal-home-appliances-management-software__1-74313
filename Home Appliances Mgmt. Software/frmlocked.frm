VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmlocked 
   BorderStyle     =   0  'None
   Caption         =   "Locked"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Candy.CandyButton cmdlogin 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   12582912
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin VB.TextBox txtpass 
      DataSource      =   "adodclocked"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin Candy.CandyButton cmdsyslocked 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "System Locked"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   12582912
      Style           =   7
      Checked         =   -1  'True
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin MSAdodcLib.Adodc adodclocked 
      Height          =   330
      Left            =   840
      Top             =   360
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
      Caption         =   "Adodclogin"
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
   Begin VB.Label lblpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   1200
      Picture         =   "frmlocked.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   0
      Picture         =   "frmlocked.frx":BA6A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmlocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdlogin_Click()
adodclocked.Recordset.MoveFirst
adodclocked.Recordset.Find "username='" & username & "'"
    
    If adodclocked.Recordset.EOF = True Then
    MsgBox "Invalid Password", vbOKOnly, "AADIK Technologies"
    txtpass.Text = ""
    txtpass.SetFocus
    ElseIf adodclocked.Recordset.Fields("password") <> txtpass.Text Then
    MsgBox "Invalid Password", vbOKOnly, "AADIK Technologies"
    txtpass.Text = ""
    txtpass.SetFocus
    Else
    Unload Me
    mdimain.Show
    End If
End Sub

