VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form frmtransactions_expenses 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Expenses Transactions"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmtransactions_expenses.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton cmdmainmenu 
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Main Menu"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   12632256
      ColorButtonUp   =   8421504
      ColorButtonDown =   8421504
      BorderBrightness=   0
      ColorBright     =   14737632
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   960
      Picture         =   "frmtransactions_expenses.frx":2E71F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   9135
   End
End
Attribute VB_Name = "frmtransactions_expenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub
