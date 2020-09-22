VERSION 5.00
Begin VB.Form frmreadme 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Read Me"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmreadme.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.Image cmdback 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmreadme.frx":552D2
      MousePointer    =   99  'Custom
      Picture         =   "frmreadme.frx":555DC
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image cmdmainmenu 
      Height          =   375
      Left            =   6600
      MouseIcon       =   "frmreadme.frx":56050
      MousePointer    =   99  'Custom
      Picture         =   "frmreadme.frx":5635A
      Stretch         =   -1  'True
      ToolTipText     =   "Main Menu"
      Top             =   6480
      Width           =   495
   End
End
Attribute VB_Name = "frmreadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub

Private Sub cmdmainmenu_Click()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
End Sub


