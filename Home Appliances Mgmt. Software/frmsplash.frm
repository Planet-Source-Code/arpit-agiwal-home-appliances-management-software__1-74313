VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   4320
      Top             =   1680
   End
   Begin Candy.CandyButton bar2 
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3360
      Width           =   100
      _ExtentX        =   185
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   -1  'True
      ColorButtonHover=   160
      ColorButtonUp   =   128
      ColorButtonDown =   240
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin Candy.CandyButton bar1 
      Height          =   255
      Left            =   1107
      TabIndex        =   1
      Top             =   3360
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   450
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   160
      ColorButtonUp   =   128
      ColorButtonDown =   240
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   4
   End
   Begin VB.Label lblload 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Home Appliances Management Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgsplash 
      Height          =   4215
      Left            =   0
      Picture         =   "frmsplash.frx":2EA5A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
bar2.Width = bar2.Width + 20
lblload.Caption = Int(bar2.Width / 51) & " %"

    If bar2.Width >= 5100 Then
    Timer1.Enabled = False
    Unload Me
    frmlogin.Show
    End If
End Sub
