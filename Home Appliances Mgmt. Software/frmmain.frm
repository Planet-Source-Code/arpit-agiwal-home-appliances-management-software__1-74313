VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmmain.frx":2EA5A
   ScaleHeight     =   7800
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox frametools 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":617C2
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdcalculator 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Calculator"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdnotepad 
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Notepad"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lbltools 
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tools"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox frameview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":65008
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdsearch 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Search"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lblview 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "View"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox frametransactions 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":6884E
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdsales 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sales"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdpurchases 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Purchases"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lbltransactions 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Transactions"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox framereports 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":6C094
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdsalesreports 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sales"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdpurchasesreports 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Purchases"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdcurrentstock 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Current Stock"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lblreports 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Reports"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox framerecords 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":6F8DA
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdemployee 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Employee"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdsupplier 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Supplier"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdcustomer 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Customer"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdproduct 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Product"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdstock 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stock"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lblrecords 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Records"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.PictureBox framefile 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5760
      Picture         =   "frmmain.frx":73120
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin Candy.CandyButton cmdnotifications 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Notifications"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   4210752
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmduseraccounts 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "User A/C"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdhistory 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "History"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdlocked 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Locked"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton cmdlogout 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Log Out"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   65535
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Candy.CandyButton lblfile 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "File"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   14737632
         Style           =   4
         Checked         =   -1  'True
         ColorButtonHover=   10485760
         ColorButtonUp   =   8388608
         ColorButtonDown =   8388608
         BorderBrightness=   0
         ColorBright     =   8388608
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin Candy.CandyButton cmdfile 
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "File"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Candy.CandyButton cmdrecords 
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Records"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Candy.CandyButton cmdreports 
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reports"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Candy.CandyButton cmdtransactions 
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Transactions"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Candy.CandyButton cmdview 
      Height          =   975
      Left            =   1320
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Candy.CandyButton cmdtools 
      Height          =   975
      Left            =   3360
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tools"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   10485760
      ColorButtonUp   =   8388608
      ColorButtonDown =   15728640
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "A R P I T    A G I W A L "
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
      Height          =   255
      Left            =   2880
      TabIndex        =   37
      Top             =   5880
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "D I N E S H   K U L D E E P"
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
      Height          =   4095
      Left            =   5400
      TabIndex        =   36
      Top             =   1680
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   8520
      X2              =   2520
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   5520
      X2              =   5520
      Y1              =   1440
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   5520
      X2              =   5520
      Y1              =   1440
      Y2              =   6000
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdaccording_Click()

End Sub

Private Sub cmdcalculator_Click()
Shell "calc.exe", vbNormalFocus
End Sub



Private Sub cmdcurrentstock_Click()
currentstockreport.Show
End Sub

Private Sub cmdcustomer_Click()
UnloadAllForms ("mdimain")
frmrecords_cust.Left = 0
frmrecords_cust.Top = 0
End Sub

Private Sub cmdemployee_Click()
UnloadAllForms ("mdimain")
frmrecords_emp.Left = 0
frmrecords_emp.Top = 0
End Sub



Private Sub cmdfile_Click()
framefile.Visible = True
framerecords.Visible = False
framereports.Visible = False
frametransactions.Visible = False
frameview.Visible = False
frametools.Visible = False
End Sub
Public Function funcmdfile()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdfile_Click
End Function

Private Sub cmdhistory_Click()
UnloadAllForms ("mdimain")
frmviewhistory.Left = 0
frmviewhistory.Top = 0
End Sub

Private Sub cmdlocked_Click()
frmlocked.Show vbModal
End Sub

Private Sub cmdlogout_Click()
Unload mdimain
If z = vbYes Then
frmlogin.Show
End If
End Sub

Private Sub cmdnotepad_Click()
Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub cmdnotifications_Click()
UnloadAllForms ("mdimain")
frmnotifications.Left = 0
frmnotifications.Top = 0
End Sub

Private Sub cmdproduct_Click()
UnloadAllForms ("mdimain")
frmrecords_product.Left = 0
frmrecords_product.Top = 0
End Sub

Private Sub cmdpurchases_Click()
UnloadAllForms ("mdimain")
frmtransactions_purchases.Left = 0
frmtransactions_purchases.Top = 0
End Sub

Private Sub cmdpurchasesreports_Click()
purchasesreport.Show
End Sub

Private Sub cmdrecords_Click()
framefile.Visible = False
framerecords.Visible = True
framereports.Visible = False
frametransactions.Visible = False
frameview.Visible = False
frametools.Visible = False
End Sub
Public Function funcmdrecords()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdrecords_Click
End Function

Private Sub cmdreports_Click()
framefile.Visible = False
framerecords.Visible = False
framereports.Visible = True
frametransactions.Visible = False
frameview.Visible = False
frametools.Visible = False
End Sub
Public Function funcmdreports()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdreports_Click
End Function

Private Sub cmdsales_Click()
UnloadAllForms ("mdimain")
frmtransactions_sales.Left = 0
frmtransactions_sales.Top = 0
End Sub



Private Sub cmdsalesreports_Click()
salesreport.Show
End Sub

Private Sub cmdsearch_Click()
UnloadAllForms ("mdimain")
frmview_search.Left = 0
frmview_search.Top = 0
End Sub

Private Sub cmdstock_Click()
UnloadAllForms ("mdimain")
frmrecords_stock.Left = 0
frmrecords_stock.Top = 0
End Sub

Private Sub cmdsupplier_Click()
UnloadAllForms ("mdimain")
frmrecords_sup.Left = 0
frmrecords_sup.Top = 0
End Sub

Private Sub cmdtools_Click()
framefile.Visible = False
framerecords.Visible = False
framereports.Visible = False
frametransactions.Visible = False
frameview.Visible = False
frametools.Visible = True
End Sub
Public Function funcmdtools()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdtools_Click
End Function

Private Sub cmdtransactions_Click()
framefile.Visible = False
framerecords.Visible = False
framereports.Visible = False
frametransactions.Visible = True
frameview.Visible = False
frametools.Visible = False
End Sub
Public Function funcmdtransactions()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdtransactions_Click
End Function

Private Sub cmduseraccounts_Click()
UnloadAllForms ("mdimain")
frmuseraccounts.Left = 0
frmuseraccounts.Top = 0
End Sub

Private Sub cmdview_Click()
framefile.Visible = False
framerecords.Visible = False
framereports.Visible = False
frametransactions.Visible = False
frameview.Visible = True
frametools.Visible = False
End Sub
Public Function funcmdview()
UnloadAllForms ("mdimain")
frmmain.Left = 0
frmmain.Top = 0
cmdview_Click
End Function




Private Sub Form_Load()
If accounttype = "Administrator" Then
cmdhistory.Enabled = True
Else
cmdhistory.Enabled = False
End If
End Sub
