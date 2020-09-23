VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Super Seven"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer6 
      Interval        =   300
      Left            =   2190
      Top             =   270
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   300
      Left            =   6240
      TabIndex        =   0
      Top             =   1485
      Width           =   405
   End
   Begin VB.Timer Timer5 
      Interval        =   50
      Left            =   2070
      Top             =   3615
   End
   Begin VB.Timer twosevens 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   225
      Top             =   3660
   End
   Begin VB.Timer WinTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   825
      Top             =   3615
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1185
      Top             =   3750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Transfer"
      Height          =   255
      Left            =   165
      TabIndex        =   9
      ToolTipText     =   "Transfer money from bonusplay to credit"
      Top             =   2910
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+$$$+"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Add cash"
      Top             =   2910
      Width           =   765
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3615
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4200
      Top             =   3885
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3915
      Top             =   3750
   End
   Begin VB.CommandButton Spin 
      Caption         =   "Spin"
      Height          =   270
      Left            =   150
      TabIndex        =   2
      Top             =   2580
      Width           =   2055
   End
   Begin MSComctlLib.ImageList IList1 
      Left            =   7320
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   46
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   55
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
            Object.Tag             =   "2000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28D8
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5041
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77AA
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F13
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C708
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EE98
            Key             =   ""
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":116E6
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13EDB
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":166D0
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18EC5
            Key             =   ""
            Object.Tag             =   "2000"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B79D
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DF2D
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":206BD
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22E4D
            Key             =   ""
            Object.Tag             =   "2000"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25725
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27F7B
            Key             =   ""
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A8A3
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D00C
            Key             =   ""
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F934
            Key             =   ""
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32017
            Key             =   ""
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":346FA
            Key             =   ""
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36DDD
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39546
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3BD9C
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E5F2
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":40E48
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":435D8
            Key             =   ""
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45E26
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4861B
            Key             =   ""
            Object.Tag             =   "140"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4AC50
            Key             =   ""
            Object.Tag             =   "140"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D285
            Key             =   ""
            Object.Tag             =   "140"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F8BA
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5204A
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":547DA
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57030
            Key             =   ""
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59713
            Key             =   ""
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C03B
            Key             =   ""
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E963
            Key             =   ""
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6128B
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":639F4
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":661E9
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":68979
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B109
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D899
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":700EF
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":72945
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7513A
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":778CA
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A05A
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C7C3
            Key             =   ""
            Object.Tag             =   "80"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7EF2C
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":816BC
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":83E4C
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":865DC
            Key             =   ""
            Object.Tag             =   "60"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wild (Not with sevens)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2700
      TabIndex        =   20
      Top             =   2925
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   28
      Left            =   2400
      Picture         =   "Form1.frx":88DD1
      Stretch         =   -1  'True
      Top             =   2895
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x 2 = Bonus 20 - 2000 (Bonusplay)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   2700
      TabIndex        =   19
      Top             =   2625
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   27
      Left            =   2400
      Picture         =   "Form1.frx":8B60F
      Stretch         =   -1  'True
      Top             =   2625
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   26
      Left            =   2400
      Picture         =   "Form1.frx":8DED7
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   25
      Left            =   2700
      Picture         =   "Form1.frx":90657
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   24
      Left            =   3000
      Picture         =   "Form1.frx":92DD7
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  20"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3300
      TabIndex        =   18
      Top             =   2340
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   23
      Left            =   2400
      Picture         =   "Form1.frx":95557
      Stretch         =   -1  'True
      Top             =   2055
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   22
      Left            =   2700
      Picture         =   "Form1.frx":97D9D
      Stretch         =   -1  'True
      Top             =   2055
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   21
      Left            =   3000
      Picture         =   "Form1.frx":9A5E3
      Stretch         =   -1  'True
      Top             =   2055
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  40"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   3300
      TabIndex        =   17
      Top             =   2055
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   20
      Left            =   2400
      Picture         =   "Form1.frx":9CE29
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   19
      Left            =   2700
      Picture         =   "Form1.frx":9F60E
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   18
      Left            =   3000
      Picture         =   "Form1.frx":A1DF3
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  60"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   3300
      TabIndex        =   16
      Top             =   1755
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   17
      Left            =   2400
      Picture         =   "Form1.frx":A45D8
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   16
      Left            =   2700
      Picture         =   "Form1.frx":A6D31
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   15
      Left            =   3000
      Picture         =   "Form1.frx":A948A
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  80"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3300
      TabIndex        =   15
      Top             =   1485
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   14
      Left            =   2400
      Picture         =   "Form1.frx":ABBE3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   13
      Left            =   2700
      Picture         =   "Form1.frx":AE421
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   12
      Left            =   3000
      Picture         =   "Form1.frx":B0C5F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 100"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   3300
      TabIndex        =   14
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   11
      Left            =   2400
      Picture         =   "Form1.frx":B349D
      Stretch         =   -1  'True
      Top             =   915
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   10
      Left            =   2700
      Picture         =   "Form1.frx":B5B70
      Stretch         =   -1  'True
      Top             =   915
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   9
      Left            =   3000
      Picture         =   "Form1.frx":B8243
      Stretch         =   -1  'True
      Top             =   915
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 120"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   3300
      TabIndex        =   13
      Top             =   915
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   8
      Left            =   2400
      Picture         =   "Form1.frx":BA916
      Stretch         =   -1  'True
      Top             =   630
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   7
      Left            =   2700
      Picture         =   "Form1.frx":BCF3B
      Stretch         =   -1  'True
      Top             =   630
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   6
      Left            =   3000
      Picture         =   "Form1.frx":BF560
      Stretch         =   -1  'True
      Top             =   630
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 140"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   3300
      TabIndex        =   12
      Top             =   615
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   5
      Left            =   2400
      Picture         =   "Form1.frx":C1B85
      Stretch         =   -1  'True
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   4
      Left            =   2700
      Picture         =   "Form1.frx":C449D
      Stretch         =   -1  'True
      Top             =   345
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   3
      Left            =   3000
      Picture         =   "Form1.frx":C6DB5
      Stretch         =   -1  'True
      Top             =   345
      Width           =   300
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 200"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3300
      TabIndex        =   11
      Top             =   330
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 2000"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3300
      TabIndex        =   10
      Top             =   60
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   2
      Left            =   3000
      Picture         =   "Form1.frx":C96CD
      Stretch         =   -1  'True
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   1
      Left            =   2700
      Picture         =   "Form1.frx":CBF95
      Stretch         =   -1  'True
      Top             =   60
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   0
      Left            =   2400
      Picture         =   "Form1.frx":CE85D
      Stretch         =   -1  'True
      Top             =   60
      Width           =   300
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1500
      X2              =   1500
      Y1              =   735
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   825
      X2              =   825
      Y1              =   720
      Y2              =   2145
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Credit"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   1500
      TabIndex        =   7
      Top             =   2175
      Width           =   705
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   2355
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "Bonusplay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   90
      Width           =   870
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   435
      TabIndex        =   3
      Top             =   330
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "LCDFont"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   210
      Left            =   180
      TabIndex        =   1
      Top             =   2355
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   150
      Top             =   720
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   1500
      Picture         =   "Form1.frx":D1125
      Top             =   1665
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   825
      Picture         =   "Form1.frx":D39ED
      Top             =   1665
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   135
      Picture         =   "Form1.frx":D62B5
      Top             =   1665
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   1500
      Picture         =   "Form1.frx":D8B7D
      Top             =   1185
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   825
      Picture         =   "Form1.frx":DB445
      Top             =   1185
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   135
      Picture         =   "Form1.frx":DDD0D
      Top             =   1185
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1500
      Picture         =   "Form1.frx":E05D5
      Top             =   720
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   825
      Picture         =   "Form1.frx":E2E9D
      Top             =   720
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   135
      Picture         =   "Form1.frx":E5765
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Last Win"
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   165
      TabIndex        =   6
      Top             =   2160
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command1_Click()
Spin.Enabled = True
Label1.Caption = Label1.Caption + 200
Spin.SetFocus
End Sub



Private Sub Command2_Click()
Dim fsum1 As Integer
Dim fsum2 As Integer
fsum1 = Label1.Caption
fsum2 = Label3.Caption
Label1.Caption = fsum1 + fsum2
Label3.Caption = "0000"
End Sub

Private Sub Form_Load()
Call sndPlaySound(ByVal App.Path & "\intro.wav", SND_ASYNC)
Randomize
r1 = Int((Rnd * 43) + 1)

r2 = Int((Rnd * 43) + 1)

r3 = Int((Rnd * 43) + 1)

Image1(0).Picture = IList1.ListImages(r1).Picture
Image1(1).Picture = IList1.ListImages(r2).Picture
Image1(2).Picture = IList1.ListImages(r3).Picture

Image1(3).Picture = IList1.ListImages(r1 + 1).Picture
Image1(4).Picture = IList1.ListImages(r2 + 1).Picture
Image1(5).Picture = IList1.ListImages(r3 + 1).Picture

Image1(6).Picture = IList1.ListImages(r1 + 2).Picture
Image1(7).Picture = IList1.ListImages(r2 + 2).Picture
Image1(8).Picture = IList1.ListImages(r3 + 2).Picture

cindex1 = r1
cindex2 = r2
cindex3 = r3
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Spin_Click()
Command3.SetFocus
Randomize
hilo = Int((Rnd * 4) + 1)
If hilo = 1 Then
lowwin = True
End If
If hilo = 2 Then
lowwin = False
End If
If hilo = 3 Then
lowwin = True
End If
If hilo = 4 Then
lowwin = True
End If
Spin.Enabled = False
wincash = 0
Call sndPlaySound(ByVal App.Path & "\spin.wav", SND_ASYNC)
If Label3.Caption > 0 Then
Label3.Caption = Label3.Caption - 20
playingup = True
Else
Label1.Caption = Label1.Caption - 10
playingup = False
End If
Randomize
r1 = Int((Rnd * 150) + 101)
Randomize
r2 = Int((Rnd * 50) + 1) + (r1 + 20)
Randomize
r3 = Int((Rnd * 100) + 50) + (r1 + 20)
rounds1 = 100
rounds2 = 100
rounds3 = 100
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True

End Sub

Private Sub Timer1_Timer()
rounds1 = rounds1 + 1
cindex1 = cindex1 + 1
counter1 = cindex1
If counter1 = 56 Then
cindex1 = 1
counter1 = 1
End If
Image1(0).Picture = IList1.ListImages(counter1).Picture

If counter1 = 1 Then
scount = 55
Else
scount = counter1 - 1
End If
Image1(3).Picture = IList1.ListImages(scount).Picture

If scount = 1 Then
xcount = 55
Else
xcount = scount - 1
End If

Image1(6).Picture = IList1.ListImages(xcount).Picture
If rounds1 = r1 Then
Timer1.Enabled = False
Call sndPlaySound(ByVal App.Path & "\stopweel1.wav", SND_ASYNC)

End If
End Sub

Private Sub Timer2_Timer()
rounds2 = rounds2 + 1
cindex2 = cindex2 + 1
counter2 = cindex2
If counter2 = 56 Then
cindex2 = 1
counter2 = 1
End If
Image1(1).Picture = IList1.ListImages(counter2).Picture

If counter2 = 1 Then
scount1 = 55
Else
scount1 = counter2 - 1
End If
Image1(4).Picture = IList1.ListImages(scount1).Picture

If scount1 = 1 Then
xcount1 = 55
Else
xcount1 = scount1 - 1
End If

Image1(7).Picture = IList1.ListImages(xcount1).Picture
If rounds2 = r2 Then
Timer2.Enabled = False
Call sndPlaySound(ByVal App.Path & "\stopweel1.wav", SND_ASYNC)





End If
End Sub

Private Sub Timer3_Timer()

rounds3 = rounds3 + 1
cindex3 = cindex3 + 1
counter3 = cindex3
If counter3 = 56 Then
cindex3 = 1
counter3 = 1
End If
Image1(2).Picture = IList1.ListImages(counter3).Picture

If counter3 = 1 Then
scount2 = 55
Else
scount2 = counter3 - 1
End If
Image1(5).Picture = IList1.ListImages(scount2).Picture

If scount2 = 1 Then
xcount2 = 55
Else
xcount2 = scount2 - 1
End If

Image1(8).Picture = IList1.ListImages(xcount2).Picture
If rounds3 = r3 Then
Timer3.Enabled = False
Call sndPlaySound(ByVal App.Path & "\stopweel1.wav", SND_ASYNC)


neste = False
flashimage = 0
flashimage1 = 1
flashimage2 = 2
checkwin counter1, counter2, counter3
Do
DoEvents
Loop While neste = False
neste = False
flashimage = 3
flashimage1 = 4
flashimage2 = 5
checkwin scount, scount1, scount2
Do
DoEvents
Loop While neste = False
neste = False
flashimage = 6
flashimage1 = 7
flashimage2 = 8
checkwin xcount, xcount1, xcount2
Do
DoEvents
Loop While neste = False
neste = False
flashimage = 0
flashimage1 = 4
flashimage2 = 8
checkwin counter1, scount1, xcount2
Do
DoEvents
Loop While neste = False
neste = False
flashimage = 2
flashimage1 = 4
flashimage2 = 6
checkwin counter3, scount1, xcount
Do
DoEvents
Loop While neste = False
sevenbonus
End If

End Sub

Public Function checkwin(counter1, counter2, counter3)
currentcash = 0
won2000 = False
found3 = False
flashcount = 0
If IList1.ListImages(counter3).Tag = IList1.ListImages(counter2).Tag And IList1.ListImages(counter3).Tag = IList1.ListImages(counter1).Tag Then
If IList1.ListImages(counter3).Tag = 2000 Then
won2000 = True
Call sndPlaySound(ByVal App.Path & "\3sevens1.wav", SND_SYNC)
wincash = 2000
WinTimer = True
Else
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
Timer4.Enabled = True
wincash = IList1.ListImages(counter3).Tag
WinTimer = True
End If
found3 = True
End If

If found3 = False And Not IList1.ListImages(counter1).Tag = 2000 And Not IList1.ListImages(counter2).Tag = 2000 And Not IList1.ListImages(counter3).Tag = 2000 Then
'1Impulse
If IList1.ListImages(counter1).Tag = 100 And IList1.ListImages(counter2).Tag = IList1.ListImages(counter3).Tag Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter2).Tag
WinTimer = True
End If

If IList1.ListImages(counter2).Tag = 100 And IList1.ListImages(counter1).Tag = IList1.ListImages(counter3).Tag Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter3).Tag
WinTimer = True
End If

If IList1.ListImages(counter3).Tag = 100 And IList1.ListImages(counter1).Tag = IList1.ListImages(counter2).Tag Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter2).Tag
WinTimer = True
End If
'2impulse
If IList1.ListImages(counter1).Tag = 100 And IList1.ListImages(counter2).Tag = 100 And Not IList1.ListImages(counter3).Tag = 100 Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter3).Tag
WinTimer = True
End If

If IList1.ListImages(counter1).Tag = 100 And IList1.ListImages(counter3).Tag = 100 And Not IList1.ListImages(counter2).Tag = 100 Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter2).Tag
WinTimer = True
End If

If IList1.ListImages(counter2).Tag = 100 And IList1.ListImages(counter3).Tag = 100 And Not IList1.ListImages(counter1).Tag = 100 Then
Timer4.Enabled = True
Call sndPlaySound(ByVal App.Path & "\stdwin.wav", SND_ASYNC)
wincash = IList1.ListImages(counter1).Tag
WinTimer = True
End If

End If
If Timer4.Enabled = False And WinTimer.Enabled = False Then
neste = True

End If

End Function

Private Sub Timer4_Timer()
neste = False
flashcount = flashcount + 1
If Image1(flashimage).BorderStyle = 1 Then
Image1(flashimage).BorderStyle = 0
Else
Image1(flashimage).BorderStyle = 1
End If

If Image1(flashimage1).BorderStyle = 1 Then
Image1(flashimage1).BorderStyle = 0
Else
Image1(flashimage1).BorderStyle = 1
End If

If Image1(flashimage2).BorderStyle = 1 Then
Image1(flashimage2).BorderStyle = 0
Else
Image1(flashimage2).BorderStyle = 1
End If
If flashcount > 10 Then
Image1(flashimage).BorderStyle = 0
Image1(flashimage1).BorderStyle = 0
Image1(flashimage2).BorderStyle = 0
Timer4.Enabled = False
End If
If Timer4.Enabled = False And WinTimer.Enabled = False Then
neste = True
End If
End Sub

Private Sub Timer5_Timer()
If WinTimer = False And twosevens = False And Timer1.Enabled = False And Timer2.Enabled = False And Timer3.Enabled = False And Timer4.Enabled = False Then
Spin.Enabled = True

End If
If Label1.Caption = 0 And Label3.Caption = 0 Then

Spin.Enabled = False
'Else
'Spin.Enabled = True
End If

End Sub



Private Sub Timer6_Timer()
If Spin.Enabled = True Then
Spin.SetFocus
End If
End Sub

Private Sub twosevens_Timer()
Label3.Caption = Label3.Caption + 20
holdbonus = holdbonus + 20
Label2.Caption = bonus
Call sndPlaySound(ByVal App.Path & "\money.wav", SND_ASYNC)
'
If holdbonus = bonus Or Label3.Caption = 2000 Then

If Label3.Caption = 2000 Then
Call sndPlaySound(ByVal App.Path & "\3sevens.wav", SND_ASYNC)
Dim fsum1 As Integer
Dim fsum2 As Integer
fsum1 = Label1.Caption
fsum2 = Label3.Caption
Label1.Caption = fsum1 + fsum2
Label3.Caption = "0000"
End If
twosevens.Enabled = False

End If
End Sub

Private Sub WinTimer_Timer()
Label3.Caption = Label3.Caption + 20
currentcash = currentcash + 20
Label2.Caption = currentcash
Call sndPlaySound(ByVal App.Path & "\money.wav", SND_ASYNC)
If currentcash = wincash Or Label3.Caption = 2000 Then

WinTimer.Enabled = False
If Label3.Caption = 2000 Then
Call sndPlaySound(ByVal App.Path & "\3sevens.wav", SND_ASYNC)
Dim fsum1 As Integer
Dim fsum2 As Integer
fsum1 = Label1.Caption
fsum2 = Label3.Caption
Label1.Caption = fsum1 + fsum2
Label3.Caption = "0000"
End If
End If
If Timer4.Enabled = False And WinTimer.Enabled = False Then
neste = True

End If
End Sub

Public Function sevenbonus()
twosevens.Enabled = False

Dim sevencount
sevencount = 0
If playingup = True And won2000 = False Then
If IList1.ListImages(counter1).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(counter2).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(counter3).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(scount).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(scount1).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(scount2).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(xcount).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(xcount1).Tag = 2000 Then
sevencount = sevencount + 1
End If
If IList1.ListImages(xcount2).Tag = 2000 Then
sevencount = sevencount + 1
End If

If sevencount = 2 Then
holdbonus = 0
bonus = getbonus
Call sndPlaySound(ByVal App.Path & "\2sevens.wav", SND_SYNC)
twosevens.Enabled = True
Do
DoEvents
Loop While twosevens.Enabled = True
End If
If sevencount = 3 Then
holdbonus = 0
bonus = getbonus
Call sndPlaySound(ByVal App.Path & "\2sevens.wav", SND_SYNC)
twosevens.Enabled = True
Do
DoEvents
Loop While twosevens.Enabled = True
holdbonus = 0
bonus = getbonus
Call sndPlaySound(ByVal App.Path & "\2sevens.wav", SND_SYNC)
twosevens.Enabled = True
Do
DoEvents
Loop While twosevens.Enabled = True
holdbonus = 0
bonus = getbonus
Call sndPlaySound(ByVal App.Path & "\2sevens.wav", SND_SYNC)
twosevens.Enabled = True
Do
DoEvents
Loop While twosevens.Enabled = True
End If
End If
End Function
