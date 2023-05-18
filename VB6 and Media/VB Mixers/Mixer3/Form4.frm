VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email me at micracom2@hotmail.com"
   ClientHeight    =   4395
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   30
      Top             =   0
   End
   Begin VB.PictureBox Picture16 
      Height          =   915
      Left            =   1830
      Picture         =   "Form4.frx":030A
      ScaleHeight     =   855
      ScaleWidth      =   1185
      TabIndex        =   108
      Top             =   60
      Width           =   1250
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   540
         X2              =   180
         Y1              =   600
         Y2              =   180
      End
      Begin VB.Image Image1 
         Height          =   150
         Left            =   510
         Picture         =   "Form4.frx":0C74
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin VB.PictureBox Picture15 
      Height          =   915
      Left            =   6930
      Picture         =   "Form4.frx":0D46
      ScaleHeight     =   855
      ScaleWidth      =   1185
      TabIndex        =   107
      Top             =   30
      Width           =   1250
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   540
         X2              =   180
         Y1              =   600
         Y2              =   180
      End
      Begin VB.Image Image2 
         Height          =   150
         Left            =   510
         Picture         =   "Form4.frx":16B0
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin MSComctlLib.Slider Slider11 
      Height          =   225
      Left            =   1590
      TabIndex        =   98
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider12 
      Height          =   225
      Left            =   2325
      TabIndex        =   99
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldPan 
      Height          =   225
      Left            =   90
      TabIndex        =   96
      Top             =   3180
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider10 
      Height          =   225
      Left            =   825
      TabIndex        =   97
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin VB.TextBox BassText 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9180
      TabIndex        =   92
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1740
      Width           =   345
   End
   Begin VB.TextBox Treblesliderte 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   8490
      TabIndex        =   91
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1740
      Width           =   345
   End
   Begin VB.PictureBox Picture10 
      Height          =   330
      Left            =   6315
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   87
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option20 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option19 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   135
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   330
      Left            =   7080
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   84
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option18 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   135
         Width           =   195
      End
      Begin VB.OptionButton Option17 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   330
      Left            =   270
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   81
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option2 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   135
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   330
      Left            =   4050
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   78
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option3 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   135
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   330
      Left            =   1845
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   75
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option5 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   135
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   330
      Left            =   2550
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   72
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option7 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   135
         Width           =   195
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   330
      Left            =   5580
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   69
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option9 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   135
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   330
      Left            =   4815
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   66
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option11 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   135
         Width           =   195
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   330
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   63
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option13 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   135
         Width           =   195
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture8 
      Height          =   330
      Left            =   3315
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   60
      Top             =   4005
      Width           =   240
      Begin VB.OptionButton Option15 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option16 
         BackColor       =   &H000000FF&
         Height          =   150
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   135
         Width           =   195
      End
   End
   Begin VB.TextBox timeWindow 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   570
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   60
      Width           =   3705
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   6960
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1770
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   6270
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5505
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4770
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3270
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2505
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1005
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox txtMasterVolume 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   270
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.TextBox txtWaveOutVolume 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4005
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "32768"
      Top             =   1755
      Width           =   345
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   270
      TabIndex        =   22
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   1080
      TabIndex        =   21
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   1845
      TabIndex        =   20
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   2550
      TabIndex        =   19
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   3315
      TabIndex        =   18
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   4050
      TabIndex        =   17
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   4830
      TabIndex        =   16
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   5580
      TabIndex        =   15
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   6315
      TabIndex        =   14
      Top             =   3690
      Width           =   210
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   7080
      TabIndex        =   13
      Top             =   3690
      Width           =   225
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H8000000B&
      Height          =   210
      Left            =   7830
      TabIndex        =   12
      Top             =   3690
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "EXIT"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4050
      Width           =   615
   End
   Begin VB.CommandButton ftrack 
      Caption         =   "Skip Forward"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1020
      Width           =   1245
   End
   Begin VB.CommandButton eject0 
      Caption         =   "Eject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1020
      Width           =   720
   End
   Begin VB.CommandButton btrack 
      Caption         =   "Skip Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   930
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1020
      Width           =   1245
   End
   Begin VB.CommandButton ff 
      Caption         =   "Fast Forward"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1260
   End
   Begin VB.CommandButton pause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   810
   End
   Begin VB.CommandButton play 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   1260
   End
   Begin VB.CommandButton stopbtn 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1020
      Width           =   845
   End
   Begin VB.CommandButton rew 
      Caption         =   "Rewind"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1020
      Width           =   1260
   End
   Begin VB.TextBox CD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "CD Player"
      Top             =   1410
      Width           =   3840
   End
   Begin ComctlLib.Slider sliderMasterVolume 
      Height          =   1230
      Left            =   90
      TabIndex        =   23
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   1230
      Left            =   840
      TabIndex        =   25
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   1230
      Left            =   2325
      TabIndex        =   26
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider9 
      Height          =   1230
      Left            =   1590
      TabIndex        =   31
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider6 
      Height          =   1365
      Left            =   7590
      TabIndex        =   32
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2408
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   1
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin VB.CommandButton eject1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1020
      Width           =   720
   End
   Begin ComctlLib.Slider trebleslider 
      Height          =   1380
      Left            =   8310
      TabIndex        =   90
      Top             =   2055
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2434
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider BassSlider 
      Height          =   1380
      Left            =   9030
      TabIndex        =   93
      Top             =   2070
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   2434
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin MSComctlLib.Slider Slider13 
      Height          =   225
      Left            =   3090
      TabIndex        =   100
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider14 
      Height          =   225
      Left            =   3825
      TabIndex        =   101
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   1230
      Left            =   3090
      TabIndex        =   27
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider sliderWaveOutVolume 
      Height          =   1230
      Left            =   3825
      TabIndex        =   24
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin MSComctlLib.Slider Slider15 
      Height          =   225
      Left            =   4590
      TabIndex        =   102
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider16 
      Height          =   225
      Left            =   5325
      TabIndex        =   103
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Slider5 
      Height          =   1230
      Left            =   5325
      TabIndex        =   29
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider4 
      Height          =   1230
      Left            =   4590
      TabIndex        =   28
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin MSComctlLib.Slider Slider17 
      Height          =   225
      Left            =   6090
      TabIndex        =   104
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider18 
      Height          =   225
      Left            =   6825
      TabIndex        =   105
      Top             =   3195
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Min             =   -10
      TickStyle       =   3
   End
   Begin ComctlLib.Slider Slider7 
      Height          =   1230
      Left            =   6825
      TabIndex        =   33
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin ComctlLib.Slider Slider8 
      Height          =   1230
      Left            =   6090
      TabIndex        =   30
      Top             =   2070
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   2170
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   200
      Max             =   65535
      SelStart        =   32768
      TickStyle       =   2
      TickFrequency   =   6535
      Value           =   32768
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8580
      Top             =   330
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "MIXER"
      Height          =   330
      Left            =   5550
      Picture         =   "Form4.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2130
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   510
      X2              =   8040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Unmute"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8190
      TabIndex        =   106
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label tracktime 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   6960
      TabIndex        =   56
      Top             =   1410
      Width           =   2955
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Treble"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8415
      TabIndex        =   95
      Top             =   3465
      Width           =   465
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Bass"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9150
      TabIndex        =   94
      Top             =   3465
      Width           =   375
   End
   Begin VB.Line Li2 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   8040
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "SBM"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8190
      TabIndex        =   59
      Top             =   3690
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Mute"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8220
      TabIndex        =   58
      Top             =   4185
      Width           =   375
   End
   Begin VB.Line Li1 
      BorderColor     =   &H00000000&
      X1              =   450
      X2              =   8190
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Label totalplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   30
      TabIndex        =   55
      Top             =   1410
      Width           =   3030
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Pc spk"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   44
      Top             =   3435
      Width           =   525
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Line in"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   43
      Top             =   3435
      Width           =   540
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "I25in"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6225
      TabIndex        =   42
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "MIDI"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5490
      TabIndex        =   41
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "TAD"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4740
      TabIndex        =   40
      Top             =   3435
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Aux"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3270
      TabIndex        =   39
      Top             =   3435
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Mic"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   38
      Top             =   3435
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "CD"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1050
      TabIndex        =   37
      Top             =   3435
      Width           =   255
   End
   Begin VB.Label lblMasterVolume 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Master"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   195
      TabIndex        =   36
      Top             =   3435
      Width           =   495
   End
   Begin VB.Label lblWaveOutVolume 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Wave"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3870
      TabIndex        =   35
      Top             =   3435
      Width           =   540
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "SBM"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7725
      TabIndex        =   34
      Top             =   3435
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim volume As Long
Dim Llevelmeter As Long
Dim Rlevelmeter As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim hmxobj As Long
Dim hmixer As Long             ' mixer handle
Dim VolCtrl As MIXERCONTROL    ' master volume control
Dim WavCtrl As MIXERCONTROL    ' wave output volume control
Dim CDVol As MIXERCONTROL      ' CD Volume
Dim LineVol As MIXERCONTROL    ' Line/In Volume
Dim MBOOST As MIXERCONTROL     ' Microphone Volume
Dim PSPKVol As MIXERCONTROL    ' PcSpeaker Volume
Dim PCSPEAKER As MIXERCONTROL
Dim AUXVol As MIXERCONTROL     ' Auxillary Volume
Dim TADVol As MIXERCONTROL     ' TAD-In Volume
Dim WaveDSVol As MIXERCONTROL  ' Wave/Direct Sound Volume
Dim MIDIVol As MIXERCONTROL    ' Midi Volume
Dim CDSPDIF As MIXERCONTROL ' Cd Digital Volume
Dim GPSPDIFVol As MIXERCONTROL ' SPDIF Volume
Dim I25InVol As MIXERCONTROL   ' I25In Volume
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL
Dim peakmetera As MIXERCONTROL
Dim rc As Long                 ' return code
Dim ok As Boolean              ' boolean return code
Dim SLIDER As MIXERCONTROL
Dim PAN As MIXERCONTROL
Dim PanC As Long
Dim MasterMuteId(0 To 1) As Long
Dim fastForwardSpeed As Long    ' seconds to seek for ff/rew
Dim fPlaying As Boolean         ' true if CD is currently playing
Dim fCDLoaded As Boolean        ' true if CD is the the player
Dim numTracks As Integer        ' number of tracks on audio CD
Dim trackLength() As String     ' array containing length of each track
Dim track As Integer            ' current track
Dim min As Integer              ' current minute on track
Dim SEC As Integer              ' current second on track
Dim cmd As String               ' string to hold mci command strings
Dim volCtrl2 As MIXERCONTROL ' vu control play
Dim volCtrl1 As MIXERCONTROL ' vu control rec

' Send a MCI command string
' If fShowError is true, display a message box on error
Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    SendMCIString "close all", False
    cmd = "close all"
    SendMCIString cmd, True
    Unload Form4
End If
SendMCIString = (rc = 0)
End Function





Private Sub Check11_Click()
If Check11.Value = 1 Then
Check1.Value = 1
Check2.Value = 1
Check3.Value = 1
Check4.Value = 1
Check5.Value = 1
Check6.Value = 1
Check7.Value = 1
Check8.Value = 1
Check9.Value = 1
Check10.Value = 1
End If
If Check11.Value = 0 Then
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
End If
End Sub



Private Sub Command1_Click()
SendMCIString "close all", False
cmd = "close all"
SendMCIString cmd, True
'Index.Enabled = True
Unload Form4
End Sub

Private Sub Command2_Click()
    'Open the mixer with deviceID 0.
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer please check if a audio mixer is installed then retry."
        Exit Sub
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, VolCtrl)
        If volume <> -1 Then
            txtMasterVolume.Text = volume \ 6553
            sliderMasterVolume.Value = 65535 - volume
        End If
    End If
   
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, WavCtrl)
        If volume <> -1 Then
            txtWaveOutVolume.Text = volume \ 6553
            sliderWaveOutVolume.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MBOOST)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MBOOST)
        If volume <> -1 Then
            Text2.Text = volume \ 6553
            Slider2.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, CDVol)
        If volume <> -1 Then
            Text1.Text = volume \ 6553
            Slider1.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, AUXVol)
        If volume <> -1 Then
            Text3.Text = volume \ 6553
            Slider3.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, TADVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, TADVol)
        If volume <> -1 Then
            Text4.Text = volume \ 6553
            Slider4.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MIDIVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MIDIVol)
        If volume <> -1 Then
            Text5.Text = volume \ 6553
            Slider5.Value = 65535 - volume
        End If
    End If

        ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, PSPKVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, PSPKVol)
        If volume <> -1 Then
            Text7.Text = volume \ 6553
            Slider7.Value = 65535 - volume
        End If
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, I25InVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, I25InVol)
        If volume <> -1 Then
            Text8.Text = volume \ 6553
            Slider8.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, LineVol)
        If volume <> -1 Then
            Text9.Text = volume \ 6553
            Slider9.Value = 65535 - volume
        End If
    End If
End Sub



Private Sub sldPan_Scroll()
Dim LefVol As Long
Dim righvol As Long

LefVol = HIGHEST_VOLUME_SETTING - lUnsigned(sldPan.Max)
righvol = HIGHEST_VOLUME_SETTING * lUnsigned(sldPan.min)

    ok = lGetVolume(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, LefVol)
    ok = lGetVolume(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, righvol)
    PanC = 65535 - CLng(sldPan.Value)
    
    lSetVolume(hmixer, LineVol, LefVol) = LefVol
    lSetVolume(hmixer, LineVol, righvol) = righvol
End Sub

Private Sub Form_Load()
'Jamie P micracom2@hotmail.com
'This is the third part of a larger program(Parts one and two are on the planet somewhere)
'that still needs a lot of work
'I have managed to get the bass and treble
'as well as many other sliders from my
'SB Live mixer moving with this program
'this is the result of many hours work
'as I am still learning about the Windows API.
'If someone can make the balance control work
'I would like to know how you did it.
'Happy computng and ps. I am still trying
'to make a peak meter work without
'the need of a peak meter on a sound card
'suggestions are very welcome.


If (App.PrevInstance = True) Then
    End
End If

' Initialize variables
Timer1.Enabled = False
fastForwardSpeed = 5
fCDLoaded = False

' If the cd is being used, then quit
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    timeWindow.Text = "Cd in use"
    End
End If
SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True
Command2_Click
Timer4.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Index.Enabled = True
SendMCIString "close all", False
End Sub



Private Sub Option1_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option10_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option17_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option18_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option19_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option20_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option9_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option11_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option12_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option13_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option14_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option15_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option16_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option2_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option3_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option4_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
unSetMuteControl hmixer, mute, 1
End Sub

Private Sub Option5_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option6_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option7_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option8_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Function Errora()
    MsgBox "Your sound card does not support a bass control"
End Function


Private Sub Timer1_Timer()
Update
End Sub

Private Sub Timer4_Timer()

Dim a As Long
Dim b As Long
Dim d As Long
Dim e As Long
Dim f As Long

a = Abs(GetVuControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER, volCtrl2))
a = Abs(getrecvucontrol(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER, volCtrl1))

b = Abs(Lvu)
d = Abs(Rvu)
e = Abs(Lrecvu)
f = Abs(Rrecvu)

If b > Llevelmeter Then
Llevelmeter = b
sposta_Llevel (Llevelmeter)
End If

If e > Llevelmeter Then
Llevelmeter = e
sposta_Llevel (Llevelmeter)
End If

If d > Rlevelmeter Then
Rlevelmeter = d
sposta_Rlevel (Rlevelmeter)
End If

If f > Rlevelmeter Then
Rlevelmeter = f
sposta_Rlevel (Rlevelmeter)
End If

If Rlevelmeter > 0 Then Rlevelmeter = Rlevelmeter - 1000
If Llevelmeter > 0 Then Llevelmeter = Llevelmeter - 1000
If Rlevelmeter < 0 Then Rlevelmeter = 0
If Llevelmeter < 0 Then Llevelmeter = 0
sposta_Rlevel (Rlevelmeter)
sposta_Llevel (Llevelmeter)
End Sub

Private Sub trebleslider_Scroll()

      ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
    If ok = False Then
    Errora
    Exit Sub
    End If
    volume = 65535 - CLng(trebleslider.Value)
    Treblesliderte.Text = volume \ 6553
    SetVolumeControl hmixer, Treble, volume
End Sub

Private Sub BassSlider_Scroll()

      ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_BASS, Bass)
    If ok = False Then
    Errora
    Exit Sub
    End If
    volume = 65535 - CLng(BassSlider.Value)
    BassText.Text = volume \ 6553
    SetVolumeControl hmixer, Bass, volume
End Sub


' Play the CD
Private Sub Play_Click()
SendMCIString "play cd", True
fPlaying = True
CD.Text = "Playing"
End Sub

Private Sub Slider6_scroll()
Dim link As Long
link = 65535 - CLng(Slider6.Value)
If Check2.Value = 1 Then
Slider1.Value = Slider6.Value
Text1.Text = link \ 6553
SetVolumeControl hmixer, CDVol, link
End If
If Check4.Value = 1 Then
Slider2.Value = Slider6.Value
Text2.Text = link \ 6553
SetVolumeControl hmixer, MBOOST, link
End If
If Check5.Value = 1 Then
Slider3.Value = Slider6.Value
Text3.Text = link \ 6553
SetVolumeControl hmixer, AUXVol, link
End If
If Check7.Value = 1 Then
Slider4.Value = Slider6.Value
Text4.Text = link \ 6553
SetVolumeControl hmixer, TADVol, link
End If
If Check8.Value = 1 Then
Slider5.Value = Slider6.Value
Text5.Text = link \ 6553
SetVolumeControl hmixer, MIDIVol, link
End If
If Check10.Value = 1 Then
Slider7.Value = Slider6.Value
Text7.Text = link \ 6553
SetVolumeControl hmixer, PSPKVol, link
End If
If Check9.Value = 1 Then
Slider8.Value = Slider6.Value
Text8.Text = link \ 6553
SetVolumeControl hmixer, I25InVol, link
End If
If Check3.Value = 1 Then
Slider9.Value = Slider6.Value
Text9.Text = link \ 6553
SetVolumeControl hmixer, LineVol, link
End If
If Check1.Value = 1 Then
sliderMasterVolume.Value = Slider6.Value
txtMasterVolume.Text = link \ 6553
SetVolumeControl hmixer, VolCtrl, link
End If
If Check6.Value = 1 Then
sliderWaveOutVolume.Value = Slider6.Value
txtWaveOutVolume.Text = link \ 6553
SetVolumeControl hmixer, WavCtrl, link

End If

End Sub

' Stop the CD play
Private Sub stopbtn_Click()
SendMCIString "stop cd wait", True
cmd = "seek cd to " & track
SendMCIString cmd, True
fPlaying = False
CD.Text = "Stopped"
Update
End Sub
' Pause the CD
Private Sub pause_Click()
SendMCIString "pause cd", True
fPlaying = False
CD.Text = "Cd Paused"
Update
End Sub

' Eject the CD
Private Sub Eject0_Click()
SendMCIString "set cd door open", True
CD.Text = "Insert CD"
eject1.Visible = True
eject0.Visible = False
Update
End Sub
Private Sub Eject1_Click()
CD.Text = "Please wait"
SendMCIString "set cd door closed", True
eject0.Visible = True
eject1.Visible = False
Update
End Sub
' Fast forward
Private Sub ff_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) + fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) + fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Rewind the CD
Private Sub rew_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) - fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) - fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
' Forward track
Private Sub ftrack_Click()
If (track < numTracks) Then
    If (fPlaying) Then
        cmd = "play cd from " & track + 1
        SendMCIString cmd, True
    Else
        cmd = "seek cd to " & track + 1
        SendMCIString cmd, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub
' Go to previous track
Private Sub btrack_Click()
Dim from As String
If (min = 0 And SEC = 0) Then
    If (track > 1) Then
        from = CStr(track - 1)
    Else
        from = CStr(numTracks)
    End If
Else
    from = CStr(track)
End If
If (fPlaying) Then
    cmd = "play cd from " & from
    SendMCIString cmd, True
Else
    cmd = "seek cd to " & from
    SendMCIString cmd, True
End If
Update
End Sub
' Update the display and state variables
Private Sub Update()
Static s As String * 30

' Check if CD is in the player
mciSendString "status cd media present", s, Len(s), 0
If (CBool(s)) Then
    ' Enable all the controls, get CD information
    If (fCDLoaded = False) Then
        mciSendString "status cd number of tracks wait", s, Len(s), 0
        numTracks = CInt(Mid$(s, 1, 2))
        eject0.Visible = True
        eject1.Visible = False
        CD.Text = "Cd Ready"
        ' If CD only has 1 track, then it's probably a data CD
        If (numTracks = 1) Then
            CD.Text = "Not audio"
            Exit Sub
        End If
        
        mciSendString "status cd length wait", s, Len(s), 0
        totalplay.Caption = "Tracks: " & numTracks & "  Total time: " & s
        ReDim trackLength(1 To numTracks)
        Dim I As Integer
        For I = 1 To numTracks
            cmd = "status cd length track " & I
            mciSendString cmd, s, Len(s), 0
            trackLength(I) = s
        Next
        timeWindow.FontSize = 18
        play.Enabled = True
        pause.Enabled = True
        ff.Enabled = True
        rew.Enabled = True
        ftrack.Enabled = True
        btrack.Enabled = True
        stopbtn.Enabled = True
        fCDLoaded = True
        SendMCIString "seek cd to 1", True
    End If
    ' Update the track time display
    mciSendString "status cd position", s, Len(s), 0
    track = CInt(Mid$(s, 1, 2))
    min = CInt(Mid$(s, 4, 2))
    SEC = CInt(Mid$(s, 7, 2))
    timeWindow.Text = "[" & Format(track, "00") & "] " & Format(min, "00") _
            & ":" & Format(SEC, "00")
    tracktime.Caption = "Track time: " & trackLength(track)
    ' Check if CD is playing
    mciSendString "status cd mode", s, Len(s), 0
    fPlaying = (Mid$(s, 1, 7) = "playing")
Else
    ' Disable all the controls, clear the display
    If (fCDLoaded = True) Then
        play.Enabled = False
        pause.Enabled = False
        ff.Enabled = False
        rew.Enabled = False
        ftrack.Enabled = False
        btrack.Enabled = False
        stopbtn.Enabled = False
        fCDLoaded = False
        fPlaying = False
        totalplay.Caption = ""
        tracktime.Caption = ""
        CD.Text = "No CD"
    End If
End If
End Sub
' Set the fast-forward speed
Private Sub ffspeed_Click()
Dim s As String
s = InputBox("Enter the new speed in seconds", "Fast Forward Speed", CStr(fastForwardSpeed))
If IsNumeric(s) Then
    fastForwardSpeed = CLng(s)
End If
End Sub

Private Sub Slider1_Scroll()

    volume = 65535 - CLng(Slider1.Value)
    Text1.Text = volume \ 6553
    SetVolumeControl hmixer, CDVol, volume
End Sub

Private Sub Slider2_Scroll()

    volume = 65535 - CLng(Slider2.Value)
    Text2.Text = volume \ 6553
    SetVolumeControl hmixer, MBOOST, volume
End Sub

Private Sub Slider3_Scroll()

    volume = 65535 - CLng(Slider3.Value)
    Text3.Text = volume \ 6553
    SetVolumeControl hmixer, AUXVol, volume
End Sub

Private Sub Slider4_Scroll()

    volume = 65535 - CLng(Slider4.Value)
    Text4.Text = volume \ 6553
    SetVolumeControl hmixer, TADVol, volume
End Sub


Private Sub Slider5_Scroll()

    volume = 65535 - CLng(Slider5.Value)
    Text5.Text = volume \ 6553
    SetVolumeControl hmixer, MIDIVol, volume
End Sub

Private Sub Slider7_Scroll()

    volume = 65535 - CLng(Slider7.Value)
    Text7.Text = volume \ 6553
    SetVolumeControl hmixer, PSPKVol, volume
End Sub

Private Sub Slider8_Scroll()

    volume = 65535 - CLng(Slider8.Value)
    Text8.Text = volume \ 6553
    SetVolumeControl hmixer, I25InVol, volume
End Sub

Private Sub Slider9_Scroll()

    volume = 65535 - CLng(Slider9.Value)
    Text9.Text = volume \ 6553
    SetVolumeControl hmixer, LineVol, volume
End Sub

Private Sub sliderMasterVolume_Scroll()

    volume = 65535 - CLng(sliderMasterVolume.Value)
    txtMasterVolume.Text = volume \ 6553
    SetVolumeControl hmixer, VolCtrl, volume
End Sub

Private Sub sliderWaveOutVolume_Scroll()

    volume = 65535 - CLng(sliderWaveOutVolume.Value)
    txtWaveOutVolume.Text = volume \ 6553
    SetVolumeControl hmixer, WavCtrl, volume
End Sub

Public Sub sposta_Rlevel(aaa As Long)

Line2.X1 = 540 + (aaa / 65535) * 130
Line2.X2 = 180 + (aaa / 65535) * 900
If aaa < 65535 / 2 Then Line2.Y2 = 200 - (aaa / 32768) * 200
If aaa > 65535 / 2 Then Line2.Y2 = 50 + (aaa / 32768) * 150
If aaa > 60000 Then
Image2.Visible = True
Else
Image2.Visible = False
End If
End Sub


Public Sub sposta_Llevel(aaa As Long)

Line1.X1 = 540 + (aaa / 65535) * 130
Line1.X2 = 180 + (aaa / 65535) * 900
If aaa < 65535 / 2 Then Line1.Y2 = 200 - (aaa / 32768) * 200
If aaa > 65535 / 2 Then Line1.Y2 = 50 + (aaa / 32768) * 150
If aaa > 60000 Then
Image1.Visible = True
Else
Image1.Visible = False
End If
End Sub
