VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.E.C CD Player + Mixer"
   ClientHeight    =   3525
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10035
   DrawWidth       =   4
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      Caption         =   "Pc spk"
      Height          =   2325
      Left            =   8430
      TabIndex        =   38
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   87
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   73
         Top             =   1800
         Width           =   225
      End
      Begin VB.PictureBox Picture9 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   70
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option17 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option18 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider7 
         Height          =   1230
         Left            =   30
         TabIndex        =   88
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "I25in"
      Height          =   2325
      Left            =   7680
      TabIndex        =   37
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   69
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture10 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   66
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option19 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton Option20 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider8 
         Height          =   1230
         Left            =   30
         TabIndex        =   86
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "MIDI"
      Height          =   2325
      Left            =   6930
      TabIndex        =   36
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   65
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture5 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   62
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option10 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider5 
         Height          =   1230
         Left            =   30
         TabIndex        =   84
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "TAD"
      Height          =   2325
      Left            =   6180
      TabIndex        =   35
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   61
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture6 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   58
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option12 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider4 
         Height          =   1230
         Left            =   30
         TabIndex        =   82
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Wave"
      Height          =   2325
      Left            =   5430
      TabIndex        =   34
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   57
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture2 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   54
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option4 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin ComctlLib.Slider sliderWaveOutVolume 
         Height          =   1230
         Left            =   30
         TabIndex        =   80
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "AUX"
      Height          =   2325
      Left            =   4680
      TabIndex        =   33
      Top             =   1170
      Width           =   705
      Begin VB.TextBox txtWaveOutVolume 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   53
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture8 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   50
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option16 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton Option15 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   1230
         Left            =   30
         TabIndex        =   78
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mic"
      Height          =   2325
      Left            =   3930
      TabIndex        =   32
      Top             =   1170
      Width           =   705
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   49
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture4 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   46
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option8 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   1230
         Left            =   30
         TabIndex        =   76
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Line In"
      Height          =   2325
      Left            =   3180
      TabIndex        =   31
      Top             =   1170
      Width           =   705
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   45
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture3 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   42
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option6 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider Slider9 
         Height          =   1230
         Left            =   30
         TabIndex        =   41
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cd vol"
      Height          =   2325
      Left            =   2430
      TabIndex        =   24
      Top             =   1170
      Width           =   705
      Begin VB.PictureBox Picture7 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   28
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option14 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   27
         Top             =   1800
         Width           =   210
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   1230
         Left            =   30
         TabIndex        =   26
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Master"
      Height          =   2325
      Left            =   1680
      TabIndex        =   17
      Top             =   1170
      Width           =   705
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   23
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   20
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option1 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H000000FF&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox txtMasterVolume 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider sliderMasterVolume 
         Height          =   1230
         Left            =   30
         TabIndex        =   19
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider Slider19 
         Height          =   225
         Left            =   60
         TabIndex        =   99
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -32767
         Max             =   32767
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Treble"
      Height          =   2325
      Left            =   930
      TabIndex        =   14
      Top             =   1170
      Width           =   705
      Begin ComctlLib.Slider trebleslider 
         Height          =   1770
         Left            =   30
         TabIndex        =   16
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin VB.TextBox Treblesliderte 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   330
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   750
      Width           =   735
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
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   750
      Width           =   1320
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
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   750
      Width           =   1320
   End
   Begin VB.CommandButton ff 
      Caption         =   "Forward"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   855
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   750
      Width           =   1005
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
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   750
      Width           =   855
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
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   750
      Width           =   1005
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
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   750
      Width           =   870
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   390
      Top             =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   " Bass"
      Height          =   2325
      Left            =   180
      TabIndex        =   11
      Top             =   1170
      Width           =   705
      Begin VB.TextBox BassText 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   345
      End
      Begin ComctlLib.Slider BassSlider 
         Height          =   1770
         Left            =   30
         TabIndex        =   13
         Top             =   510
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "SBM"
      Height          =   2325
      Left            =   9180
      TabIndex        =   39
      Top             =   1170
      Width           =   705
      Begin VB.CheckBox Check11 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   74
         ToolTipText     =   "Links all selected faders to SBM fader"
         Top             =   1830
         Width           =   225
      End
      Begin ComctlLib.Slider Slider6 
         Height          =   1260
         Left            =   60
         TabIndex        =   89
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2223
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "SBM"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   90
         Top             =   1830
         Width           =   375
      End
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   735
      Width           =   735
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
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   735
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "MIXER"
      Height          =   330
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame14 
      Height          =   765
      Left            =   2940
      TabIndex        =   91
      Top             =   -60
      Width           =   4155
      Begin VB.TextBox timeWindow 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   570
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   150
         Width           =   3975
      End
   End
   Begin VB.Frame Frame16 
      Height          =   765
      Left            =   7140
      TabIndex        =   93
      Top             =   -60
      Width           =   2865
      Begin VB.TextBox CD 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "CD Player"
         Top             =   450
         Width           =   2745
      End
      Begin VB.Label tracktime 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   94
         Top             =   150
         Width           =   2745
      End
   End
   Begin VB.Frame Frame17 
      Height          =   765
      Left            =   30
      TabIndex        =   95
      Top             =   -60
      Width           =   2865
      Begin VB.Label totalplay 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   97
         Top             =   150
         Width           =   2745
      End
      Begin VB.Label trtime 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   96
         Top             =   450
         Width           =   2745
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim hmixer As Long             ' mixer handle
Dim VolCtrl As MIXERCONTROL    ' master volume control
Dim WavCtrl As MIXERCONTROL    ' wave output volume control
Dim CDVol As MIXERCONTROL      ' CD Volume
Dim LineVol As MIXERCONTROL    ' Line/In Volume
Dim MBOOST As MIXERCONTROL     ' Microphone Volume
Dim PSPKVol As MIXERCONTROL    ' PcSpeaker Volume

Dim AUXVol As MIXERCONTROL     ' Auxillary Volume
Dim TADVol As MIXERCONTROL     ' TAD-In Volume

Dim MIDIVol As MIXERCONTROL    ' Midi Volume

Dim I25InVol As MIXERCONTROL   ' I25In Volume
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL

Dim rc As Long                 ' return code
Dim ok As Boolean              ' boolean return code




Dim fastForwardSpeed As Long    ' seconds to seek for ff/rew
Dim fPlaying As Boolean         ' true if CD is currently playing
Dim fCDLoaded As Boolean        ' true if CD is the the player
Dim numTracks As Integer        ' number of tracks on audio CD
Dim trackLength() As String     ' array containing length of each track
Dim track As Integer            ' current track
Dim min As Integer              ' current minute on track
Dim SEC As Integer              ' current second on track
Dim cmd As String               ' string to hold mci command strings

' Send a MCI command string
' If fShowError is true, display a message box on error
Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hWnd)
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
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_BASS, Bass)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Bass)
        If volume <> -1 Then
           BassText.Text = volume \ 6553
           BassSlider.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Treble)
        If volume <> -1 Then
            Treblesliderte.Text = volume \ 6553
            trebleslider.Value = 65535 - volume
        End If
    End If
End Sub

Private Sub Form_Load()
MsgBox "If you have any comments please email me at micracom2@hotmail.com This program still needs work and if any one can make the balance work correctly I would like to know how it was done."

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

Private Sub Slider19_Scroll()
    volL = CLng(32767.5 - Slider19)
    volR = CLng(32767.5 + Slider19)
    SetPANControl hmixer, VolCtrl, volL, volR
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
        totalplay.Caption = "Tracks: " & numTracks
        trtime.Caption = "Total time: " & s
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
Private Sub Timer1_Timer()
Update
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


