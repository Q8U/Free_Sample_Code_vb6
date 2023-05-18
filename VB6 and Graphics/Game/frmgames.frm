VERSION 5.00
Begin VB.Form frmGameScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " 4000 A.D."
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "FRMGAMES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrUpdateMessageBox 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1155
      Top             =   390
   End
   Begin VB.PictureBox picNuclear 
      Height          =   315
      Left            =   1560
      Picture         =   "FRMGAMES.frx":08CA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   135
      Top             =   1035
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picPlanet5 
      Height          =   285
      Left            =   465
      Picture         =   "FRMGAMES.frx":09C4
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   134
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPlanet4 
      Height          =   285
      Left            =   45
      Picture         =   "FRMGAMES.frx":0EE6
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   133
      Top             =   465
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPlanet3 
      Height          =   285
      Left            =   735
      Picture         =   "FRMGAMES.frx":1418
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   132
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPlanet2 
      Height          =   285
      Left            =   405
      Picture         =   "FRMGAMES.frx":194A
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   131
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPlanet1 
      Height          =   270
      Left            =   45
      Picture         =   "FRMGAMES.frx":1E7C
      ScaleHeight     =   210
      ScaleWidth      =   165
      TabIndex        =   130
      Top             =   90
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picTiny 
      Height          =   510
      Left            =   -45
      Picture         =   "FRMGAMES.frx":23AE
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   129
      Top             =   1635
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame fraLanding 
      Caption         =   "Attack/Land Ships"
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
      Height          =   945
      Left            =   7635
      TabIndex        =   113
      Top             =   5910
      Width           =   1860
      Begin VB.CommandButton cmdLandShip2 
         Caption         =   "Ship 2"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   115
         Top             =   540
         Width           =   1500
      End
      Begin VB.CommandButton cmdLandShip1 
         Caption         =   "Ship 1"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   114
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.PictureBox picTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      Picture         =   "FRMGAMES.frx":2C78
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   112
      Top             =   990
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtPlayerName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   6105
      TabIndex        =   84
      Text            =   "Player 1"
      Top             =   6480
      Width           =   1485
   End
   Begin VB.TextBox txtTurnNumber 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   5190
      TabIndex        =   55
      Text            =   "Turn "
      Top             =   6480
      Width           =   885
   End
   Begin VB.CommandButton cmdEndTurn 
      Caption         =   "&Save Turn"
      Height          =   330
      Left            =   180
      TabIndex        =   45
      Top             =   5490
      Width           =   1200
   End
   Begin VB.Frame fraUpgrade 
      Caption         =   "Resource Mgmt"
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
      Height          =   5820
      Left            =   7620
      TabIndex        =   44
      Top             =   30
      Width           =   1875
      Begin VB.CommandButton cmdRepairIndustry 
         Caption         =   "Repair Industry"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   138
         Top             =   4830
         Width           =   1500
      End
      Begin VB.CommandButton cmdCleanup 
         Caption         =   "Detoxify Planet"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   127
         Top             =   5130
         Width           =   1500
      End
      Begin VB.CommandButton cmdRegenerate 
         Caption         =   "Regenerate Planet"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   126
         Top             =   5430
         Width           =   1500
      End
      Begin VB.PictureBox picUpgrade 
         BackColor       =   &H00C0C0C0&
         Height          =   660
         Index           =   7
         Left            =   1230
         Picture         =   "FRMGAMES.frx":3542
         ScaleHeight     =   600
         ScaleMode       =   0  'User
         ScaleWidth      =   495
         TabIndex        =   125
         Top             =   1890
         Width           =   555
      End
      Begin VB.PictureBox picUpgrade 
         BackColor       =   &H00C0C0C0&
         Height          =   645
         Index           =   6
         Left            =   1260
         Picture         =   "FRMGAMES.frx":3FD8
         ScaleHeight     =   585
         ScaleMode       =   0  'User
         ScaleWidth      =   465
         TabIndex        =   124
         Top             =   1215
         Width           =   525
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scanner"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   123
         Top             =   4530
         Width           =   1500
      End
      Begin VB.CommandButton cmdPlanetName 
         Height          =   270
         Left            =   105
         TabIndex        =   120
         Top             =   210
         Width           =   1665
      End
      Begin VB.Frame fraTactical 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   60
         TabIndex        =   110
         Top             =   3660
         Width           =   1755
         Begin VB.CommandButton cmdLaunchBioRocket 
            Caption         =   "BioHazard Rocket"
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   128
            Top             =   450
            Width           =   1500
         End
         Begin VB.CommandButton cmdLaunch 
            Caption         =   "&Launch Ship"
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   111
            Top             =   150
            Width           =   1500
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   270
         Left            =   450
         TabIndex        =   90
         Top             =   3375
         Width           =   1050
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   3045
         Width           =   555
      End
      Begin VB.HScrollBar hsbQuantity 
         Enabled         =   0   'False
         Height          =   240
         LargeChange     =   2
         Left            =   90
         Max             =   50
         TabIndex        =   86
         Top             =   2760
         Width           =   1425
      End
      Begin VB.PictureBox picUpgrade 
         BackColor       =   &H00000000&
         Height          =   660
         Index           =   5
         Left            =   105
         Picture         =   "FRMGAMES.frx":493A
         ScaleHeight     =   600
         ScaleWidth      =   510
         TabIndex        =   53
         Top             =   1890
         Width           =   570
      End
      Begin VB.PictureBox picUpgrade 
         BackColor       =   &H00808080&
         Height          =   645
         Index           =   4
         Left            =   690
         Picture         =   "FRMGAMES.frx":4C44
         ScaleHeight     =   585
         ScaleWidth      =   480
         TabIndex        =   52
         Top             =   1215
         Width           =   540
      End
      Begin VB.PictureBox picUpgrade 
         Height          =   645
         Index           =   3
         Left            =   105
         Picture         =   "FRMGAMES.frx":5D52
         ScaleHeight     =   585
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   1215
         Width           =   555
      End
      Begin VB.PictureBox picUpgrade 
         BackColor       =   &H00800080&
         Height          =   675
         Index           =   2
         Left            =   1260
         Picture         =   "FRMGAMES.frx":67D4
         ScaleHeight     =   615
         ScaleWidth      =   480
         TabIndex        =   50
         Top             =   525
         Width           =   540
      End
      Begin VB.PictureBox picUpgrade 
         Height          =   675
         Index           =   1
         Left            =   660
         Picture         =   "FRMGAMES.frx":78E2
         ScaleHeight     =   615
         ScaleWidth      =   525
         TabIndex        =   49
         Top             =   525
         Width           =   585
      End
      Begin VB.PictureBox picUpgrade 
         Height          =   675
         Index           =   0
         Left            =   105
         Picture         =   "FRMGAMES.frx":8330
         ScaleHeight     =   615
         ScaleWidth      =   480
         TabIndex        =   48
         Top             =   525
         Width           =   540
      End
      Begin VB.Label lblTotal 
         Caption         =   "Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   88
         Top             =   3075
         Width           =   480
      End
      Begin VB.Label lblQuantity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1530
         TabIndex        =   87
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   105
         TabIndex        =   85
         Top             =   2580
         Width           =   1590
      End
   End
   Begin VB.Frame fraEnemyWarpPath 
      Caption         =   "Player 2 Warp Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1515
      TabIndex        =   23
      Top             =   30
      Width           =   6015
      Begin VB.CommandButton cmdPreviewEnemy2 
         Caption         =   "Ship 2"
         Height          =   255
         Left            =   75
         TabIndex        =   117
         Top             =   585
         Width           =   775
      End
      Begin VB.CommandButton cmdPreviewEnemy1 
         Caption         =   "Ship 1"
         Height          =   255
         Left            =   75
         TabIndex        =   116
         Top             =   255
         Width           =   775
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000C0&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   7
         Left            =   5295
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   98
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000040C0&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   6
         Left            =   4665
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   97
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   5
         Left            =   4035
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   96
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   4
         Left            =   3405
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   95
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   3
         Left            =   2775
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   94
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   2
         Left            =   2140
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   93
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   1
         Left            =   1515
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   92
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picEnemyPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   0
         Left            =   885
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   91
         Top             =   195
         Width           =   625
      End
   End
   Begin VB.Frame fraWarpPath 
      Caption         =   "Player 1 Warp Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1545
      TabIndex        =   12
      Top             =   5445
      Width           =   6045
      Begin VB.CommandButton cmdPreviewShip2 
         Caption         =   "Ship 2"
         Height          =   255
         Left            =   75
         TabIndex        =   119
         Top             =   585
         Width           =   775
      End
      Begin VB.CommandButton cmdPreviewShip1 
         Caption         =   "Ship 1"
         Height          =   255
         Left            =   75
         TabIndex        =   118
         Top             =   255
         Width           =   775
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   7
         Left            =   5310
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   106
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   6
         Left            =   4680
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   105
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   5
         Left            =   4050
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   104
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   4
         Left            =   3420
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   103
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   3
         Left            =   2790
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   102
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   2
         Left            =   2160
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   101
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   1
         Left            =   1530
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   100
         Top             =   195
         Width           =   625
      End
      Begin VB.PictureBox picPlayerPath 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Index           =   0
         Left            =   900
         ScaleHeight     =   615
         ScaleWidth      =   570
         TabIndex        =   99
         Top             =   195
         Width           =   625
      End
   End
   Begin VB.Frame fraPlayerStats 
      Caption         =   "Player Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   5
      Top             =   2430
      Width           =   1440
      Begin VB.TextBox txtProduction 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   405
         MultiLine       =   -1  'True
         TabIndex        =   137
         Top             =   2160
         Width           =   660
      End
      Begin VB.TextBox txtNumAssaultTroops 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   675
         MultiLine       =   -1  'True
         TabIndex        =   109
         Top             =   1020
         Width           =   660
      End
      Begin VB.TextBox txtNumResources 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   405
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1590
         Width           =   660
      End
      Begin VB.TextBox txtNumTroops 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   675
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   645
         Width           =   660
      End
      Begin VB.TextBox txtNumPlanets 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   675
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Resources/Turn:"
         Height          =   210
         Left            =   105
         TabIndex        =   136
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label lblAssaultTroops 
         Caption         =   "Mechs:"
         Height          =   225
         Left            =   90
         TabIndex        =   108
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label lblResources 
         Caption         =   "Total Resources:"
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   1365
         Width           =   1260
      End
      Begin VB.Label lblNumTroops 
         Caption         =   "Troops:"
         Height          =   225
         Left            =   75
         TabIndex        =   7
         Top             =   675
         Width           =   570
      End
      Begin VB.Label lblnumplanets 
         Caption         =   "Planets:"
         Height          =   210
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Height          =   330
      Left            =   180
      TabIndex        =   4
      Top             =   5820
      Width           =   1200
   End
   Begin VB.Timer tmrRandomSounds 
      Interval        =   35000
      Left            =   495
      Top             =   6180
   End
   Begin VB.Frame fraMessages 
      Caption         =   "Messages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   45
      TabIndex        =   2
      Top             =   1005
      Width           =   1455
      Begin VB.CommandButton cmdViewSend 
         Caption         =   "&View/Send"
         Height          =   300
         Left            =   60
         TabIndex        =   54
         Top             =   975
         Width           =   1275
      End
      Begin VB.TextBox txtMessages 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "FRMGAMES.frx":8D7E
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6480
      Width           =   5070
   End
   Begin VB.PictureBox picGalaxy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   1800
      ScaleHeight     =   4035
      ScaleWidth      =   5670
      TabIndex        =   0
      Top             =   1230
      Width           =   5730
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   49
         Left            =   5220
         Picture         =   "FRMGAMES.frx":8D99
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   83
         Top             =   3510
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   48
         Left            =   4770
         Picture         =   "FRMGAMES.frx":92CB
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   82
         Top             =   3120
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   47
         Left            =   4320
         Picture         =   "FRMGAMES.frx":97FD
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   81
         Top             =   3660
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   46
         Left            =   3540
         Picture         =   "FRMGAMES.frx":9D1F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   80
         Top             =   3360
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   45
         Left            =   2820
         Picture         =   "FRMGAMES.frx":A251
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   79
         Top             =   3300
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   44
         Left            =   2160
         Picture         =   "FRMGAMES.frx":A783
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   78
         Top             =   3660
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   43
         Left            =   1620
         Picture         =   "FRMGAMES.frx":ACB5
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   77
         Top             =   3285
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   42
         Left            =   960
         Picture         =   "FRMGAMES.frx":B1E7
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   76
         Top             =   3660
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   41
         Left            =   600
         Picture         =   "FRMGAMES.frx":B709
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   75
         Top             =   3120
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   40
         Left            =   225
         Picture         =   "FRMGAMES.frx":BC3B
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   74
         Top             =   3555
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   39
         Left            =   5340
         Picture         =   "FRMGAMES.frx":C16D
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   73
         Top             =   2820
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   38
         Left            =   4920
         Picture         =   "FRMGAMES.frx":C69F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   72
         Top             =   2400
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   37
         Left            =   4230
         Picture         =   "FRMGAMES.frx":CBD1
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   71
         Top             =   2820
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   36
         Left            =   3600
         Picture         =   "FRMGAMES.frx":D103
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   70
         Top             =   2580
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   35
         Left            =   2820
         Picture         =   "FRMGAMES.frx":D625
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   69
         Top             =   2700
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   34
         Left            =   2100
         Picture         =   "FRMGAMES.frx":DB57
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   68
         Top             =   2760
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   33
         Left            =   1620
         Picture         =   "FRMGAMES.frx":E089
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   67
         Top             =   2355
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   32
         Left            =   1080
         Picture         =   "FRMGAMES.frx":E5BB
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   66
         Top             =   2880
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   31
         Left            =   120
         Picture         =   "FRMGAMES.frx":EAED
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   65
         Top             =   2760
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   30
         Left            =   900
         Picture         =   "FRMGAMES.frx":F01F
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   64
         Top             =   2160
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   29
         Left            =   5340
         Picture         =   "FRMGAMES.frx":F541
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   63
         Top             =   1740
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   28
         Left            =   4695
         Picture         =   "FRMGAMES.frx":FA63
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   62
         Top             =   1680
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   27
         Left            =   4200
         Picture         =   "FRMGAMES.frx":FF95
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   61
         Top             =   2220
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   26
         Left            =   3720
         Picture         =   "FRMGAMES.frx":104C7
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   60
         Top             =   1650
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   25
         Left            =   3240
         Picture         =   "FRMGAMES.frx":109F9
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   59
         Top             =   1980
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   24
         Left            =   2580
         Picture         =   "FRMGAMES.frx":10F2B
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   58
         Top             =   2145
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   23
         Left            =   1995
         Picture         =   "FRMGAMES.frx":1144D
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   57
         Top             =   1845
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   22
         Left            =   1350
         Picture         =   "FRMGAMES.frx":1197F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   56
         Top             =   1575
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   21
         Left            =   630
         Picture         =   "FRMGAMES.frx":11EB1
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   47
         Top             =   1545
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   20
         Left            =   180
         Picture         =   "FRMGAMES.frx":123E3
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   46
         Top             =   2100
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   19
         Left            =   5280
         Picture         =   "FRMGAMES.frx":12915
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   43
         Top             =   1080
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   18
         Left            =   4740
         Picture         =   "FRMGAMES.frx":12E47
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   42
         Top             =   780
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   17
         Left            =   4230
         Picture         =   "FRMGAMES.frx":13379
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   41
         Top             =   1215
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   16
         Left            =   3660
         Picture         =   "FRMGAMES.frx":138AB
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   40
         Top             =   960
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   2970
         Picture         =   "FRMGAMES.frx":13DCD
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   39
         Top             =   1365
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   2475
         Picture         =   "FRMGAMES.frx":142FF
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   38
         Top             =   870
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   2205
         Picture         =   "FRMGAMES.frx":14831
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   37
         Top             =   1365
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   1620
         Picture         =   "FRMGAMES.frx":14D63
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   36
         Top             =   975
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   810
         Picture         =   "FRMGAMES.frx":15295
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   35
         Top             =   900
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   165
         Picture         =   "FRMGAMES.frx":157C7
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   34
         Top             =   1125
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   5265
         Picture         =   "FRMGAMES.frx":15CE9
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   33
         Top             =   345
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   4725
         Picture         =   "FRMGAMES.frx":1621B
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   32
         Top             =   135
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   4140
         Picture         =   "FRMGAMES.frx":1673D
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   31
         Top             =   480
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3615
         Picture         =   "FRMGAMES.frx":16C6F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   30
         Top             =   105
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   3060
         Picture         =   "FRMGAMES.frx":171A1
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   29
         Top             =   525
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   2610
         Picture         =   "FRMGAMES.frx":176D3
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   28
         Top             =   180
         Width           =   210
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   1905
         Picture         =   "FRMGAMES.frx":17BF5
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   27
         Top             =   315
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   1140
         Picture         =   "FRMGAMES.frx":18127
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   26
         Top             =   360
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   570
         Picture         =   "FRMGAMES.frx":18659
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   25
         Top             =   60
         Width           =   225
      End
      Begin VB.PictureBox picPlanet 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   150
         Picture         =   "FRMGAMES.frx":18B8B
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   24
         Top             =   480
         Width           =   225
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   90
      TabIndex        =   107
      Top             =   5145
      Width           =   1410
   End
   Begin VB.Label lblTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "A.D."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   300
      TabIndex        =   122
      Top             =   450
      Width           =   945
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   510
      Left            =   225
      TabIndex        =   121
      Top             =   30
      Width           =   1050
   End
   Begin VB.Label lblE 
      Caption         =   "E"
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
      Left            =   6900
      TabIndex        =   22
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lblD 
      Caption         =   "D"
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
      Left            =   5820
      TabIndex        =   21
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lblC 
      Caption         =   "C"
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
      Left            =   4620
      TabIndex        =   20
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lblB 
      Caption         =   "B"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lblA 
      Caption         =   "A"
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
      Left            =   2340
      TabIndex        =   18
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label lblFive 
      Caption         =   "5"
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
      Left            =   1620
      TabIndex        =   17
      Top             =   4800
      Width           =   195
   End
   Begin VB.Label lblFour 
      Caption         =   "4"
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
      Left            =   1620
      TabIndex        =   16
      Top             =   3960
      Width           =   195
   End
   Begin VB.Label lblThree 
      Caption         =   "3"
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
      Left            =   1620
      TabIndex        =   15
      Top             =   3180
      Width           =   135
   End
   Begin VB.Label lblTwo 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label lblOne 
      Caption         =   "1"
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
      Left            =   1620
      TabIndex        =   13
      Top             =   1560
      Width           =   195
   End
End
Attribute VB_Name = "frmGameScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare the function keys
Const vbKeyF1 = &H70    'for help
Const vbKeyF2 = &H71    'to save the game
Const vbKeyF3 = &H72    'to quit without saving
Const vbKeyF4 = &H73

Dim CloakingChecked(2) As Boolean    'so cloaked ships are only checked once/turn
                                     'to see if they're hidden or not
Public Warp7WarningGiven As Boolean
Public Warp8WarningGiven As Boolean  'to limit the warp path warnings to once at
                                    'the beginning of the turn, and again when
                                    'the player hits cmdendturn
                                    

Public NumPlanets1 As Integer       'checks number of planets in range on warp 8
                                    'if zero, ship is destroyed
                                
Public NumPlanets2 As Integer       'same, for ship 2

Public GridLinesOn As Boolean       'whether or not the grid lines are showing

Public ContaminationWarningGiven As Boolean  'to only show contamination results msgbox once/turn


Private Sub cmdAbort_Click()
'abort the game without saving

PlaySoundEffect "Button2"

If MsgBox("Are you sure you want to quit without saving?", vbQuestion + vbYesNo, "Abort Turn") = vbYes Then
    PlaySoundEffect "Abort"
    Dim Counter
    For Counter = 1 To 100000
    Next Counter
    'deregister help file
    QuitHelp
    
    'turn off scanner
    If ScannerOn Then
        cmdScan_Click
        ScannerOn = False
        cmdScan.Enabled = False
    End If


    If TurnNumber = 1 Then
        'for some as-yet unknown reason (at least to me),
        'the program will not shut all the way down on turn 1
        'unless I use End - I know it makes no sense, but...
        End
        
    Else
       '***Alternative to using End:
   
        Dim F As Long
   
        'fade form into taskbar
        Me.WindowState = 1
   
        'count forms opened
        For F = Forms.Count - 1 To 0 Step -1
           Unload Forms(F)
        Next F

        'close any open files
        If (Forms.Count = 0) Then Close

        'set all open forms to Nothing
        Set frmGameScreen = Nothing
    End If
   
End If

End Sub

Private Sub cmdAbort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen
txtStatus.Text = "Quit without saving"

End Sub


Private Sub cmdCleanup_Click()
'detoxify planet contaminated by biorocket

Dim ResourceFlag As Boolean   'adds second line to msgbox, showing extra resources
Dim Q As Integer              'extra resources if currently < 2

'added check for money

If Planet(ActivePlanet).Contaminated And Player(Current).BioCleanupResearched Then
    UnitCost = 10
    If Player(Current).NumResources >= UnitCost Then
        PlaySoundEffect "Button2"
        If MsgBox("Detoxify this planet?", vbYesNo, "") = vbYes Then
        'deduct cost of cleanup
        Player(Current).NumResources = Player(Current).NumResources - UnitCost
        UpdatePlayerStats

            Planet(ActivePlanet).Contaminated = False
            'add 1-2 resources if planet has < 2
            If Planet(ActivePlanet).Resources < 2 Then
                'rejuvenate 1-2 resources
                Randomize
                Q = Int(Rnd * 1) + 1
                Planet(ActivePlanet).Resources = Planet(ActivePlanet).Resources + Q
            End If
            Dim Msg As String
            Msg = "Planetary environment now suitable for Humans"
            If ResourceFlag Then
                Msg = Msg + Chr(13) + "Resource Production increased to " + Str(Planet(ActivePlanet).Resources)
                ResourceFlag = False
            End If
            PlaySoundEffect "Quiet"
            MsgBox Msg, , "Detoxification complete"
        
            'restore planet picture
            Select Case Planet(ActivePlanet).BackGround
            Case 1
               picPlanet(ActivePlanet).Picture = picPlanet1.Picture
            Case 2
                picPlanet(ActivePlanet).Picture = picPlanet2.Picture
            Case 3
                picPlanet(ActivePlanet).Picture = picPlanet3.Picture
            Case 4
                picPlanet(ActivePlanet).Picture = picPlanet4.Picture
            Case 5
                picPlanet(ActivePlanet).Picture = picPlanet5.Picture
            End Select
        
        End If
    Else
        'insufficient funds
        PlaySoundEffect "Quiet"
        MsgBox "Detoxification costs 10 resource units", vbExclamation, "Insufficient Resources"
    End If
    
End If

End Sub

Private Sub cmdCleanup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Detoxify Planet Contaminated By BioHazard"

End Sub


Private Sub cmdEndTurn_Click()
'save the game
PlaySoundEffect "Button2"

If MsgBox("Save this turn?", vbYesNo + vbQuestion, "Ending Turn") = vbYes Then

'advance ships on warp path if not landed
If Player(Current).Ship(0).WarpPosition = 8 Or Player(Current).Ship(1).WarpPosition = 8 Then
  'don't end turn until player lands ship
  PlaySoundEffect "Disintegrate"
  MsgBox "Your ship must land this turn", vbExclamation, "Warp Limit"
  Exit Sub
End If
  
'confirm ending turn, then give message to player

Dim z As Integer
For z = 0 To 1
    If Player(Current).Ship(z).Launched Then
        Player(Current).Ship(z).WarpPosition = Player(Current).Ship(z).WarpPosition + 1
    End If
Next z

'set it to the next turn if player 2 is finishing up
If Current = 1 Then
    TurnNumber = TurnNumber + 1
End If
  
  
'turn off scanner
If ScannerOn Then
    cmdScan_Click
    ScannerOn = False
    cmdScan.Enabled = False
End If


'save planet and player info to a file, then end
WriteBigFile
  
'here's where zlib is used to compress the file just written
frmCompress.Show Modal
 
'frmCompress calls frmContinue for choice of main menu or quit
'if main menu, then this form is unloaded by frmContinue
   
End If

End Sub

Private Sub cmdEndTurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtStatus.ForeColor = vbGreen
  txtStatus.Text = "End Turn, Prepare Information File for Upload"
End Sub








Private Sub cmdLandShip1_Click()
'figure out which planets are available for landing

Dim Count As Integer
Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim RangeLow, RangeHigh
Dim xpos, ypos, radius


'set toggle state:
If ReadyToLand1 = True Then
    ReadyToLand1 = False
ElseIf ReadyToLand1 = False Then
    ReadyToLand1 = True
End If

If ReadyToLand1 Then
   cmdLandShip2.Enabled = False
   cmdPreviewShip1.Enabled = False
   cmdPreviewShip2.Enabled = False
   cmdPreviewEnemy1.Enabled = False
   cmdPreviewEnemy2.Enabled = False
ElseIf ReadyToLand1 = False Then
    If Player(Current).Ship(1).Launched Then
        cmdLandShip2.Enabled = True
    End If
    
    cmdPreviewShip1.Enabled = True
    cmdPreviewShip2.Enabled = True
    cmdPreviewEnemy1.Enabled = True
    cmdPreviewEnemy2.Enabled = True
    
End If

'set activeship to appropriate ship number
activeship = 0

'clear the board of any inrange settings
Dim z
For z = 0 To 49
    Planet(z).InRange = False
Next z

'Check for UltraWarp and set ranges
If Player(Current).UltraWarpResearched Then
    'increased range
    RangeLow = Player(Current).Ship(0).WarpPosition * 250
    RangeHigh = RangeLow + 700
ElseIf Player(Current).UltraWarpResearched = False Then
    'lower ranges
    RangeLow = Player(Current).Ship(0).WarpPosition * 250
    RangeHigh = RangeLow + 350
End If

'ship's starting position - originating planet
x1 = Player(Current).Ship(0).CenterX
y1 = Player(Current).Ship(0).CenterY

'check distance from home planet to each other planet
'if within the range, set planet's InRange to true
For Count = 0 To 49
    X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
    Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
    a = Abs(x1 - X2)
    b = Abs(y1 - Y2)
   
    Distance = Int(Sqr(a ^ 2 + b ^ 2))

    If Distance >= RangeLow And Distance <= RangeHigh And picPlanet(Count).Visible Then
        'planet is within range - set value
        Planet(Count).InRange = True
    End If
    
Next Count


'if in range, draw circle
For Count = 0 To 49
    If Planet(Count).InRange And picPlanet(Count).Visible Then
        'find center of the picturebox and draw circle
        xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
        radius = (picPlanet(Count).Width / 2) + 45
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbYellow
    End If
Next Count

End Sub

Private Sub cmdLandShip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Prepare Ship 1 for Attack/Landing"
End Sub


Private Sub cmdLandShip2_Click()
'figure out which planets are available for landing

Dim Count As Integer
Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim RangeLow, RangeHigh
Dim xpos, ypos, radius

'set toggle state:
If ReadyToLand2 = True Then
    ReadyToLand2 = False
ElseIf ReadyToLand2 = False Then
    ReadyToLand2 = True
End If

If ReadyToLand2 Then
    cmdLandShip1.Enabled = False
    cmdPreviewShip1.Enabled = False
    cmdPreviewShip2.Enabled = False
    cmdPreviewEnemy1.Enabled = False
    cmdPreviewEnemy2.Enabled = False
ElseIf ReadyToLand2 = False Then
    '**see if other ship launched
    If Player(Current).Ship(0).Launched Then
        cmdLandShip1.Enabled = True
    End If
    cmdPreviewShip1.Enabled = True
    cmdPreviewShip2.Enabled = True
    cmdPreviewEnemy1.Enabled = True
    cmdPreviewEnemy2.Enabled = True
End If

'set activeship to appropriate ship number
activeship = 1

'clear the board of any inrange settings
Dim z
For z = 0 To 49
    Planet(z).InRange = False
Next z

'Check for UltraWarp and set ranges
If Player(Current).UltraWarpResearched Then
    'increased range
    RangeLow = Player(Current).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 700
ElseIf Player(Current).UltraWarpResearched = False Then
    'lower ranges
    RangeLow = Player(Current).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 350
End If


'ship's starting position - originating planet
x1 = Player(Current).Ship(1).CenterX
y1 = Player(Current).Ship(1).CenterY
    
'check distance from home planet to each other planet
'if within the range, set planet's InRange to true

For Count = 0 To 49
    X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
    Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)

    a = Abs(x1 - X2)
    b = Abs(y1 - Y2)
   
    Distance = Int(Sqr(a ^ 2 + b ^ 2))
    
    If Distance >= RangeLow And Distance <= RangeHigh Then
        'planet is within range - set value
        Planet(Count).InRange = True
   End If
   
Next Count

'if in range, draw circle
For Count = 0 To 49
    If Planet(Count).InRange And picPlanet(Count).Visible Then
        'find center of the picturebox and draw circle
        xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
        radius = (picPlanet(Count).Width / 2) + 45
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbYellow
    End If
Next Count

End Sub


Private Sub cmdLandShip2_LostFocus()

'EraseCircles

End Sub


Private Sub cmdLandShip2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Prepare Ship 2 for Attack/Landing"

End Sub


Private Sub cmdLaunch_Click()
'launch a ship

PlaySoundEffect "Button4"

'check to see if a new planet - can't take off this turn
Dim msg1 As String

msg1 = "Engines being re-tooled after landing:" & Chr(13)
msg1 = msg1 + "Ship cannot launch this turn."

If Planet(ActivePlanet).JustLanded Then
    PlaySoundEffect "Quiet"
    MsgBox msg1, vbOKOnly + vbExclamation, "Operation Terminated"
    Exit Sub
ElseIf Player(Current).Ship(0).Launched And Player(Current).Ship(1).Launched Then
    PlaySoundEffect "Quiet"
    MsgBox "No ships available for launch", vbOKOnly, "Launch Cancelled"
    Exit Sub
Else
    'load the launch form
    frmLaunch.Show Modal
End If

End Sub

Private Sub cmdLaunch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Prepare a ship for launch"

End Sub


Private Sub cmdLaunchBioRocket_Click()
'launch biorocket - check if planets in range

PlaySoundEffect "Button4"

Dim Range
Dim x1, y1, X2, Y2
Dim Count
Dim a, b
Dim Distance
Dim RocketCost

RocketCost = 30

'set toggle state:
If BioRocketOn = True Then
    BioRocketOn = False
    Planet(ActivePlanet).LaunchSite = False
ElseIf BioRocketOn = False Then
    BioRocketOn = True
    Planet(ActivePlanet).LaunchSite = True
End If


If Player(Current).BioRocketResearched And Player(Current).NumResources >= RocketCost Then
    'continue
    'set range
    If Player(Current).LongBioResearched Then
        Range = 1200
    Else
        Range = 750
    End If
    
    'clear the board of any inbiorange settings
    Dim z
    For z = 0 To 49
        Planet(z).InBioRange = False
    Next z
    
    'rocket's centre - originating planet
    x1 = picPlanet(ActivePlanet).Left + (picPlanet(ActivePlanet).Width / 2)
    y1 = picPlanet(ActivePlanet).Top + (picPlanet(ActivePlanet).Height / 2)

    'check distance from source planet to each other planet
    'if within the range and visible (ie. different galaxy sizes),
    'then set planet's InBioRange to true

    For Count = 0 To 49
       X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
       Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
       a = Abs(x1 - X2)
       b = Abs(y1 - Y2)
   
       Distance = Int(Sqr(a ^ 2 + b ^ 2))
   
       If Distance <= Range And picPlanet(Count).Visible Then
          Planet(Count).InBioRange = True
          Planet(Count).BioDistance = Distance
       End If
    Next Count

    'Disallow targeting of planets you own
    For Count = 0 To 49
        If Planet(Count).Owner = Current Then
            Planet(Count).InBioRange = False
        End If
    Next Count
    
    'if in range and visible, draw line
    For Count = 0 To 49
        If Planet(Count).InBioRange And picPlanet(Count).Visible Then
            'find center of the picturebox and draw line
            X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
            Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
            picGalaxy.DrawWidth = 2
            picGalaxy.DrawMode = 7
            picGalaxy.Line (x1, y1)-(X2, Y2), vbMagenta
        End If
    Next Count

        
Else
    'not enough money, or don't have technology
    PlaySoundEffect "Quiet"
    MsgBox "Insufficient funds - BioHazard Rockets" + Chr(13) + "cost 30 resource units each", vbOKOnly + vbInformation, "Launch Cancelled"
    Exit Sub
End If

End Sub


Private Sub cmdLaunchBioRocket_LostFocus()

'EraseLines (ActivePlanet)


    
End Sub


Private Sub cmdLaunchBioRocket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Launch BioHazard Rocket"

End Sub


Private Sub cmdOK_Click()
'put the purchase order through
'and update the player and planet stats
'uses horizontal scroll bar hsbQuantity to set # of units purchased

PlaySoundEffect "Button1"

Select Case lblItemName.Caption
Case "Missile Defences"
     'Buy missile defenses for planet
     'Cost: 10, Adds 5 to combatstrength
     'exit sub if planet already has missiles
     If Planet(ActivePlanet).HaveMissiles Then
        Exit Sub
     End If
     
     Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
     UpdatePlayerStats
     
     'add to planet's combatstrength if hsbquantity >0
     If hsbQuantity.Value > 0 Then
         Planet(ActivePlanet).HaveMissiles = True
         SetCombatStrength (ActivePlanet)
     End If
     
Case "Planetary Shield"
     'Buy shield defense for planet
     'Cost:20, Adds 25 to combatstrength

     Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
     UpdatePlayerStats
     
     'add to planet's combatstrength if hsbquantity>0
     If hsbQuantity.Value > 0 Then
        Planet(ActivePlanet).HaveShields = True
        SetCombatStrength (ActivePlanet)
     End If
       
Case "Troops"
    'update #troops for player and for planet
    Player(Current).NumTroops = Player(Current).NumTroops + Quantity
    Planet(ActivePlanet).Troops = Planet(ActivePlanet).Troops + Quantity
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    UpdatePlayerStats
    SetCombatStrength (ActivePlanet)

Case "Assault Mechs"
    'update #assault troops for player and planet
    Player(Current).NumAssaultTroops = Player(Current).NumAssaultTroops + Quantity
    Planet(ActivePlanet).AssaultTroops = Planet(ActivePlanet).AssaultTroops + Quantity
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    UpdatePlayerStats
    SetCombatStrength (ActivePlanet)
     
Case "Improved Resource Production"
     'Increase planet's resource production
     'Cost: 15, Adds random 2-5 to planet's resources
     
     'set flag that this planet has improved resources if hsbquantity>0
     If hsbQuantity.Value > 0 Then
        'add to planet's resources
         Dim Increase As Integer
         Increase = Int(Rnd * 3) + 2
         Planet(ActivePlanet).Resources = Planet(ActivePlanet).Resources + Increase
         'set improved flag
         Planet(ActivePlanet).ImprovedResources = True
         'charge the player
         Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
         UpdatePlayerStats
     End If
     
Case "Scanner"
     'Cost: 25, lets player see detailed info on surrounding planets
     Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
     UpdatePlayerStats
    
     'set flag that this planet has scanner
     '***if quantity=1
     If hsbQuantity.Value > 0 Then
         Planet(ActivePlanet).HaveScanner = True
     End If
     
Case "Scanner Jamming Device"
     'Cost: 15
     'PurchasePrice = 15
     Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
     UpdatePlayerStats
    
     'set flag that this planet has scanner jamming device
     If hsbQuantity.Value > 0 Then
         Planet(ActivePlanet).HaveJammer = True
     End If
End Select

'reset the scroll bar and labels
cmdLaunch.Enabled = False
ClearFrame

'disable landing frame only if no ships launched
If Player(Current).Ship(0).Launched = False And Player(Current).Ship(1).Launched = False Then
    fraLanding.Enabled = False
End If

End Sub

























Private Sub cmdPlanetName_Click()
'go to the landscape view

frmLandscape.Show Modal

End Sub

Private Sub cmdPlanetName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Click here to view the planet"
End Sub


Private Sub cmdPreviewEnemy1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'figure out which planets are available for enemy landing next turn

'exit sub if warp scanner not researched
If Player(Current).WarpScannerResearched = False Then
    'don't do the preview at all
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
    Exit Sub
End If

'***********************
'In form_activate, there is a base 98% chance of a cloaked ship staying hidden.
'If the current player has the warp scanner, there is an 75% chance, +/- 5%, of remaining
'hidden.  I'm not sure if that's too high...
'Either the chance of being hidden in the first place is reduced if the player
'has a warp scanner (ie. by 50%), or I have to recalculate the chance of successfully
'previewing the ship's landing sites here.
'***********************

If Player(Other).Ship(0).HaveCloakingDevice And Player(Other).Ship(0).Hidden Then
    'don't do the preview at all
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
    Exit Sub
End If

'first, see if there's a ship in the warp path
If Player(Other).Ship(0).Launched Then

Dim Count As Integer

Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim RangeLow, RangeHigh
Dim xpos, ypos, radius

'clear the board of any inrange settings
Dim z
For z = 0 To 49
    Planet(z).InRange = False
Next z

'Check for UltraWarp and set ranges
If Player(Other).UltraWarpResearched Then
    'increased range
    RangeLow = Player(Other).Ship(0).WarpPosition * 250
    RangeHigh = RangeLow + 700
ElseIf Player(Other).UltraWarpResearched = False Then
    'lower ranges
    RangeLow = Player(Other).Ship(0).WarpPosition * 250
    RangeHigh = RangeLow + 350
End If


'ship's starting position - originating planet
x1 = Player(Other).Ship(0).CenterX
y1 = Player(Other).Ship(0).CenterY

For Count = 0 To 49
    'check distance from home planet to each other planet
    'if within the range, set planet's InRange to true
    'use centre of planets, not top left corner
    X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
    Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)

    a = Abs(x1 - X2)
    b = Abs(y1 - Y2)
    Distance = Int(Sqr(a ^ 2 + b ^ 2))
    
    If Distance >= RangeLow And Distance <= RangeHigh And picPlanet(Count).Visible Then
        'planet is within range - set value
        Planet(Count).InRange = True
        'find center of the picturebox and draw circle
        xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
        radius = (picPlanet(Count).Width / 2) + 45
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbRed
    End If

Next Count

'if no ship in the warp path, error message
Else
   PlaySoundEffect "Quiet"
   MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
End If

End Sub


Private Sub cmdPreviewEnemy1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'redraw the circles to erase them

'if cloaked, don't do this routine at all
If Player(Other).Ship(0).HaveCloakingDevice And Player(Other).Ship(0).Hidden Then
    'don't do the preview at all
    Exit Sub
End If

Dim xpos, ypos, radius
Dim z
    For z = 0 To 49
    If Planet(z).InRange And picPlanet(z).Visible Then
        'find center of the picturebox and draw circle
        xpos = picPlanet(z).Left + (picPlanet(z).Width / 2)
        ypos = picPlanet(z).Top + (picPlanet(z).Height / 2)
        radius = (picPlanet(z).Width / 2) + 45
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbRed
    End If
    Next z
End Sub


Private Sub cmdPreviewEnemy2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'figure out which planets are available for enemy landing next turn

'exit sub if warp scanner not researched
If Player(Current).WarpScannerResearched = False Then
    'don't do the preview at all
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
    Exit Sub
End If

If Player(Other).Ship(1).HaveCloakingDevice And Player(Other).Ship(1).Hidden Then
    'don't do the preview at all
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
    Exit Sub
End If

If Player(Other).Ship(1).Launched Then

Dim Count As Integer

Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim RangeLow, RangeHigh
Dim xpos, ypos, radius

'clear the board of any inrange settings
Dim z
For z = 0 To 49
    Planet(z).InRange = False
Next z

'Check for UltraWarp and set ranges
If Player(Other).UltraWarpResearched Then
    'increased range
    RangeLow = Player(Other).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 700
ElseIf Player(Other).UltraWarpResearched = False Then
    'lower ranges
    RangeLow = Player(Other).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 350
End If


'ship's starting position - originating planet
x1 = Player(Other).Ship(1).CenterX
y1 = Player(Other).Ship(1).CenterY

For Count = 0 To 49
    'check distance from home planet to each other planet
    'if within the range, set planet's InRange to true

    'use centre of planets, not top left corner
    X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
    Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)

    a = Abs(x1 - X2)
    b = Abs(y1 - Y2)
    Distance = Int(Sqr(a ^ 2 + b ^ 2))
    
    If Distance >= RangeLow And Distance <= RangeHigh And picPlanet(Count).Visible Then
        'planet is within range - set value
        Planet(Count).InRange = True
        'find center of the picturebox and draw circle
        xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
        radius = (picPlanet(Count).Width / 2) + 45
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbRed
    End If

Next Count

Else
   PlaySoundEffect "Quiet"
   MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
   
End If

End Sub


Private Sub cmdPreviewEnemy2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'redraw the circles to erase them

'if cloaked, don't do this routine at all
If Player(Other).Ship(1).HaveCloakingDevice And Player(Other).Ship(1).Hidden Then
    'don't do the preview at all
    Exit Sub
End If

Dim xpos, ypos, radius
Dim z
    For z = 0 To 49
    If Planet(z).InRange And picPlanet(z).Visible Then
        'find center of the picturebox and draw circle
        xpos = picPlanet(z).Left + (picPlanet(z).Width / 2)
        ypos = picPlanet(z).Top + (picPlanet(z).Height / 2)
        radius = (picPlanet(z).Width / 2) + 45
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbRed
    End If
    Next z
End Sub


Private Sub cmdPreviewShip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'figure out which planets are available for landing

Dim NumPlanetsInRange As Integer   'if none, tell player

If Player(Current).Ship(0).Launched Then

    'figure out which planets are available for landing
    Dim Count As Integer
    Dim x1, y1, X2, Y2
    Dim a As Integer
    Dim b As Integer
    Dim Distance
    Dim RangeLow, RangeHigh
    Dim xpos, ypos, radius

    'set activeship to appropriate ship number
    activeship = 0

    'clear the board of any inrange settings
    Dim z
    For z = 0 To 49
        Planet(z).InRange = False
    Next z

'Check for UltraWarp and set ranges
    If Player(Current).UltraWarpResearched Then
        'increased range
        RangeLow = Player(Current).Ship(0).WarpPosition * 250
        RangeHigh = RangeLow + 700
    ElseIf Player(Current).UltraWarpResearched = False Then
        'lower ranges
        RangeLow = Player(Current).Ship(0).WarpPosition * 250
        RangeHigh = RangeLow + 350
    End If
  
    'ship's starting position - originating planet
    x1 = Player(Current).Ship(0).CenterX
    y1 = Player(Current).Ship(0).CenterY
    
    'check distance from home planet to each other planet
    'if within the range, set planet's InRange to true

    For Count = 0 To 49
        'use centre of planets, not top left corner
        X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
        a = Abs(x1 - X2)
        b = Abs(y1 - Y2)
   
        Distance = Int(Sqr(a ^ 2 + b ^ 2))
        If Distance >= RangeLow And Distance <= RangeHigh Then
            'planet is within range - set value
            Planet(Count).InRange = True
            NumPlanetsInRange = NumPlanetsInRange + 1
        End If
    Next Count

    If NumPlanetsInRange = 0 Then
        PlaySoundEffect "Quiet"
        MsgBox "No planets in range this turn", vbOKOnly + vbInformation, "Attempting to Land Ship 1"
        Exit Sub
    End If
    
    'if in range, draw circle
    For Count = 0 To 49
        If Planet(Count).InRange And picPlanet(Count).Visible Then
            'find center of the picturebox and draw circle
            xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
            ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
            radius = (picPlanet(Count).Width / 2) + 45
            picGalaxy.DrawMode = 7
            picGalaxy.DrawWidth = 1
            picGalaxy.Circle (xpos, ypos), radius, vbYellow
        End If
    Next Count

Else
    'ship not launched
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
End If


End Sub


Private Sub cmdPreviewShip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""
End Sub


Private Sub cmdPreviewShip1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'redraw the circles to erase them
EraseCircles


End Sub


Private Sub cmdPreviewShip2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'figure out which planets are available for landing

Dim NumPlanetsInRange As Integer   'if none, tell player

If Player(Current).Ship(1).Launched Then

'figure out which planets are available for landing

Dim Count As Integer
Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim RangeLow, RangeHigh
Dim xpos, ypos, radius

'set activeship to appropriate ship number
    activeship = 1

'clear the board of any inrange settings
    Dim z
    For z = 0 To 49
        Planet(z).InRange = False
    Next z
    
'Check for UltraWarp and set ranges
If Player(Current).UltraWarpResearched Then
    'increased range
    RangeLow = Player(Current).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 700
ElseIf Player(Current).UltraWarpResearched = False Then
    'lower ranges
    RangeLow = Player(Current).Ship(1).WarpPosition * 250
    RangeHigh = RangeLow + 350
End If

'ship's starting position - originating planet
    x1 = Player(Current).Ship(1).CenterX
    y1 = Player(Current).Ship(1).CenterY

'check distance from home planet to each other planet
'if within the range, set planet's InRange to true

        For Count = 0 To 49
            X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
            Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
            a = Abs(x1 - X2)
            b = Abs(y1 - Y2)
   
            Distance = Int(Sqr(a ^ 2 + b ^ 2))
            If Distance >= RangeLow And Distance <= RangeHigh Then
                'planet is within range - set value
                Planet(Count).InRange = True
                NumPlanetsInRange = NumPlanetsInRange + 1
            End If
        Next Count
    
    If NumPlanetsInRange = 0 Then
        PlaySoundEffect "Quiet"
        MsgBox "No planets in range this turn", vbOKOnly + vbInformation, "Attempting to Land Ship 2"
        Exit Sub
    End If
    
    
    'if in range, draw circle
    For Count = 0 To 49
        If Planet(Count).InRange And picPlanet(Count).Visible Then
            'find center of the picturebox and draw circle
            xpos = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
            ypos = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
            radius = (picPlanet(Count).Width / 2) + 45
            picGalaxy.DrawMode = 7
            picGalaxy.DrawWidth = 1
            picGalaxy.Circle (xpos, ypos), radius, vbYellow
        End If
    Next Count

Else
    PlaySoundEffect "Quiet"
    MsgBox "No landing data available", vbInformation + vbOKOnly, "Probe Error"
End If

End Sub


Private Sub cmdPreviewShip2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""

End Sub


Private Sub cmdPreviewShip2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'redraw the circles to erase them
EraseCircles

End Sub


Private Sub cmdRegenerate_Click()
'Regenerate planet after detoxification, after biorocket damage

If Planet(ActivePlanet).NukedResources And Player(Current).RegenerationResearched Then
    'check if enough money
    UnitCost = 15   'defined globally in .bas module
    
    If Player(Current).NumResources >= UnitCost Then
        'continue
        PlaySoundEffect "Button2"
        If MsgBox("Regenerate this planet?", vbYesNo, "") = vbYes Then
           'deduct cost
            Player(Current).NumResources = Player(Current).NumResources - UnitCost
            UpdatePlayerStats

            Randomize
            Dim Num
            Num = Int(Rnd * 3) + 2
            
            Planet(ActivePlanet).Resources = Planet(ActivePlanet).Resources + Num
            
            If Planet(ActivePlanet).Resources > 8 Then
                Planet(ActivePlanet).Resources = 8
            End If
            
            PlaySoundEffect "Quiet"
            MsgBox "Resource regeneration successful." + Chr(13) + "Production Capacity Now" + Str(Planet(ActivePlanet).Resources) + " Resources/Turn", , "Regeneration Complete"
            Planet(ActivePlanet).NukedResources = False
        End If
    Else
        'insufficient funds
        PlaySoundEffect "Quiet"
        MsgBox "Regeneration costs 15 resource units", vbExclamation, "Insufficient Resources"
    End If
End If



End Sub

Private Sub cmdRegenerate_LostFocus()
RegenerateOn = False

End Sub


Private Sub cmdRegenerate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "Regenerate Barren Planet's Resource Capacity"


End Sub


Private Sub cmdRepairIndustry_Click()

'start by checking for enough money - cost of 10-15
'if not enough, msgbox saying what it costs, then exit sub
'else, do the rest

Dim Msg As String  'for msgbox

If Planet(ActivePlanet).Damaged Then
    'check if enough money
    UnitCost = 10   'defined globally in .bas module

    If Player(Current).NumResources >= UnitCost Then
        'deduct cost of repairs
        Player(Current).NumResources = Player(Current).NumResources - UnitCost
        UpdatePlayerStats
        
        'add 2-3 resources
        Dim Repairs As Integer
        Randomize
        Repairs = Int(Rnd * 1) + 2
    
        'add repairs to planet's resources
        Planet(ActivePlanet).Resources = Planet(ActivePlanet).Resources + Repairs
    
        'ceiling of 8 resources
        If Planet(ActivePlanet).Resources > 8 Then
            Planet(ActivePlanet).Resources = 8
        End If
    
        If Planet(ActivePlanet).Resources > 3 Then
            'msgbox showing amt of improvement, and that further improvements need factory
            Msg = "Resource Production on " + Planet(ActivePlanet).Name + " increased by" + Str(Repairs) + " to" + Str(Planet(ActivePlanet).Resources) + " units" + Chr(13)
            Msg = Msg + "Further improvements require an Advanced Resource Production Facility"
            PlaySoundEffect "Quiet"
            MsgBox Msg, , "Industry Repair Results"
            
            'set Damaged ppty to false to prevent further repairs
            Planet(ActivePlanet).Damaged = False
            cmdRepairIndustry.Enabled = False
            
        Else
            'msgbox showing amt of improvement, and further repairs are ok
            Msg = "Resource Production on " + Planet(ActivePlanet).Name + " increased by" + Str(Repairs) + " to" + Str(Planet(ActivePlanet).Resources) + " units" + Chr(13)
            Msg = Msg + "Further repairs are possible."
            PlaySoundEffect "Quiet"
            MsgBox Msg, , "Industry Repair Results"
        End If
    Else
        PlaySoundEffect "Quiet"
        MsgBox "Repairs cost 10 resource units", vbExclamation, "Insufficient Resources"
    End If
     
End If

End Sub

Private Sub cmdScan_Click()
'figure out which planets are in scanner range
'draw red circle, allow full details in mousemove

PlaySoundEffect "Button4"

Dim Count As Integer
Dim x1, y1, X2, Y2
Dim a As Integer
Dim b As Integer
Dim Distance
Dim Range
Dim xpos, ypos, radius

'set toggle
If ScannerOn = True Then
    ScannerOn = False
ElseIf ScannerOn = False Then
    ScannerOn = True
End If

'if scanner on, turn off buttons that screw things up
If ScannerOn Then
    cmdPreviewShip1.Enabled = False
    cmdPreviewShip2.Enabled = False
    cmdPreviewEnemy1.Enabled = False
    cmdPreviewEnemy2.Enabled = False
    cmdLandShip1.Enabled = False
    cmdLandShip2.Enabled = False

ElseIf ScannerOn = False Then
    cmdPreviewShip1.Enabled = True
    cmdPreviewShip2.Enabled = True
    cmdPreviewEnemy1.Enabled = True
    cmdPreviewEnemy2.Enabled = True
    
    If Player(Current).Ship(0).Launched Then
        cmdLandShip1.Enabled = True
    End If
    
    If Player(Current).Ship(1).Launched Then
        cmdLandShip2.Enabled = True
    End If
    
End If

'clear the board of any inrange settings
Dim z
For z = 0 To 49
   Planet(z).InScannerRange = False
Next z

'set range
If Player(Current).DeepScannerResearched Then
    Range = 1800
Else
    Range = 1200
End If

'scanner's centre - originating planet
x1 = picPlanet(ActivePlanet).Left + (picPlanet(ActivePlanet).Width / 2)
y1 = picPlanet(ActivePlanet).Top + (picPlanet(ActivePlanet).Height / 2)

'check distance from home planet to each other planet
'if within the range, set planet's InRange to true

For Count = 0 To 49
   X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
   Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
   a = Abs(x1 - X2)
   b = Abs(y1 - Y2)
   
   Distance = Int(Sqr(a ^ 2 + b ^ 2))
   
   If Distance <= Range Then
   'check for jammer, or set as in range
        If Planet(Count).HaveJammer Then
            Planet(Count).InScannerRange = False
        ElseIf Planet(Count).HaveJammer = False Then
           Planet(Count).InScannerRange = True
        End If
   End If
Next Count

'draw red circle showing range of scanner
radius = Range
picGalaxy.DrawMode = 7
picGalaxy.Circle (x1, y1), radius, vbRed

End Sub

Private Sub cmdScan_LostFocus()
'**getting rid of this seems to clear up some conflict
'between the scanner and the landing buttons...
'scanner gets turned off when a foreign planet is clicked.

'cmdScan_Click

'clear the board of any inrange settings
'Dim z As Integer
'For z = 0 To 49
'   Planet(z).InScannerRange = False
'Next z

End Sub


Private Sub cmdViewSend_Click()
'play monitor sound
PlaySoundEffect "Button3"

'***see if ships being landed, if so, turn them off
If ReadyToLand1 Then
    cmdLandShip1_Click
End If

If ReadyToLand2 Then
    cmdLandShip2_Click
End If
    
'bring up the message console form
frmMessageConsole.Show Modal

'turn off scanner
If ScannerOn Then
    cmdScan_Click
    ScannerOn = False
    cmdScan.Enabled = False
End If

End Sub



Private Sub cmdViewSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "View or Send Messages"

End Sub


Private Sub Form_Activate()

'set value of cmdland1 and 2 toggles:
ReadyToLand1 = False
ReadyToLand2 = False


'put player number under owned planets
Dim Count As Integer
For Count = 0 To 49
    If Planet(Count).Owner = Current Then
        picGalaxy.CurrentX = picPlanet(Count).Left + (picPlanet(Count).Width / 2) - 25
        picGalaxy.CurrentY = picPlanet(Count).Top + picPlanet(Count).Height + 15
        picGalaxy.ForeColor = vbYellow
        picGalaxy.Print Str(Current + 1)
    ElseIf Planet(Count).Owner = Other Then
        picGalaxy.CurrentX = picPlanet(Count).Left + (picPlanet(Count).Width / 2) - 25
        picGalaxy.CurrentY = picPlanet(Count).Top + picPlanet(Count).Height + 15
        picGalaxy.ForeColor = vbRed
        picGalaxy.Print Str(Other + 1)
    End If
Next Count

'put ships on the warp path with coordinates printed below
'*****
RefreshWarpPath
'*****
Dim z     'counter
Dim j, k  'hold warp positions - easier to type & read
For z = 0 To 1
    If Player(Other).Ship(z).Launched Then
        k = Player(Other).Ship(z).WarpPosition
       
        If Player(Other).Ship(z).HaveCloakingDevice And CloakingChecked(z) = False Then
            '
            '******see if ship is hidden*********
            Randomize

            Dim ChanceOfHiding As Integer
            Dim Result
            
            ChanceOfHiding = 98 'set base chance of being hidden
            
            'chance lowered if current player has the warp scanner
            If Player(Current).WarpScannerResearched Then
                ChanceOfHiding = 75
                ChanceOfHiding = ChanceOfHiding + Int(Rnd * 5)
                ChanceOfHiding = ChanceOfHiding - Int(Rnd * 5)
            End If
            
            Result = Int(Rnd * 100) + 1
            
            If Result <= ChanceOfHiding Then
                'don't show picture on path
                Debug.Print "Chance of hiding: ", ChanceOfHiding, "greater than/equal to Result: ", Result, "therefore, NOT showing"
                Player(Other).Ship(z).Hidden = True
                CloakingChecked(z) = True
            Else
                'ship is not hidden this turn
                Player(Other).Ship(z).Hidden = False
                Debug.Print "Chance: ", ChanceOfHiding, "Result: ", Result, "NOT HIDDEN!"
                
                'show picture
                picEnemyPath(k - 1).Picture = picTemp.Picture
                'set cursor at bottom right corner
                picEnemyPath(k - 1).CurrentX = 290
                picEnemyPath(k - 1).CurrentY = 425
                picEnemyPath(k - 1).Print Player(Other).Ship(z).Coordinate
                CloakingChecked(z) = True
            End If
        Else
            'put picture on enemy path - IF NOT HIDDEN!!
            If Player(Other).Ship(z).Hidden = False Then
                picEnemyPath(k - 1).Picture = picTemp.Picture
                'set cursor at bottom right corner
                picEnemyPath(k - 1).CurrentX = 290
                picEnemyPath(k - 1).CurrentY = 425
                picEnemyPath(k - 1).Print Player(Other).Ship(z).Coordinate
            End If
        End If
        
    End If
Next z

'enable the landing frame if either ship is in warp
If Player(Current).Ship(0).Launched Or Player(Current).Ship(1).Launched Then
    fraLanding.Enabled = True
End If

'enable the landing buttons if ships are launched
If Player(Current).Ship(0).Launched Then
   cmdLandShip1.Enabled = True
End If

If Player(Current).Ship(1).Launched Then
  cmdLandShip2.Enabled = True
End If

'enable the preview buttons
cmdPreviewShip1.Enabled = True
cmdPreviewShip2.Enabled = True
cmdPreviewEnemy1.Enabled = True
cmdPreviewEnemy2.Enabled = True


'********************************
'check for contamination - set picture
Dim h
For h = 0 To 49
    If Planet(h).Contaminated Then
        picPlanet(h).Cls
        picPlanet(h).Picture = picNuclear.Picture
    End If
Next h

'Check for Contaminated Planets - damage done every turn humans on planet
Dim X As Integer
If ContaminationWarningGiven = False Then
    For X = 0 To 49
      If Planet(X).Owner = Current And Planet(X).Contaminated Then
        'figure out damage
        BioDamage (X)
        ContaminationWarningGiven = True
      End If
    Next X
End If

'unload cover form
Unload frmCover


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'trap the function keys
Select Case KeyCode
    Case vbKeyEscape
        'show quick help screen
        PlaySoundEffect "Button3"
        'call sub from Declare.bas
        ShowQuickHelp
        
    Case vbKeyF2
        'Save game and exit
        cmdEndTurn.Value = True
        
    Case vbKeyF3
        'Exit without saving
        cmdAbort.Value = True
        
    Case vbKeyF4
        'toggle grid lines
        If GridLinesOn = False Then
            DrawGridLines
            GridLinesOn = True
        ElseIf GridLinesOn = True Then
            EraseGridLines
            GridLinesOn = False
        End If
        
    Case vbKeyF5
        'toggle sound
        If SoundOn = False Then
            SoundOn = True
        ElseIf SoundOn = True Then
            SoundOn = False
        End If
        
    Case vbKeyF6
        
        PlaySoundEffect "Button3"
        If ReadyToLand1 Then
            cmdLandShip1_Click
        End If

        If ReadyToLand2 Then
            cmdLandShip2_Click
        End If
    
        'bring up the message console form
        frmConfigure.Show Modal

        'turn off scanner
        If ScannerOn Then
            cmdScan_Click
            ScannerOn = False
            cmdScan.Enabled = False
        End If

End Select

End Sub

Private Sub Form_Load()
'test
'Player(Current).NumResources = 666
'Planet(8).Owner = Current
'Planet(8).Sabotaged = True
'Planet(8).SabotageReduction = 0
'Player(Other).NumPlanets = 10
'Player(Other).WasBig = True

'refresh this variable
NumPlanetsCaptured = 0


'for messages to players in frmAnnounce
'to set a random 3rd line
Dim alternate1 As String
Dim alternate2 As String
Dim alternate3 As String

Dim choice As Integer

Me.WindowState = DefaultGameSize

    
SetAppHelp (Me.hWnd) 'register help file with win engine

Call DisableX(Me) 'disable the X button on the control box

'turn off the planets that aren't in play
'due to galaxy size
Select Case GalaxySize
Case 30   '20 planets turned off
    picPlanet(2).Visible = False
    picPlanet(5).Visible = False
    picPlanet(7).Visible = False
    picPlanet(10).Visible = False
    picPlanet(12).Visible = False
    picPlanet(14).Visible = False
    picPlanet(16).Visible = False
    picPlanet(19).Visible = False
    picPlanet(21).Visible = False
    picPlanet(22).Visible = False
    picPlanet(26).Visible = False
    picPlanet(28).Visible = False
    picPlanet(31).Visible = False
    picPlanet(33).Visible = False
    picPlanet(34).Visible = False
    picPlanet(37).Visible = False
    picPlanet(38).Visible = False
    picPlanet(43).Visible = False
    picPlanet(45).Visible = False
    picPlanet(46).Visible = False
    picPlanet(49).Visible = False
    
    'adjust location of some planets, to even them out
    picPlanet(0).Top = 600
    picPlanet(0).Left = 300
    
    picPlanet(9).Top = 400
    picPlanet(9).Left = 5200
    
    
Case 40   '10 planets turned off
    picPlanet(4).Visible = False
    picPlanet(11).Visible = False
    picPlanet(16).Visible = False
    picPlanet(20).Visible = False
    picPlanet(22).Visible = False
    picPlanet(28).Visible = False
    'picPlanet(31).Visible = False
    picPlanet(35).Visible = False
    picPlanet(37).Visible = False
    picPlanet(39).Visible = False
    picPlanet(44).Visible = False
End Select

'to trap the function keys:
KeyPreview = True

'reset justlanded to false - not necessary, but just in case...
Dim v
For v = 0 To 49
    Planet(v).JustLanded = False
Next v

DrawGalaxy

'set the form caption with current player's number
'and set player's name in name box
Personalize

'print message in message box
If IncomingMessage = "" Then
    txtMessages.FontBold = False
    txtMessages.ForeColor = vbGreen
    txtMessages.Text = "No Messages at this time"
Else
    txtMessages.ForeColor = vbYellow
    txtMessages.FontBold = True
    txtMessages.Text = "Incoming Message..."
    frmMessageConsole.txtMessageBox.Text = IncomingMessage
    'Beep
End If

'fill in turn number box
txtTurnNumber.Text = "Turn " & Str(TurnNumber)

'******BioRocket*******
'** I moved the part that checks for contaminated planets & calls BioDamage procedure
'** to the form_activate event - seems to work, showing msgbox on top of gamescreen
'** instead of on space background...with flag to only show it once/turn

Dim X As Integer
'see if biohazard rockets landed
For X = 0 To 49
    If Planet(X).BioRocketETA = TurnNumber Then
        Detonation (X)
    End If
Next X

'********End BioRocket*********


'******Aliens********
CheckForAliens
'update troops on alien planets
UpdateAliens
'check for expansion to neutral planets
AlienExpansion
'********************

'Update players resources
UpdateResources

'fill in planet, troop and resources text boxes
UpdatePlayerStats

'initialize the colours for the status text box
Dim yellow, red, Green
yellow = &HFFFF&    'for player 1
red = &HFF&         'for player 2
Green = &HFF00&     'for neutral
'

'enable the landing frame if either ship is in warp
If Player(Current).Ship(0).Launched Or Player(Current).Ship(1).Launched Then
    fraLanding.Enabled = True
End If

'enable the landing buttons if ships are launched
If Player(Current).Ship(0).Launched Then
   cmdLandShip1.Enabled = True
End If

If Player(Current).Ship(1).Launched Then
  cmdLandShip2.Enabled = True
End If

'********************************************
'***********MULTI-TURN RESEARCH**************
If Player(Current).MechResearchDone = TurnNumber Then
    'MsgBox "Turnnumber=" + Str(TurnNumber) + "player(current).mechresearchdone=" + Str(Player(Current).MechResearchDone)
    
    TechLevel = 1   'techlevel set in module declare.bas
    Player(Current).MechResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).ShieldResearchDone = TurnNumber Then
    TechLevel = 2
    Player(Current).ShieldResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).ResourceResearchDone = TurnNumber Then
    TechLevel = 3
    Player(Current).ResourcesResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).ScannerResearchDone = TurnNumber Then
    TechLevel = 5
    Player(Current).ScannerResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).BigShipResearchDone = TurnNumber Then
    TechLevel = 6
    Player(Current).BigShipResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).UltraWarpResearchDone = TurnNumber Then
    TechLevel = 7
    Player(Current).UltraWarpResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).DeepScannerResearchDone = TurnNumber Then
    TechLevel = 8
    Player(Current).DeepScannerResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).JammerResearchDone = TurnNumber Then
    TechLevel = 9
    Player(Current).JammerResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).CloakingResearchDone = TurnNumber Then
    TechLevel = 10
    Player(Current).CloakingResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).Missile1ResearchDone = TurnNumber Then
    TechLevel = 11
    Player(Current).Missile1Researched = True
    frmTechDone.Show Modal
End If

If Player(Current).Missile2ResearchDone = TurnNumber Then
    TechLevel = 12
    Player(Current).Missile2Researched = True
    frmTechDone.Show Modal
End If

If Player(Current).LaserResearchDone = TurnNumber Then
    TechLevel = 13
    Player(Current).LaserResearched = True
    'recalculate all player's planet's combatstrengths
    'but not the troops on ships...
    Dim Q
    For Q = 0 To 49
        If Planet(Q).Owner = Current Then
            SetCombatStrength (Q)
        End If
    Next Q
    
    frmTechDone.Show Modal
End If

If Player(Current).PlasmaResearchDone = TurnNumber Then
    TechLevel = 14
    Player(Current).PlasmaResearched = True
    'reset CS for all troops - except those on ships
    For Q = 0 To 49
        If Planet(Q).Owner = Current Then
            SetCombatStrength (Q)
        End If
    Next Q
    frmTechDone.Show Modal
End If

If Player(Current).BioRocketResearchDone = TurnNumber Then
    TechLevel = 15
    Player(Current).BioRocketResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).LongBioResearchDone = TurnNumber Then
    TechLevel = 16
    Player(Current).LongBioResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).ShipShield1ResearchDone = TurnNumber Then
    TechLevel = 17
    Player(Current).ShipShield1Researched = True
    frmTechDone.Show Modal
End If

If Player(Current).ShipShield2ResearchDone = TurnNumber Then
    TechLevel = 18
    Player(Current).ShipShield2Researched = True
    frmTechDone.Show Modal
End If

If Player(Current).BioCleanupResearchDone = TurnNumber Then
    TechLevel = 19
    Player(Current).BioCleanupResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).RegenerationResearchDone = TurnNumber Then
    TechLevel = 20
    Player(Current).RegenerationResearched = True
    frmTechDone.Show Modal
End If

If Player(Current).WarpScannerResearchDone = TurnNumber Then
    TechLevel = 21
    Player(Current).WarpScannerResearched = True
    frmTechDone.Show Modal
End If

'**********************************


'******ANNOUNCEMENTS***********
'warn player when enemy has more than 10 planets
If Player(Other).NumPlanets > 9 And Player(Current).Message1Given = False Then
    MessageType = "Expanding"
    Announceline1 = "Reports Show " + Player(Other).Name + "'s Reach"
    Announceline2 = "Has Extended Across" + Str(Player(Other).NumPlanets) + " Systems."
    Announceline3 = ""
    frmAnnounce.Show Modal
    
    Player(Current).Message1Given = True
End If

'second warning when more than 20 planets
If Player(Other).NumPlanets > 19 And Player(Current).Message2Given = False Then
    MessageType = "Expanding"
    Announceline1 = Player(Other).Name + "'s Empire Now Spans" + Str(Player(Other).NumPlanets) + " Systems."
    Announceline2 = "Your People Demand Action!"
    Announceline3 = ""
    frmAnnounce.Show Modal

    Player(Current).Message2Given = True
    Player(Other).WasBig = True
End If

'message if enemy's empire is shrinking
If Player(Other).NumPlanets < 12 And Player(Other).WasBig Then
    MessageType = "Expanding"
    Announceline1 = Player(Other).Name + "'s Empire is Crumbling -"
    Announceline2 = "Fewer Than 15 Systems Remain..."
    Announceline3 = "Victory Is At Hand!"
    frmAnnounce.Show Modal

    Player(Other).WasBig = False
End If

'message if planet captured by opponent
For X = 0 To 49
    If Planet(X).Owner = Other And Planet(X).Captured = True Then
        'tell player which planet is lost
        MessageType = "Captured"
        Announceline1 = Player(Other).Name + " has triumphed over our forces"
        Announceline2 = "stationed on " + Planet(X).Name + " ---"
        
        '***set a random third line
        Randomize
        choice = Int(Rnd * 3) + 1
                    
        alternate1 = "We must avenge them!"
        alternate2 = "This humiliation cannot be accepted..."
        alternate3 = "We must reclaim this world for the Empire!"
        
        Select Case choice
           Case 1
               Announceline3 = alternate1
           Case 2
               Announceline3 = alternate2
           Case 3
               Announceline3 = alternate3
        End Select
       
        frmAnnounce.Show Modal
        
        'reset captured variable
        Planet(X).Captured = False
     End If
Next X

'message that other player's invasion failed
For X = 0 To 49
    If Planet(X).Owner = Current And Planet(X).FailedInvasion = True Then
    
        MessageType = "Failed Invasion"
        
        '***set a random FIRST line
        Randomize
        choice = Int(Rnd * 3) + 1
                    
        alternate1 = "Our Forces on " + Planet(X).Name + " have stopped an invasion!"
        alternate2 = Planet(X).Name + " reports a failed invasion!"
        alternate3 = Player(Other).Name + " was turned back at " + Planet(X).Name + "!"
        
        Select Case choice
           Case 1
               Announceline1 = alternate1
           Case 2
               Announceline1 = alternate2
           Case 3
               Announceline1 = alternate3
        End Select

        'Announceline1 = "Our Forces on " + Planet(X).Name + " have repelled an invasion!"
        Announceline2 = "Troop Losses: " + Str(Planet(X).FailedInvasionTroopLosses)
        
        If Player(Current).MechResearched = True Then
            Announceline3 = "Mech Losses: " + Str(Planet(X).FailedInvasionMechLosses)
        Else
            Announceline3 = ""
        End If
        
        frmAnnounce.Show Modal

               
        'reset values to prevent repeat messages
        Planet(X).FailedInvasion = False
        Planet(X).FailedInvasionTroopLosses = 0
        Planet(X).FailedInvasionMechLosses = 0
        
    End If
Next X




'message if biorocket did not detonate - warn player that enemy tried
'to use biorocket on them...
For X = 0 To 49
    If Planet(X).Owner = Current And Planet(X).BioFailed Then
        'warn player that enemy tried to use biorocket
        MessageType = "BioFailed"
        Announceline1 = "Leaders on " + Planet(X).Name + " confirm the destruction"
        Announceline2 = "of an incoming BioChemical Rocket"
        
        '***set a random third line
        Randomize
        choice = Int(Rnd * 3) + 1
                    
        alternate1 = "A catastrophe has been narrowly avoided!"
        alternate2 = "A swift reprisal is demanded!"
        alternate3 = "The planet is saved!"
        
        Select Case choice
           Case 1
               Announceline3 = alternate1
           Case 2
               Announceline3 = alternate2
           Case 3
               Announceline3 = alternate3
        End Select

        frmAnnounce.Show Modal
        
        'reset biofailure stuff to avoid repeat messages
        Planet(X).BioFailed = False
    End If
Next X

'Message if planet was sabotaged - successful or not
For X = 0 To 49
    If Planet(X).Owner = Current And Planet(X).Sabotaged Then
        'MessageType = "Sabotage"
        If Planet(X).SabotageReduction = 0 Then
            'mission was a FAILURE
            MessageType = "Sabotage Failed"
            
            '**this should be in frmAnnounce as well, not just a msgbox...
            
            Announceline1 = "Reports from " + Planet(X).Name + " confirm that enemy"
            Announceline2 = "forces tried to sabotage its resource production --"
            Announceline3 = "The cowards were destroyed!"
            
            frmAnnounce.Show Modal

            'Dim msg1 As String
            'Dim msg2 As String
            'Dim msg3 As String
            
            'msg1 = "Reports from " + Planet(X).Name + " confirm that enemy"
            'msg2 = "forces tried to sabotage its resource production --"
            'msg3 = "The cowards were destroyed!"
            
            'PlaySoundEffect "Warning"
            'MsgBox msg1 + Chr(13) + msg2 + Chr(13) + msg3, vbExclamation, "Sabotage Alert!"
            
            
            
            'don't repeat this message every turn!!!
            Planet(X).Sabotaged = False
            Planet(X).SabotageReduction = 0
            Planet(X).SabotagedFactory = False

        ElseIf Planet(X).SabotageReduction > 0 Or Planet(X).SabotageReduction = -1 Then
            'NOTE: variable can be -1 if the planet already had 0 resources when sabotaged...
            
            'mission was a SUCCESS
            MessageType = "Sabotage"
            
            Announceline1 = "Reports from " + Planet(X).Name + " confirm that saboteurs"
            If Planet(X).SabotagedFactory Then
                Announceline2 = "have destroyed their resource production!"
            Else
                Announceline2 = "have breached the security perimeter!"
            End If
            Announceline3 = "Resource production reduced to " + Str(Planet(X).Resources)
            
            
            frmAnnounce.Show Modal
        
            'reset sabotage attributes, to avoid the message coming up every turn
            'and to allow the planet to be sabotaged again...
            Planet(X).Sabotaged = False
            Planet(X).SabotageReduction = 0
            Planet(X).SabotagedFactory = False
            
        'ElseIf Planet(X).SabotageReduction = -1 Then
            'do nothing - sabotaged a planet already at 0 resources
            
            'reset sabotage attributes, to avoid the message coming up every turn
            'and to allow the planet to be sabotaged again...
        '    Planet(X).Sabotaged = False
        '    Planet(X).SabotageReduction = 0
        '    Planet(X).SabotagedFactory = False
        End If
        
    End If
Next X


End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen
txtStatus.Text = ""
End Sub











Private Sub Form_Resize()
RefreshWarpPath
End Sub

Private Sub fraEnemyWarpPath_DragDrop(Source As Control, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen

End Sub

Private Sub fraLanding_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""

End Sub


Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""
txtStatus.ForeColor = vbGreen

End Sub


Private Sub fraPlayerStats_DragDrop(Source As Control, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen
txtStatus.Text = "Player Stats"

End Sub

Private Sub fraTactical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""

End Sub


Private Sub fraUpgrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtStatus.ForeColor = vbGreen
  txtStatus.Text = ""
End Sub




Private Sub fraWarpPath_DragDrop(Source As Control, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen

End Sub

Private Sub fraWarpPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""

End Sub


Private Sub hsbQuantity_Change()
'update the quantity label and projected cost

lblQuantity.Caption = Str(hsbQuantity.Value)
Quantity = hsbQuantity.Value
PurchasePrice = Quantity * UnitCost
txtTotal.Text = Str(PurchasePrice)

End Sub

Private Sub hsbQuantity_Scroll()
'update the quantity label and projected cost

lblQuantity.Caption = Str(hsbQuantity.Value)
Quantity = hsbQuantity.Value
PurchasePrice = Quantity * UnitCost
txtTotal.Text = Str(PurchasePrice)
End Sub


Private Sub lblTitle_DblClick()
'runs code under lblTitle2

lblTitle2_DblClick

End Sub


Private Sub lblTitle2_DblClick()
'displays the About box
Dim Msg As String
Dim title, dialogtype

dialogtype = vbOKOnly + vbInformation
title = "About 4000 A.D."
Msg = "4000 A.D. - version 2.5 " + Chr(13)
Msg = Msg + "c. 1997-1999 Gordon Stewart - All Rights Reserved " + Chr(13) + Chr(13)
Msg = Msg + "Visit 4000 A.D. on the internet at:" + Chr(13)
Msg = Msg + "http://www.interlog.com/~gordons/4000ad.html" + Chr(13) + Chr(13)
Msg = Msg + "Free source code now available"

PlaySoundEffect "Quiet"

MsgBox Msg, dialogtype, title



End Sub


Private Sub picGalaxy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtStatus.Text = ""
  
End Sub




Private Sub picPlanet_Click(Index As Integer)

'match the right ship number - with activeship

Dim CurrentPlayer, Enemy
CurrentPlayer = Current
Enemy = Other

Dim RandomLosses    'for calculating losses in battles...

'placeholders for troop numbers
Dim a, b
a = Player(Current).Ship(activeship).Troops
b = Player(Current).Ship(activeship).AssaultTroops

'****biorocket
'this bit seems to fix a problem with clicking your other planets
If Planet(Index).LaunchSite = False And Planet(Index).InBioRange = False And BioRocketOn Then
    EraseLines
End If

If Planet(Index).InBioRange And BioRocketOn Then
    PlaySoundEffect "Button2"
    If MsgBox("Target this planet?", vbYesNo + vbQuestion, "BioHazard Launch") = vbYes Then
        If Planet(Index).BioRocketETA > 0 Then
            PlaySoundEffect "Quiet"
            MsgBox "This planet has already been targeted", , "BioHazard Launch Aborted"
            'erase lines and exit sub
            EraseLines
            'erase circles if necessary
            If ReadyToLand1 Or ReadyToLand2 Then
                EraseCircles
            End If
        Else
            'if planet inbiorange, and not targeted previously...
            TargetBioRocket (Index)
        End If
        
    Else    'erase lines and exit sub
        EraseLines
        'erase circles if necessary
        If ReadyToLand1 Or ReadyToLand2 Then
            EraseCircles
        End If
    End If
        
    Exit Sub
End If
'**end biorocket

If (ReadyToLand1 = True Or ReadyToLand2 = True) And Planet(Index).InRange And Player(Current).Ship(activeship).Launched Then
    PlaySoundEffect "Quiet"
    If MsgBox("Do you want to land here?", vbYesNo + vbQuestion, "Landing") = vbYes Then
        'get rid of the circles
        EraseCircles
        
        '*****SABOTAGE*****
        If Player(Current).Ship(activeship).Sabotage Then
            'goto sabotage routine
            Call SabotageLanding(Index, activeship)
            ReInitializeShip (activeship)
            Exit Sub
        End If
        '****END OF SABOTAGE *****
        
        
        'select type of landing - friendly, neutral or attack
        Select Case Planet(Index).Owner
        
        Case CurrentPlayer
            'landing on planet player owns
            PlaySoundEffect "Quiet"
            MsgBox "Landing successful", , " "
            'update planet, ship stats, player stats
            UpdateNumPlanets
            UpdatePlayerStats
            
            'disable the landing button and remove from warp path
            If activeship = 0 Then
                cmdLandShip1.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            End If
            
            If activeship = 1 Then
                cmdLandShip2.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            End If
            
            
            'set ship values to unlaunched and empty
            ReInitializeShip (activeship)
            
            RefreshWarpPath
          
            'add the players stuff to the planet
            Planet(Index).Troops = Planet(Index).Troops + a
            Planet(Index).AssaultTroops = Planet(Index).AssaultTroops + b
            SetCombatStrength (Index)
            
                
            'set ship values to unlaunched and empty
            ReInitializeShip (activeship)
            
            'clear the management frame
            ClearFrame
            
            cmdPreviewShip1.Enabled = True
            cmdPreviewShip2.Enabled = True
            cmdPreviewEnemy1.Enabled = True
            cmdPreviewEnemy2.Enabled = True
            
        Case Enemy
            'landing on other player's planet
            PlaySoundEffect "Attack"
            
            Randomize
            
            'set placeholder for current player's combat strength
            AttackStrength = Player(Current).Ship(activeship).CombatStrength
            DefenceStrength = Planet(Index).CombatStrength
            
            'set placeholders for troops on the planet
            Dim c, d
            c = Planet(Index).Troops
            d = Planet(Index).AssaultTroops
            
            'compare the strengths
            If AttackStrength > DefenceStrength Then
                'ATTACKER WINS
                NumPlanetsCaptured = NumPlanetsCaptured + 1
                Planet(Index).Captured = True
                                
                'don't allow player to launch from this planet this turn
                Planet(Index).JustLanded = True
                    
                'winner loses % of troops based on defensive CS
                'initial value
                TroopLosses = 0
                If a > 0 And DefenceStrength > 0 Then
                    
                    Select Case DefenceStrength
                    Case 0      'no losses
                        TroopLosses = 0
                    Case 1 To 2  '0-10% losses
                        TroopLosses = Int(Rnd * 10)
                    Case 3 To 4  '5-15% losses
                        TroopLosses = Int(Rnd * 10) + 5
                    Case 5 To 7  '10-20%
                        TroopLosses = Int(Rnd * 10) + 10
                    Case 8 To 10  '15-25%
                        TroopLosses = Int(Rnd * 10) + 15
                    Case 11 To 15 '15-30%
                        TroopLosses = Int(Rnd * 15) + 15
                    Case 16 To 200 '20-35%
                        TroopLosses = Int(Rnd * 15) + 20
                    Case Else    '20-45%
                        TroopLosses = Int(Rnd * 25) + 20
                    End Select
                    
                    a = a - (a * (TroopLosses / 100))
                    
                    Debug.Print "**Troop Losses = " & Str(TroopLosses) & "%"
                End If
                
                'winner loses % of assault troops
                'initial value
                
                AssaultLosses = 0
                If b > 0 And DefenceStrength > 0 Then
                    Select Case DefenceStrength
                      Case 0   '0% losses
                          AssaultLosses = 0
                      Case 1 To 3  '0-10%
                          AssaultLosses = Int(Rnd * 10)
                      Case 4 To 7  '0-15%
                          AssaultLosses = Int(Rnd * 15)
                      Case 8 To 12 '5-15%
                          AssaultLosses = Int(Rnd * 10) + 5
                      Case 13 To 17 '10-20%
                          AssaultLosses = Int(Rnd * 10) + 10
                      Case 18 To 250 '15-25%
                          AssaultLosses = Int(Rnd * 10) + 15
                      Case Else    '15-35%
                          AssaultLosses = Int(Rnd * 20) + 15
                    End Select
                    
                    b = b - (b * (AssaultLosses / 100))

                End If
                                                              
                'planet troops, combatstrength changes
                Planet(Index).Troops = a
                Planet(Index).AssaultTroops = b

                'update other player's stats with c & d (above)
                Player(Other).NumTroops = Player(Other).NumTroops - Planet(Index).Troops
                Player(Other).NumAssaultTroops = Player(Other).NumAssaultTroops - d
                Player(Other).NumPlanets = Player(Other).NumPlanets - 1
                                
               'set this variable for frmlandscape:
                ActivePlanet = Index
                
                'show the results of the battle
                'owner of planet not changed yet, to show results of battle - see right below
                frmLandscape.Show Modal
                
                'planet changes owners
                Planet(Index).Owner = Current
                                
                'set planet's combat strength
                SetCombatStrength (Index)

                'update player: resources, numplanets
                UpdateNumPlanets
                UpdatePlayerStats
            
                'clear the management frame
                ClearFrame
                
                cmdPreviewShip1.Enabled = True
                cmdPreviewShip2.Enabled = True
                cmdPreviewEnemy1.Enabled = True
                cmdPreviewEnemy2.Enabled = True
            '************************************
            ElseIf AttackStrength <= Planet(Index).CombatStrength Then
                'DEFENDER WINS, and tie goes to the defender
                
                'notify other player next turn
                Planet(Index).FailedInvasion = True
                
                'lose all your troops
                Player(Current).NumTroops = Player(Current).NumTroops - a
                Player(Current).NumAssaultTroops = Player(Current).NumAssaultTroops - b
                        
                If Planet(Index).Troops > 0 Then
                   'defender wins, but loses % of troops based on attacker's CS
                   'initial value
                   TroopLosses = 0
                    
                    Select Case AttackStrength
                    Case 0      'no losses
                        TroopLosses = 0
                    Case 1 To 2  '0-10% losses
                        TroopLosses = Int(Rnd * 10)
                    Case 3 To 4  '0-15% losses
                        TroopLosses = Int(Rnd * 15)
                    Case 5 To 8  '5-15%
                        TroopLosses = Int(Rnd * 10) + 5
                    Case 9 To 12  '10-20%
                        TroopLosses = Int(Rnd * 10) + 10
                    Case 13 To 15 '10-30%
                        TroopLosses = Int(Rnd * 20) + 10
                    Case 16 To 20 '15-35%
                        TroopLosses = Int(Rnd * 20) + 15
                    Case 21 To 25 '15-40%
                        TroopLosses = Int(Rnd * 25) + 15
                    Case 26 To 35 '20-45%
                        TroopLosses = Int(Rnd * 25) + 20
                    Case 36 To 45 '25-40%
                        TroopLosses = Int(Rnd * 15) + 25
                    Case Else    '25-50%
                        TroopLosses = Int(Rnd * 25) + 25
                    End Select
                    
                    'for next turn's notification of failed invasion
                    Planet(Index).FailedInvasionTroopLosses = Int(c * (TroopLosses / 100))
                                        
                    c = c - (c * (TroopLosses / 100))

                    Planet(Index).Troops = c
                End If
                
                If Planet(Index).AssaultTroops > 0 Then
                    'defender/winner loses % of assault troops
                    'initial value 0
                    AssaultLosses = 0
                    
                    Select Case AttackStrength
                    Case 0 To 5 '0% losses
                        AssaultLosses = 0
                    Case 6 To 10 '0-10%
                        AssaultLosses = Int(Rnd * 10)
                        Debug.Print "Arrgh! Attackstrength:" + Str(AttackStrength), "Assaultlosses:" + Str(AssaultLosses)
                    Case 11 To 15 '0-15%
                        AssaultLosses = Int(Rnd * 15)
                    Case 16 To 25 '5-20%
                        AssaultLosses = Int(Rnd * 15) + 5
                    Case 26 To 40 '10-20%
                        AssaultLosses = Int(Rnd * 10) + 10
                    Case 41 To 60 '15-30%
                        AssaultLosses = Int(Rnd * 15) + 15
                    Case 61 To 80 '20-35%
                        AssaultLosses = Int(Rnd * 15) + 20
                    Case Else    '20-40%
                        AssaultLosses = Int(Rnd * 20) + 20
                    End Select
                    
                    Planet(Index).FailedInvasionMechLosses = Int(d * (AssaultLosses / 100))
                    
                    d = d - (d * (AssaultLosses / 100))

                    Planet(Index).AssaultTroops = d
                End If
                                  
                'update planet's combat strength
                SetCombatStrength (Index)
                UpdateNumPlanets
                UpdatePlayerStats
                      
                'clear management frame
                ClearFrame
                
                cmdPreviewShip1.Enabled = True
                cmdPreviewShip2.Enabled = True
                cmdPreviewEnemy1.Enabled = True
                cmdPreviewEnemy2.Enabled = True
               
               'set this variable for frmlandscape:
                ActivePlanet = Index
                'show the results of the battle
                frmLandscape.Show Modal
                             
            End If
            
            'disable the landing button and remove from warp path
            If activeship = 0 Then
                cmdLandShip1.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()

            ElseIf activeship = 1 Then
                cmdLandShip2.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            
            End If
            
            'set ship values to unlaunched and empty
            ReInitializeShip (activeship)
        
            RefreshWarpPath
            
        Case 2
            'landing on UNOWNED planet
            PlaySoundEffect "Quiet"
            MsgBox "A new planet for you!", , " "
            Planet(Index).Owner = Current
            Planet(Index).JustLanded = True    'so you can't take off again this turn
            
            'add the players stuff to the planet
            Planet(Index).Troops = a
            Planet(Index).AssaultTroops = b
            SetCombatStrength (Index)
            
            'update player: resources, numplanets
            UpdateNumPlanets
            UpdatePlayerStats
            
            'clear management frame
            ClearFrame
            
            cmdPreviewShip1.Enabled = True
            cmdPreviewShip2.Enabled = True
            cmdPreviewEnemy1.Enabled = True
            cmdPreviewEnemy2.Enabled = True
            
            'deal with the landing button and warp path picture
            If activeship = 0 Then
                cmdLandShip1.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            ElseIf activeship = 1 Then
                cmdLandShip2.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            End If
            
            'set ship values to unlaunched and empty
            ReInitializeShip (activeship)
            
            RefreshWarpPath
            
            'show the landscape screen with troop stats
            ActivePlanet = Index
            frmLandscape.Show Modal

        Case 3
            AttackAliens (Index)
        End Select
        
    
    Else  'answered NO to msgbox - turn off circles, set rtl false
      EraseCircles
      ReadyToLand1 = False
      ReadyToLand2 = False
      
      cmdPreviewShip1.Enabled = True
      cmdPreviewShip2.Enabled = True
      cmdPreviewEnemy1.Enabled = True
      cmdPreviewEnemy2.Enabled = True
      
      
      If Player(Current).Ship(0).Launched Then
         cmdLandShip1.Enabled = True
      End If
      
      If Player(Current).Ship(1).Launched Then
         cmdLandShip2.Enabled = True
      End If
      
      Exit Sub

    End If
    'quit without showing "Hey! Not your planet box"
Exit Sub

End If

'******************************
'*****Planet Management********
'******************************
If Planet(Index).Owner = Current Then
    On Error Resume Next
    PlaySoundEffect "Ambient3"

    'PlayRandomSound
    On Error GoTo 0
    'clear frame first
    ClearFrame
        
   'enable the management box and the ship box
   fraUpgrade.Enabled = True
   cmdOK.Enabled = True
   hsbQuantity.Enabled = False
   txtTotal.Text = ""
   fraLanding.Enabled = True
   fraTactical.Enabled = True
   
   'show what planet is active, set it to active
   ActivePlanet = Index
   cmdPlanetName.Caption = Planet(Index).Name
   
   '********turn on buttons as available:
   'launch button enabled if at least one ship available
   If Player(Current).Ship(0).Launched And Player(Current).Ship(1).Launched Then
        cmdLaunch.Enabled = False
   Else
        cmdLaunch.Enabled = True
   End If
   
   'scanner button enabled if planet has a scanner
   If Planet(Index).HaveScanner Then
        cmdScan.Enabled = True
   ElseIf Planet(Index).HaveScanner = False Then
        cmdScan.Enabled = False
   End If
   
   'RepairIndustry button enabled if planet is damaged
   If Planet(Index).Damaged Then
        cmdRepairIndustry.Enabled = True
   End If
   
   'Biohazard button enabled if tech researched
   If Player(Current).BioRocketResearched Then
       cmdLaunchBioRocket.Enabled = True
   End If
   
   'regenerate button enabled if tech researched and if needed
   If Player(Current).RegenerationResearched And Planet(Index).NukedResources Then
       cmdRegenerate.Enabled = True
   End If
    
   'detoxify button enabled if tech researched and if needed
   If Player(Current).BioCleanupResearched And Planet(Index).Contaminated Then
        cmdCleanup.Enabled = True
   End If

Else   'planet not owned by current player
    PlaySoundEffect "Access"
    MsgBox "Hey! Not your planet", vbOKOnly + vbExclamation, "Access Denied"
        
   'disable mgmt frame, erase circles, set rtl to false
   ClearFrame
   
   'should only erase circles if part of a landing scenario...
   If ReadyToLand1 = True Or ReadyToLand2 = True Then
       EraseCircles
       ReadyToLand1 = False
       ReadyToLand2 = False
   End If
   
   'clear the board of any inrange scanner settings
   Dim z As Integer
   For z = 0 To 49
       Planet(z).InScannerRange = False
   Next z
   
   Dim Count
   For Count = 0 To 49
     If Planet(Count).LaunchSite = True Then
        EraseLines
     End If
   Next Count
    
   cmdPreviewShip1.Enabled = True
   cmdPreviewShip2.Enabled = True
   cmdPreviewEnemy1.Enabled = True
   cmdPreviewEnemy2.Enabled = True
   
   If Player(Current).Ship(0).Launched Then
      cmdLandShip1.Enabled = True
   End If
   
   If Player(Current).Ship(1).Launched Then
      cmdLandShip2.Enabled = True
   End If
   
   
End If

End Sub

Private Sub picPlanet_DblClick(Index As Integer)

'go to planet view screen
cmdPlanetName_Click

End Sub


Private Sub picPlanet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'public const neutral and alien in .bas module
'Neutral = 2
'Alien = 3
If ScannerOn Then
    'show full stats if in range
    If Planet(Index).Owner = Current Then
        txtStatus.ForeColor = vbYellow
        'show full stats
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources & "  Troops:" & Planet(Index).Troops _
        & "  Mechs: " & Planet(Index).AssaultTroops & _
        "  CS:" & Planet(Index).CombatStrength
    
    ElseIf Planet(Index).Owner = Neutral Then
        'show basic stats in green
        txtStatus.ForeColor = vbGreen
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources
        
    ElseIf Planet(Index).Owner = Alien And Planet(Index).InScannerRange Then
        'show full stats in blue
        txtStatus.ForeColor = vbBlue
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources & "  Troops:" & Planet(Index).Troops _
        & "  Mechs:" & Planet(Index).AssaultTroops & _
        "  CS:" & Planet(Index).CombatStrength
    
    ElseIf Planet(Index).Owner = Other And Planet(Index).InScannerRange Then
        'show full stats in red
        txtStatus.ForeColor = vbRed
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources & "  Troops:" & Planet(Index).Troops _
        & "  Mechs:" & Planet(Index).AssaultTroops & _
        "  CS:" & Planet(Index).CombatStrength
    ElseIf Planet(Index).Owner = Alien Then
        'show basic stats in blue
        txtStatus.ForeColor = vbBlue
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources

    ElseIf Planet(Index).Owner = Other Then
        'show basic stats in red
        txtStatus.ForeColor = vbRed
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources
        
    End If
    

ElseIf ScannerOn = False Then
    'full stats for owned planets, basic for everything else
    Select Case Planet(Index).Owner
    Case Current
        txtStatus.ForeColor = vbYellow
        'show full stats
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources & "  Troops:" & Planet(Index).Troops _
        & "  Mechs:" & Planet(Index).AssaultTroops & _
        "  CS:" & Planet(Index).CombatStrength
    Case Neutral
        'show basic stats in green
        txtStatus.ForeColor = vbGreen
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources
    Case Alien
        'show full stats in blue
        txtStatus.ForeColor = vbBlue
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources
    
    Case Other
        'show full stats in red
        txtStatus.ForeColor = vbRed
        txtStatus.Text = Planet(Index).Name & ":  Resources:" & Planet(Index).Resources

    End Select

End If

End Sub




Private Sub picUpgrade_DblClick(Index As Integer)
'turn off scanner if it's on
If ScannerOn Then
    cmdScan_Click
    ScannerOn = False
End If


'start buying procedure
'see what is being bought/researched
txtStatus.ForeColor = vbGreen

Select Case Index
Case 0
    'missile defences for planet
    'fixed price of 10/planet
    '***moving Unitcost into the 'if tech researched' part
    'UnitCost = 10
    
    'exit sub if planet already has missiles
     If Planet(ActivePlanet).HaveMissiles Then
        PlaySoundEffect "Quiet"
        MsgBox "This planet is already equipped with missiles", vbInformation, "Transaction Cancelled"
        Exit Sub
     End If
    
    'check if technology researched
    If Player(Current).Missile1Researched Then
        'check if enough money!
        UnitCost = 10
        If Player(Current).NumResources >= UnitCost Then
            'first, enable the scrollbar
            hsbQuantity.Enabled = True
            hsbQuantity.Value = 0
            lblItemName = "Missile Defences"
            lblQuantity = Str(hsbQuantity.Value)
            txtTotal.Text = Str(hsbQuantity.Value)
            'get max value for scroll bar
            hsbQuantity.Max = 1
        Else
            PlaySoundEffect "Quiet"
            MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If
    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If

Case 1
    'Planetary shield - fixed cost/planet
    '****testing moving this into 'if tech researched'
    'UnitCost = 15
    
    'exit sub if planet already has a shield
     If Planet(ActivePlanet).HaveShields Then
        PlaySoundEffect "Quiet"
        MsgBox "This planet is already protected by a planetary shield", vbInformation, "Transaction Cancelled"
        Exit Sub
     End If
     
    If Player(Current).ShieldResearched Then
        'check if enough money!
        UnitCost = 15
        If Player(Current).NumResources >= UnitCost Then
            'enable the scrollbar
            hsbQuantity.Enabled = True
            hsbQuantity.Value = 0
            lblItemName = "Planetary Shield"
            lblQuantity = Str(hsbQuantity.Value)
            txtTotal.Text = Str(hsbQuantity.Value)
            'get max value for scroll bar
            hsbQuantity.Max = 1
        Else
            PlaySoundEffect "Quiet"
            MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If

    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If

Case 2
    'improved resource production - fixed cost/planet
    'unitcost moved into 'if tech researched' part
    'UnitCost = 15
        
        'check for resources already high enough - ie 5
     If Planet(ActivePlanet).ImprovedResources Or Planet(ActivePlanet).Resources > 5 Then
        PlaySoundEffect "Quiet"
        MsgBox "This planet has already maximized its production capacity", vbInformation, "Transaction Cancelled"
        Exit Sub
     End If

    'check if player has researched this item
    If Player(Current).ResourcesResearched Then
        'check if enough money!
        UnitCost = 15
        If Player(Current).NumResources >= UnitCost Then
            'first, enable the scrollbar
            hsbQuantity.Enabled = True
            hsbQuantity.Value = 0
            lblItemName = "Improved Resource Production"
            lblQuantity = Str(hsbQuantity.Value)
            txtTotal.Text = Str(hsbQuantity.Value)
            'get max value for scroll bar
            hsbQuantity.Max = 1
        Else
            PlaySoundEffect "Quiet"
            MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If
    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If
  
Case 3
    'regular troops
    
    UnitCost = 1
    'check if enough money!
    If Player(Current).NumResources >= UnitCost Then
        'first, enable the scrollbar
        hsbQuantity.Enabled = True
        hsbQuantity.Value = 0
        lblItemName = "Troops"
        lblQuantity = Str(hsbQuantity.Value)
        txtTotal.Text = Str(hsbQuantity.Value)
        'get max value for scroll bar
        hsbQuantity.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
    End If
    
Case 4
    'buying assault troops
    'UnitCost = 4
    'check if player has researched this item
    If Player(Current).MechResearched Then
        UnitCost = 4
        'check if enough money
        If Player(Current).NumResources >= UnitCost Then
           'enable the scrollbar
           hsbQuantity.Enabled = True
           hsbQuantity.Value = 0
           lblItemName = "Assault Mechs"
           lblQuantity = Str(hsbQuantity.Value)
           txtTotal.Text = Str(hsbQuantity.Value)
           'get max value for scroll bar
           hsbQuantity.Max = Int(Player(Current).NumResources / UnitCost)
        Else
           PlaySoundEffect "Quiet"
           MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If
    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If
    
Case 5
    'here is where I load a tech research form
    'with buttons for assault troops, ship tech, resource tech, and planet defenses
    frmResearch.Show Modal

Case 6
    'scanners
    'UnitCost = 25
    
    'exit sub if planet already has a shield
     If Planet(ActivePlanet).HaveScanner Then
        PlaySoundEffect "Quiet"
        MsgBox "This planet already has a scanner", vbInformation, "Transaction Cancelled"
        Exit Sub
     End If
     
    If Player(Current).ScannerResearched Then
        UnitCost = 25
        'check if enough money!
        If Player(Current).NumResources >= UnitCost Then
            'enable the scrollbar
            hsbQuantity.Enabled = True
            hsbQuantity.Value = 0
            lblItemName = "Scanner"
            lblQuantity = Str(hsbQuantity.Value)
            txtTotal.Text = Str(hsbQuantity.Value)
            'get max value for scroll bar
            hsbQuantity.Max = 1
        Else
            PlaySoundEffect "Quiet"
            MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If

    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If

Case 7
    'jammers
    'UnitCost = 15
    
    'exit sub if planet already has a shield
     If Planet(ActivePlanet).HaveJammer Then
        PlaySoundEffect "Quiet"
        MsgBox "This planet already has a jamming device", vbInformation, "Transaction Cancelled"
        Exit Sub
     End If
     
    If Player(Current).JammerResearched Then
        'check if enough money!
        UnitCost = 15
        If Player(Current).NumResources >= UnitCost Then
            'enable the scrollbar
            hsbQuantity.Enabled = True
            hsbQuantity.Value = 0
            lblItemName = "Scanner Jamming Device"
            lblQuantity = Str(hsbQuantity.Value)
            txtTotal.Text = Str(hsbQuantity.Value)
            'get max value for scroll bar
            hsbQuantity.Max = 1
        Else
            PlaySoundEffect "Quiet"
            MsgBox "Insufficient resources", vbExclamation, "Transaction Denied"
        End If

    Else
        PlaySoundEffect "Quiet"
        MsgBox "You do not have this technology", vbExclamation, "Access Denied"
    End If

End Select

End Sub

Private Sub picUpgrade_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.ForeColor = vbGreen


'show text describing each of the upgrade icons
Select Case Index
Case 0
    txtStatus.Text = "Planetary Missile Defences"
Case 1
    txtStatus.Text = "Planetary Shield Defenses"
Case 2
    txtStatus.Text = "Improved Resource Production Facility"
Case 3
    txtStatus.Text = "Recruit Troops"
Case 4
    txtStatus.Text = "Recruit Assault Troops"
Case 5
    txtStatus.Text = "Research Advanced Technologies"
Case 6
    txtStatus.Text = "Space Scanner"
Case 7
    txtStatus.Text = "Anti-Scanning Jammer"
End Select

End Sub






Private Sub tmrRandomSounds_Timer()
'play random sound effect
PlayRandomSound

End Sub


Private Sub tmrUpdateMessageBox_Timer()
txtMessages.Text = "No New Messages..."
tmrUpdateMessageBox.Enabled = False

End Sub

Private Sub txtMessages_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub txtMessages_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtMessages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = "View or Send Messages"
End Sub


Private Sub txtNumAssaultTroops_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub txtNumAssaultTroops_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtNumPlanets_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0

End Sub

Private Sub txtNumPlanets_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtNumResources_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub txtNumResources_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtNumTroops_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub txtNumTroops_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtPlayerName_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub txtPlayerName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtPlayerName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""
End Sub

Private Sub txtProduction_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0

End Sub


Private Sub txtProduction_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub


Private Sub txtStatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub


Private Sub txtStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = ""
End Sub











Public Sub DrawGalaxy()
Randomize
'Called by the form_Activate procedure

'draw white stars on the playing field
Dim a, X, Y
For a = 1 To 700
    X = Int(Rnd * picGalaxy.ScaleWidth)
    Y = Int(Rnd * picGalaxy.ScaleHeight)
    picGalaxy.PSet (X, Y)
Next a

'draw darker stars for depth
    Dim grey
    grey = &H808080
    For a = 1 To 1000
        X = Int(Rnd * picGalaxy.ScaleWidth)
        Y = Int(Rnd * picGalaxy.ScaleHeight)
        picGalaxy.PSet (X, Y), grey
    Next a
       
'drawing grid lines moved to its own proc
End Sub

Private Sub txtTurnNumber_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub txtTurnNumber_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub








Public Sub UpdatePlayerStats()
'fill in values for player stats box

txtNumPlanets.Text = Str(Player(Current).NumPlanets)
txtNumTroops.Text = Str(Player(Current).NumTroops)
txtNumResources.Text = Str(Player(Current).NumResources)
txtNumAssaultTroops.Text = Str(Player(Current).NumAssaultTroops)

'redo resources/turn box
Dim h, i
i = 0
For h = 0 To 49
   If Planet(h).Owner = Current Then
        i = i + Planet(h).Resources
    End If
Next h

txtProduction.Text = Str(i)

End Sub

Public Sub Personalize()
'set caption and player name to current player
If Current = 0 Then
   frmGameScreen.Caption = " 4000 A.D.  (Player 1)   "
ElseIf Current = 1 Then
   frmGameScreen.Caption = " 4000 A.D.  (Player 2)   "
End If

If Player(1).Name = "" Then
    Player(1).Name = "?"
End If

frmGameScreen.Caption = frmGameScreen.Caption + "                            " + Player(0).Name + " vs. " + Player(1).Name

If Player(1).Name = "?" Then
    Player(1).Name = ""
End If


'set the current player's name in the name box
txtPlayerName.Text = Player(Current).Name

End Sub

Public Sub UpdateResources()
'at start of each turn, add up resources
'current plus resources of each owned planet

Dim i
For i = 0 To 49
  If Planet(i).Owner = Current Then
     Player(Current).NumResources = Player(Current).NumResources + Planet(i).Resources
  End If
Next i
  
End Sub







Public Sub UpdateNumPlanets()
'at start of each turn - and after getting new planets,
'and after battles, add up number of planets owned
'and total troops/assault troops

'first, set numplanet to zero to avoid recounting
Player(Current).NumPlanets = 0
Player(Current).NumTroops = 0
Player(Current).NumAssaultTroops = 0

Dim i

For i = 0 To 49
  If Planet(i).Owner = Current Then
     Player(Current).NumPlanets = Player(Current).NumPlanets + 1
     Player(Current).NumTroops = Player(Current).NumTroops + Planet(i).Troops
     Player(Current).NumAssaultTroops = Player(Current).NumAssaultTroops + Planet(i).AssaultTroops

  End If
Next i

End Sub

Public Sub ReInitializeShip(activeship As Integer)
'set all the values to zero, not launched, etc
'using ActiveShip as the indicator

With Player(Current).Ship(activeship)
    .Launched = False
    .Troops = 0
    .AssaultTroops = 0
    .CombatStrength = 0
    .HaveShields = False
    .HaveWeapons = False
    .HaveCloakingDevice = False
    .Sabotage = False
    .WarpPosition = 0
    .Coordinate = ""
    .CenterX = 0
    .CenterY = 0
End With
    
End Sub

Public Sub ReInitializePlanet(Index As Integer)
'add the players stuff to the planet

Planet(Index).Troops = Player(Current).Ship(activeship).Troops
Planet(Index).AssaultTroops = Player(Current).Ship(activeship).AssaultTroops

End Sub

Public Sub EraseCircles()
'redraw the circles to erase them
'***Only erases yellow circles!!!

Dim xpos, ypos, radius
Dim z
    For z = 0 To 49
    If Planet(z).InRange And picPlanet(z).Visible Then
        'find center of the picturebox and draw circle
        xpos = picPlanet(z).Left + (picPlanet(z).Width / 2)
        ypos = picPlanet(z).Top + (picPlanet(z).Height / 2)
        radius = (picPlanet(z).Width / 2) + 45
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 1
        picGalaxy.Circle (xpos, ypos), radius, vbYellow
    End If
    Next z

'reset the inrange value to false to prevent screwy drawing
For z = 0 To 49
    Planet(z).InRange = False
Next z

'set the readytoland value to false to prevent screwy drawing...
ReadyToLand1 = False
ReadyToLand2 = False

End Sub

Private Sub txtTurnNumber_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtStatus.Text = " "
End Sub



Public Sub ClearFrame()
'clear up the management frame - set everything to zero etc.
'reset the scroll bar and labels

hsbQuantity.Enabled = False
lblItemName.Caption = ""
lblQuantity.Caption = ""
txtTotal.Text = ""
'reset the purchase price to zero
'**this is used to initialize txtTotal.text
PurchasePrice = 0

'disable the management and ship frames again
cmdPlanetName.Caption = ""
fraUpgrade.Enabled = False
cmdOK.Enabled = False

'turn off scanner
If ScannerOn Then
    cmdScan_Click
    ScannerOn = False
    cmdScan.Enabled = False
End If

If RegenerateOn Then
    RegenerateOn = False
    cmdRegenerate.Enabled = False
End If

If BioRocketOn Then
    BioRocketOn = False
End If

'turn off all the buttons
cmdRepairIndustry.Enabled = False
cmdLaunchBioRocket.Enabled = False
cmdRegenerate.Enabled = False
cmdCleanup.Enabled = False

'deal with landing buttons - conflict with scanner button
If ReadyToLand1 Then
    cmdLandShip1_Click
End If

If ReadyToLand2 Then
    cmdLandShip2_Click
End If

End Sub



Public Sub CheckForAliens()
'see if weak planet is attacked by aliens

Dim ChanceOfInvasion
Dim Result As Integer
Dim X As Integer
Dim PlanetCS As Integer
Dim AlienCS As Integer

'*****************
'**Experimented with increasing minimums as game went on,
'**but decided it would affect play balance too much
'Dim MinTroops As Integer    'threshold # for invasion test
'Dim MinMechs As Integer

'calculate the strength of planet that is weak enough for invasion
'MinTroops = Int(Rnd * (TurnNumber / 5)) + 2
'MinMechs = Int(Rnd * (TurnNumber / 8)) + 1
'******************

Randomize

For X = 0 To 49
    If Planet(X).Owner = Current Then
        If Planet(X).Troops < 4 And Planet(X).AssaultTroops < 2 Then
            '5-15% chance
            ChanceOfInvasion = Int(Rnd * 10) + 5
            Result = Int(Rnd * 100) + 1
            
            If Result <= ChanceOfInvasion Then
                'planet attacked
                'compare combatstrengths
                '***Alien strength builds with turn number!
                AlienCS = Int(Rnd * 10) + Int(TurnNumber / 3)
                
                If AlienCS > Planet(X).CombatStrength Then
                    'Aliens Win
                    With Planet(X)
                        .Owner = Alien
                        .Troops = AlienCS - Planet(X).CombatStrength
                        .Resources = Int(Rnd * 6) + 1
                    End With
                    SetCombatStrength (X)
                    'MsgBox Planet(X).Name + " Overrun By The Melnikons!", vbOKOnly + vbExclamation, "Alien Invasion!"
                    '******
                    'Set various messages for overrun
                    Dim alternate1 As String
                    Dim alternate2 As String
                    Dim alternate3 As String
                    Dim alternate4 As String
                    Dim choice As Integer
                    Randomize
                    choice = Int(Rnd * 4) + 1
                    
                    alternate1 = "Defenses Breached...Casualties Mounting..."
                    alternate2 = "Security Perimeter Down...Shields Failed"
                    alternate3 = "Warning Systems Sabotaged...Need Help!..."
                    alternate4 = "They're Everywhere! If we can just--"

                    MessageType = "Overrun"
                    Announceline1 = "Melnikon Invasion Reported On " + Planet(X).Name + "!"
                    Select Case choice
                    Case 1
                        Announceline2 = alternate1
                    Case 2
                        Announceline2 = alternate2
                    Case 3
                        Announceline2 = alternate3
                    Case 4
                        Announceline2 = alternate4
                    End Select
                    Announceline3 = "<End Of Transmission>"
                    frmAnnounce.Show Modal
                    
                ElseIf AlienCS <= Planet(X).CombatStrength Then
                    'Player Wins
                    Dim Msg As String
                    Msg = "Melnikon Invasion Force Destroyed!"
                                     
                    AlienCS = Int(AlienCS / 2)
                    'lose troops if any
                    If Planet(X).Troops > 0 Then
                        Planet(X).Troops = Planet(X).Troops - AlienCS
                        If Planet(X).Troops < 0 Then
                            Planet(X).Troops = 0
                        End If
                    End If
                    
                    'lose some mechs if any
                    If Planet(X).AssaultTroops > 0 Then
                        Planet(X).AssaultTroops = Planet(X).AssaultTroops - Int(AlienCS / 2)
                        If Planet(X).AssaultTroops < 0 Then
                            Planet(X).AssaultTroops = 0
                        End If
                    End If
                    
                    SetCombatStrength (X)
                    
                    MessageType = "Victorious"
                    Announceline1 = "Melnikon Invasion of " + Planet(X).Name + " Defeated!"
                    If Planet(X).Owner = Current Then
                        Announceline2 = "Troop Losses:" + Str(AlienCS)
                        Announceline3 = "Mech Losses:" + Str(Int(AlienCS / 2))
                    Else
                        Announceline2 = ""
                        Announceline3 = ""
                    End If
                    frmAnnounce.Show Modal

                    'MsgBox Msg, vbOKOnly + vbExclamation, "Report From " + Planet(X).Name
                    'Debug.Print "AlienCS/2:" + Str(AlienCS)
                End If
            End If
        End If
    End If
Next X


End Sub

Public Sub AttackAliens(Index As Integer)
'attack procedure when landing on alien planets
Dim RandomLosses
Dim a, b

'set random number generator
Randomize
            
'set placeholder for current player's combat strength
AttackStrength = Player(Current).Ship(activeship).CombatStrength
DefenceStrength = Planet(Index).CombatStrength
            
a = Player(Current).Ship(activeship).Troops
b = Player(Current).Ship(activeship).AssaultTroops

'set placeholders for troops on the planet
Dim c, d
c = Planet(Index).Troops
d = Planet(Index).AssaultTroops

PlaySoundEffect "Attack"
'MsgBox "Attack!", , " "
                       
'compare the strengths
If AttackStrength > DefenceStrength Then
    'ATTACKER WINS
    'don't allow player to launch from this planet this turn
    Planet(Index).JustLanded = True
                
    If a > 0 And DefenceStrength > 0 Then
        'winner loses 10-50% of troops
        RandomLosses = Int(Rnd * 4) + 1
        RandomLosses = RandomLosses / 10
        a = a - (a * RandomLosses)
        TroopLosses = RandomLosses * 100
    End If
                                               
    If b > 0 And DefenceStrength > 0 Then
        'winner loses 10-30% of assault troops
        RandomLosses = Int(Rnd * 3) + 1
        RandomLosses = RandomLosses / 10
        b = b - (b * RandomLosses)
        AssaultLosses = RandomLosses * 100
    End If
                                                              
    'planet troops, combatstrength changes
    Planet(Index).Troops = a
    Planet(Index).AssaultTroops = b

    'set this variable for frmlandscape:
    ActivePlanet = Index
    
    'show the results of the battle
    frmLandscape.Show Modal
                
    'planet changes owners
    Planet(Index).Owner = Current
                                
    'set planet's combat strength
    SetCombatStrength (Index)
                
    'update player: resources, numplanets
    UpdateNumPlanets
    UpdatePlayerStats
            
    'clear the management frame
    ClearFrame
    
    cmdPreviewShip1.Enabled = True
    cmdPreviewShip2.Enabled = True
    cmdPreviewEnemy1.Enabled = True
    cmdPreviewEnemy2.Enabled = True
    
    '************************************
ElseIf AttackStrength <= Planet(Index).CombatStrength Then
    'defender wins, and tie goes to the defender
               
    'current player loses all troops
    Player(Current).NumTroops = Player(Current).NumTroops - a
    Player(Current).NumAssaultTroops = Player(Current).NumAssaultTroops - b
                        
    If Planet(Index).Troops > 0 Then
        'winner loses 10-50% of troops
        RandomLosses = Int(Rnd * 4) + 1
        RandomLosses = RandomLosses / 10
        c = c - (c * RandomLosses)
        TroopLosses = RandomLosses * 100
                    
        Planet(Index).Troops = c
    End If
                
    If Planet(Index).AssaultTroops > 0 Then
        'winner loses 10-40% of assault troops
        RandomLosses = Int(Rnd * 4) + 1
        RandomLosses = RandomLosses / 10
        d = d - (d * RandomLosses)
        AssaultLosses = RandomLosses * 100
        Planet(Index).AssaultTroops = d
    End If
                                  
    'update planet's combat strength
    SetCombatStrength (Index)
    UpdateNumPlanets
    UpdatePlayerStats
                      
    'clear management frame
    ClearFrame
    
    cmdPreviewShip1.Enabled = True
    cmdPreviewShip2.Enabled = True
    cmdPreviewEnemy1.Enabled = True
    cmdPreviewEnemy2.Enabled = True
    
    'set this variable for frmlandscape:
    ActivePlanet = Index
                
    'show the results of the battle
    frmLandscape.Show Modal
                             
End If
            
'disable the landing button and remove from warp path
If activeship = 0 Then
    cmdLandShip1.Enabled = False
    'get rid of the ship picture
    picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
ElseIf activeship = 1 Then
    cmdLandShip2.Enabled = False
    'get rid of the ship picture
    picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
End If
            
'set ship values to unlaunched and empty
ReInitializeShip (activeship)

RefreshWarpPath

End Sub



Public Sub SabotageLanding(Index As Integer, activeship As Integer)
'determine results of sabotage mission

Dim CurrentPlayer, Enemy
CurrentPlayer = Current
Enemy = Other

Select Case Planet(Index).Owner

Case CurrentPlayer
    'do nothing, mission wasted by landing on player's own planet
    PlaySoundEffect "Quiet"
    MsgBox "Mission aborted", vbOKOnly, " "
    
Case Enemy
    Dim Success As Integer
    Dim Result As Integer
    Success = 95
    
    'success varies with number of troops, mechs and technology on planet
    Select Case Planet(Index).Troops
    Case 0 To 5
        Success = Success - 1
    Case 6 To 10
        Success = Success - 2
    Case 11 To 15
        Success = Success - 4
    Case Else
        Success = Success - 6
    End Select
    
    'mechs
    Select Case Planet(Index).AssaultTroops
    Case 0 To 1
        Success = Success - 1
    Case 2 To 5
        Success = Success - 3
    Case Else
        Success = Success - 5
    End Select
    
    'missiles
    If Planet(Index).HaveMissiles Then
        'check what level of missiles other player has researched
        If Player(Other).Missile2Researched Then
            Success = Success - 4
        Else
            Success = Success - 2
        End If
    End If
    
    If Planet(Index).HaveShields Then
        Success = Success - 3
    End If
    
    If Planet(Index).HaveScanner Then
        Success = Success - 5
    End If
    
    Result = Int(Rnd * 100)
    
    Debug.Print "Result: ", Result, " Success: ", Success
    
    If Result <= Success Then
        'mission successful
        
        Dim Damage As Integer
        Dim Reduction
        Dim FactoryFlag As Boolean
        Dim msg1 As String
        Dim msg2 As String
        
        msg1 = "Advanced production facilities destroyed!"
        msg2 = "Planet's resource production eliminated!"
        Damage = Int(Rnd * 4) + 2  '2-6 damage
        
        If Planet(Index).ImprovedResources Then
            'factory destroyed
            Planet(Index).ImprovedResources = False
            FactoryFlag = True
            'tell other player his factory destroyed
            Planet(Index).SabotagedFactory = True
        End If
        
        'Fixes divide by zero error!!
        If Planet(Index).Resources < 1 Then
            Reduction = -1          'NOTE: if the same planet is sabotaged twice in 1 turn
                                    'there may not be a message to the other player if the
                                    'first sabotage reduced resources to zero
        Else
            Reduction = Damage / Planet(Index).Resources
            Reduction = Int(Reduction * 100)
        End If
        
        Planet(Index).Resources = Planet(Index).Resources - Damage
        
        If Planet(Index).Resources < 0 Then
            Planet(Index).Resources = 0
        End If
        
        If Planet(Index).Resources > 0 And FactoryFlag Then
            PlaySoundEffect "Quiet"
            MsgBox "Mission accomplished!" + Chr(13) + msg1 + Chr(13) + "Resource production reduced by " + Str(Reduction) + "%", vbOKOnly, "Sabotage Mission Results"
        ElseIf Planet(Index).Resources > 0 And Not FactoryFlag Then
            PlaySoundEffect "Quiet"
            MsgBox "Resource production on " + Planet(Index).Name + " crippled!" + Chr(13) + "Resource production reduced by" + Str(Reduction) + "%", vbOKOnly, "Sabotage Mission Results"
        ElseIf Planet(Index).Resources <= 0 Then
            PlaySoundEffect "Quiet"
            MsgBox msg2, , "Sabotage Mission Results"
        End If
        
        '***should set flag to give other player message at startup
        Planet(Index).Sabotaged = True
              
        Planet(Index).SabotageReduction = Reduction
        
        'set flag to enable cmdRepairIndustry
        Planet(Index).Damaged = True
        
    ElseIf Result > Success Then
        'MISSION FAILED
        
        PlaySoundEffect "Quiet"
        MsgBox "Mission failed - ship destroyed in orbit around " + Planet(Index).Name, vbOKOnly, "Sabotage Mission Results"
        'set flag to tell other player next turn
        Planet(Index).Sabotaged = True
        Planet(Index).SabotageReduction = 0 'this will show that mission failed, show different message
        Planet(Index).SabotagedFactory = False 'factory not destroyed
    End If
    
    
Case Neutral
    'do nothing
    PlaySoundEffect "Quiet"
    MsgBox "Mission failed - neutral planet", vbOKOnly, "Sabotage Mission Results"

Case Alien
    'reduce resources to 0
    Planet(Index).Resources = 0
    PlaySoundEffect "Quiet"
    MsgBox "Alien resource production facilities destroyed", vbOKOnly + vbExclamation, "Sabotage Mission Results"

End Select

EraseShip:
            'disable the landing button and remove from warp path
            If activeship = 0 Then
                cmdLandShip1.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            End If
            
            If activeship = 1 Then
                cmdLandShip2.Enabled = False
                'get rid of the ship picture
                picPlayerPath(Player(Current).Ship(activeship).WarpPosition - 1).Picture = LoadPicture()
            End If
            
            'set ship values to unlaunched and empty
            ReInitializeShip (activeship)
            
            RefreshWarpPath
            

End Sub

Public Sub BioDamage(X As Integer)
'this happens every turn on contaminated planets
Randomize
If Planet(X).Troops = 0 Or Planet(X).BioRocketETA = TurnNumber Then
    Exit Sub
End If

If Planet(X).Troops <= 2 Then
    Planet(X).Troops = Planet(X).Troops - 1
    If Planet(X).Troops <= 0 Then
       PlaySoundEffect "Quiet"
       MsgBox "Contamination warning: all troops dead.", vbOKOnly + vbExclamation, "Report From: " + Planet(X).Name
       Planet(X).Troops = 0
       Exit Sub
    End If
    'don't show msgbox if it's the same turn the rocket hits
    If Planet(X).BioRocketETA = TurnNumber Then
        'do nothing
    Else
        PlaySoundEffect "Warning"
        MsgBox "Contamination warning: 1 troop dead.", vbOKOnly + vbExclamation, "Report From: " + Planet(X).Name
    End If
    
ElseIf Planet(X).Troops > 2 Then
    Dim Damage As Integer
    Dim Dead As Integer
    
    'kill of 25-50% of troops
    Damage = Int(Rnd * 25) + 25
    
    'Damage = Damage / 100
    Dead = Int(Planet(X).Troops * Damage)
    Dead = Int(Dead / 100)
    
    'should be at least 1 dead per turn
    If Dead < 1 Then
        Dead = 1
    End If
    Planet(X).Troops = Planet(X).Troops - Dead
    
    Dim msg1 As String
    If Dead = 1 Then
        msg1 = " troop dead."
    Else
        msg1 = " troops dead."
    End If
    
    If Planet(X).BioRocketETA = TurnNumber Then
        'do nothing
    Else
        PlaySoundEffect "Warning"
        MsgBox "Contamination warning: " + Str(Dead) + msg1, vbOKOnly + vbExclamation, "Report From: " + Planet(X).Name
    End If
    
End If



End Sub





Public Sub EraseLines()
'redraw the lines to erase them

Dim x1, y1
Dim z As Integer

For z = 0 To 49
    If Planet(z).LaunchSite Then
        x1 = picPlanet(z).Left + (picPlanet(z).Width / 2)
        y1 = picPlanet(z).Top + (picPlanet(z).Height / 2)
        Exit For
    End If
Next z

Dim X2, Y2
For z = 0 To 49
    If Planet(z).InBioRange Then
        'find center of the picturebox and draw line
        X2 = picPlanet(z).Left + (picPlanet(z).Width / 2)
        Y2 = picPlanet(z).Top + (picPlanet(z).Height / 2)
        'set drawmode
        picGalaxy.DrawMode = 7
        picGalaxy.DrawWidth = 2
        
        picGalaxy.Line (x1, y1)-(X2, Y2), vbMagenta
    End If
Next z

'reset the inbiorange value to false to prevent screwy drawing
For z = 0 To 49
    Planet(z).InBioRange = False
Next z

'reset the launchsite to false
For z = 0 To 49
    If Planet(z).LaunchSite Then
        Planet(z).LaunchSite = False
        Exit For
    End If
Next z

'reset biorocketon to false
BioRocketOn = False

End Sub



Public Sub TargetBioRocket(Index As Integer)
'let player know how long it will take for biorocket to reach target

Dim ETA As Integer
Dim RocketCost As Integer

RocketCost = 30

'final cost check
If BioRocketOn Then
    'continue with sub
    ETA = Int(Planet(Index).BioDistance / 300)
    Planet(Index).BioRocketETA = TurnNumber + ETA

    If SoundOn Then
        'select which turn# sound file to play
        Select Case ETA
        Case 1
           Sound = App.Path + "\1turn.wav"
           sndPlaySound Sound, 3
        Case 2
           Sound = App.Path + "\2turns.wav"
           sndPlaySound Sound, 3
        Case 3
            Sound = App.Path + "\3turns.wav"
            sndPlaySound Sound, 3
        Case 4
           Sound = App.Path + "\4turns.wav"
           sndPlaySound Sound, 3
        Case 5
           Sound = App.Path + "\5turns.wav"
           sndPlaySound Sound, 3
        Case 6
           Sound = App.Path + "\6turns.wav"
           sndPlaySound Sound, 3
        Case 7
           Sound = App.Path + "\7turns.wav"
           sndPlaySound Sound, 3
        End Select
        
    End If
    
    PlaySoundEffect "Button3"
    MsgBox "Planet " + Planet(Index).Name + " targeted." + Chr(13) + Chr(13) + "BioHazard Rocket ETA:" + Str(ETA) + " Turns.", , " "

    'deduct funds
    Player(Current).NumResources = Player(Current).NumResources - RocketCost
    UpdatePlayerStats

    'erase lines and reset launchsite to false
    EraseLines
    
    'clear frame
    ClearFrame
Else
    Exit Sub
End If


End Sub

Public Sub WriteBigFile()
'save all the general info - message to other player...

Dim Filename As String
Dim ShortName As String


ShortName = "\gameinfo.txt"
Filename = App.Path & ShortName

'get a free file number
gFileNum = FreeFile

'create the file
Open Filename For Output As gFileNum

'write galaxysize
Write #gFileNum, GalaxySize

'write the planet data
Dim i
For i = 0 To 49
  Write #gFileNum, Planet(i).Name, Planet(i).Owner, Planet(i).Troops, _
  Planet(i).AssaultTroops, Planet(i).CombatStrength, Planet(i).Coordinate, _
  Planet(i).Resources, Planet(i).HaveMissiles, Planet(i).HaveShields, _
  Planet(i).ImprovedResources, Planet(i).HaveScanner, Planet(i).BackGround, _
  Planet(i).HaveJammer, Planet(i).BioRocketETA, Planet(i).Contaminated, Planet(i).NukedResources, _
  Planet(i).Sabotaged, Planet(i).SabotageReduction, Planet(i).SabotagedFactory, Planet(i).Damaged, _
  Planet(i).BioFailed
Next

'write the player data
For i = 0 To 1
  Write #gFileNum, Current, TurnNumber, Player(i).Name, Player(i).NumTroops, Player(i).NumAssaultTroops, Player(i).NumPlanets, Player(i).NumResources, _
  Player(i).HomePlanet, Player(i).Message1Given, Player(i).Message2Given, Player(i).WasBig, Player(i).Missile1ResearchDone, Player(i).Missile1Researched, _
  Player(i).Missile2ResearchDone, Player(i).Missile2Researched, Player(i).ShieldResearchDone, Player(i).ShieldResearched, _
  Player(i).LaserResearchDone, Player(i).LaserResearched, Player(i).PlasmaResearchDone, Player(i).PlasmaResearched, Player(i).MechResearchDone, Player(i).MechResearched, _
  Player(i).BioRocketResearchDone, Player(i).BioRocketResearched, Player(i).LongBioResearchDone, Player(i).LongBioResearched, Player(i).ShipShield1ResearchDone, Player(i).ShipShield1Researched, _
  Player(i).ShipShield2ResearchDone, Player(i).ShipShield2Researched, Player(i).BigShipResearchDone, Player(i).BigShipResearched, Player(i).UltraWarpResearchDone, Player(i).UltraWarpResearched, _
  Player(i).CloakingResearchDone, Player(i).CloakingResearched, Player(i).ResourceResearchDone, Player(i).ResourcesResearched, Player(i).BioCleanupResearchDone, Player(i).BioCleanupResearched, _
  Player(i).RegenerationResearchDone, Player(i).RegenerationResearched, Player(i).ScannerResearchDone, Player(i).ScannerResearched, Player(i).DeepScannerResearchDone, Player(i).DeepScannerResearched, _
  Player(i).JammerResearchDone, Player(i).JammerResearched, Player(i).WarpScannerResearchDone, Player(i).WarpScannerResearched
Next

'write the ship data
For i = 0 To 1
  Write #gFileNum, Player(0).Ship(i).Launched, Player(0).Ship(i).HaveCloakingDevice, _
  Player(0).Ship(i).Troops, Player(0).Ship(i).AssaultTroops, _
  Player(0).Ship(i).CombatStrength, Player(0).Ship(i).WarpPosition, _
  Player(0).Ship(i).Coordinate, Player(0).Ship(i).CenterX, _
  Player(0).Ship(i).CenterY, Player(0).Ship(i).ShipNumber, Player(0).Ship(i).Sabotage, _
  Player(1).Ship(i).Launched, Player(1).Ship(i).HaveCloakingDevice, _
  Player(1).Ship(i).Troops, Player(1).Ship(i).AssaultTroops, _
  Player(1).Ship(i).CombatStrength, Player(1).Ship(i).WarpPosition, _
  Player(1).Ship(i).Coordinate, Player(1).Ship(i).CenterX, _
  Player(1).Ship(i).CenterY, Player(1).Ship(i).ShipNumber, Player(1).Ship(i).Sabotage
Next

'write the general data
Write #gFileNum, OutgoingMessage, Player(Current).WasBig

'captured planet data
Write #gFileNum, NumPlanetsCaptured
For i = 0 To 49
    Write #gFileNum, Planet(i).Captured
Next i

'failed invasion data
Write #gFileNum, NumFailedInvasions
For i = 0 To 49
    Write #gFileNum, Planet(i).FailedInvasion, Planet(i).FailedInvasionTroopLosses, Planet(i).FailedInvasionMechLosses
Next i

'close the file
Close gFileNum


End Sub

Public Sub Detonation(X As Integer)
Randomize

'see if rocket hits, or is shot down by planet defenses
Dim Success As Integer
Dim Result As Integer

'set base chance of success
Success = 100
    
'success varies with technology on planet
If Planet(X).HaveMissiles Then
    If Player(Current).Missile1Researched Then
        Success = Success - 2
    End If
    
    If Player(Current).Missile2Researched Then
        Success = Success - 3
    End If
End If

If Planet(X).HaveShields Then
    Success = Success - 7
End If

If Planet(X).HaveScanner Then
   Success = Success - 2
End If
    
Result = Int(Rnd * 100)

If Result <= Success Then
    'rocket hits
    'explosion - kills troops and drastically reduces resource production
    If Planet(X).Owner = Current Then   'duh!
        'tally the damage to troops and resources
        If Planet(X).Troops > 0 Then
            Randomize
            Dim Damage As Integer
            Dim Dead As Integer
     
            'kill of 25-50% of troops
            Damage = Int(Rnd * 25) + 25
            Dead = Int(Planet(X).Troops * Damage)
            Dead = Int(Dead / 100)
 
            'should be at least 1 killed/turn
            If Dead < 1 Then
                Dead = 1
            End If
            Planet(X).Troops = Planet(X).Troops - Dead
        End If

        'resource production hurt bad
        Dim ProductionDamage        '3-8 damage to resources
    
        ProductionDamage = Int(Rnd * 5) + 3
        Planet(X).Resources = Planet(X).Resources - ProductionDamage
        
        If Planet(X).Resources < 0 Then
            Planet(X).Resources = 0
        End If
    
        'reset improved resource production to false, let them build again
        Planet(X).ImprovedResources = False

    End If  'end of if planet.owner=current loop
  
    'Announcement screen
    MessageType = "Explosion"
    If Planet(X).Owner = Current Then
        Announceline1 = "A Massive Explosion On " + Planet(X).Name + "!"
        Announceline2 = "Troop Losses: " + Str(Dead)
        Announceline3 = "Resource Production Reduced To:" + Str(Planet(X).Resources)
    Else
        Announceline1 = "Successful BioRocket Detonation On " + Planet(X).Name + "!"
        Announceline2 = ""
        Announceline3 = ""
    End If
    
    frmAnnounce.Show Modal

    Planet(X).Contaminated = True
    Planet(X).NukedResources = True

Else
    'result > success, therefore the missile didn't detonate!
    PlaySoundEffect "Quiet"
    MsgBox "BioRocket destroyed by defensive systems on " + Planet(X).Name + "!"
    'Debug.Print "Result:"; Result, "Success:"; Success
    
    'set warning to show other player on their next turn
    Planet(X).BioFailed = True
    
End If

End Sub



Public Sub RefreshWarpPath()
'put ships on the warp path with coordinates printed below
Dim z     'counter
Dim j, k  'hold warp positions - easier to type & read


j = Player(Current).Ship(0).WarpPosition
k = Player(Current).Ship(1).WarpPosition

If Player(Current).Ship(0).Launched And Player(Current).Ship(1).Launched And j = k Then
    'both ships on same warp box
    'check if in 2nd to last box
    If j = 7 And Warp7WarningGiven = False Then
        Dim Msg As String
        Msg = "Your ships must land next turn. Be advised" + Chr(13)
        Msg = Msg + "of the risk that no suitable planet will be" + Chr(13)
        Msg = Msg + "in range before the warp path disintegrates."
        'play warning and show message
        PlaySoundEffect "Disintegrate"
        MsgBox Msg, vbOKOnly + vbExclamation, "Warp Path Warning"
        Warp7WarningGiven = True
    End If
    'check if in last warp box
    If j = 8 And Warp8WarningGiven = False Then
        PlaySoundEffect "Disintegrate"
        LostInSpace
        If NumPlanets1 = 0 Then
            PlaySoundEffect "Warning"
            MsgBox "Ship 1 Destroyed", vbCritical, "Warp Path Disintegration"
            ReInitializeShip (0)
        End If
        
        If NumPlanets2 = 0 Then
            PlaySoundEffect "Warning"
            MsgBox "Ship 2 Destroyed", vbCritical, "Warp Path Disintegration"
            ReInitializeShip (1)
        End If
        
        If NumPlanets1 > 0 And NumPlanets2 > 0 Then
            PlaySoundEffect "Warning"
            MsgBox "Your ships must land this turn", vbOKOnly + vbCritical, "Warp Path Disintegrating!"
        End If
        Warp8WarningGiven = True
        'this is a flag to only show this warning once at the start
        'and once at the end of the turn
    End If
   
    picPlayerPath(j - 1).Picture = picTiny.Picture
    'set ship 1 coordinate on top left
    picPlayerPath(j - 1).CurrentX = 0
    picPlayerPath(j - 1).CurrentY = 0
    picPlayerPath(j - 1).Print Player(Current).Ship(0).Coordinate
    
    '***Put an S instead of CS if a sabotage mission
    If Player(Current).Ship(0).Sabotage Then
        picPlayerPath(j - 1).CurrentX = 450
        picPlayerPath(j - 1).CurrentY = 0
        picPlayerPath(j - 1).FontBold = True
        picPlayerPath(j - 1).Print "S"
        picPlayerPath(j - 1).FontBold = False
    Else
        'ship 1 CS on top right
        '***change currentx depending on value of CS
        If Player(Current).Ship(0).CombatStrength > 99 Then
            picPlayerPath(j - 1).CurrentX = 300
        ElseIf Player(Current).Ship(0).CombatStrength > 9 Then
            picPlayerPath(j - 1).CurrentX = 360
        Else
            picPlayerPath(j - 1).CurrentX = 440
        End If
        '***
        picPlayerPath(j - 1).CurrentY = 0
        picPlayerPath(j - 1).Print Player(Current).Ship(0).CombatStrength
    End If
    
    'set ship 2 coordinate at bottom left corner
    picPlayerPath(j - 1).CurrentX = 0
    picPlayerPath(j - 1).CurrentY = 465
    picPlayerPath(j - 1).Print Player(Current).Ship(1).Coordinate
        
    '***print S if a sabotage mission
    If Player(Current).Ship(1).Sabotage Then
        picPlayerPath(j - 1).CurrentX = 450
        picPlayerPath(j - 1).CurrentY = 465
        picPlayerPath(j - 1).FontBold = True
        picPlayerPath(j - 1).Print "S"
        picPlayerPath(j - 1).FontBold = False

    Else
        'set ship 2 CS at bottom right corner
        '***change currentx depending on value of CS
        If Player(Current).Ship(1).CombatStrength > 99 Then
            picPlayerPath(j - 1).CurrentX = 300
        ElseIf Player(Current).Ship(1).CombatStrength > 9 Then
            picPlayerPath(j - 1).CurrentX = 360
        Else
            picPlayerPath(j - 1).CurrentX = 440
        End If
        '***
        picPlayerPath(j - 1).CurrentY = 465
        picPlayerPath(j - 1).Print Player(Current).Ship(1).CombatStrength
    End If
    
Else
'ships not in same box
For z = 0 To 1
    If Player(Current).Ship(z).Launched Then
        j = Player(Current).Ship(z).WarpPosition
        'check if in 2nd to last box
        If j = 7 And Warp7WarningGiven = False Then
            Dim message As String
            message = "Your ship must land next turn. Be advised" + Chr(13)
            message = message + "of the risk that no suitable planet will be" + Chr(13)
            message = message + "in range before the warp path disintegrates."
            PlaySoundEffect "Warning"
            MsgBox message, vbOKOnly + vbExclamation, "Warp Path Warning"
            Warp7WarningGiven = True
        End If

        'check if in last warp box
        If j = 8 And Warp8WarningGiven = False Then
            PlaySoundEffect "Disintegrate"
            LostInSpace
            If NumPlanets1 = 0 Then
                PlaySoundEffect "Warning"
                MsgBox "Ship 1 Destroyed", vbCritical, "Warp Path Disintegration"
                ReInitializeShip (0)
            End If
        
            If NumPlanets2 = 0 Then
                PlaySoundEffect "Warning"
                MsgBox "Ship 2 Destroyed", vbCritical, "Warp Path Disintegration"
                ReInitializeShip (1)
            End If
                
            If NumPlanets1 > 0 And NumPlanets2 > 0 Then
                PlaySoundEffect "Warning"
                MsgBox "Your ship must land this turn", vbOKOnly + vbCritical, "Warp Path Disintegrating!"
            End If
            
            Warp8WarningGiven = True
        End If
        
    
            picPlayerPath(j - 1).Picture = picTemp.Picture
            'set shipnumber at top right
            picPlayerPath(j - 1).CurrentX = 435
            picPlayerPath(j - 1).CurrentY = 0
            picPlayerPath(j - 1).Print Player(Current).Ship(z).ShipNumber + 1
            'set coordinate at bottom left corner
            picPlayerPath(j - 1).CurrentX = 50
            picPlayerPath(j - 1).CurrentY = 465
            picPlayerPath(j - 1).Print Player(Current).Ship(z).Coordinate

        
        'set cursor at bottom right corner
        '**print S if sabotage mission
        If Player(Current).Ship(z).Sabotage Then
            picPlayerPath(j - 1).CurrentX = 450
            picPlayerPath(j - 1).CurrentY = 465
            picPlayerPath(j - 1).FontBold = True
            picPlayerPath(j - 1).Print "S"
            picPlayerPath(j - 1).FontBold = False

        Else
            '***change currentx depending on value of CS
            If Player(Current).Ship(z).CombatStrength > 99 Then
                picPlayerPath(j - 1).CurrentX = 300
            ElseIf Player(Current).Ship(z).CombatStrength > 9 Then
                picPlayerPath(j - 1).CurrentX = 360
            Else
                picPlayerPath(j - 1).CurrentX = 440
            End If
            '***

            picPlayerPath(j - 1).CurrentY = 465
            picPlayerPath(j - 1).Print Player(Current).Ship(z).CombatStrength
        End If
    End If
Next z

End If

If Player(Current).Ship(0).Launched Then
    cmdLandShip1.Enabled = True
ElseIf Player(Current).Ship(1).Launched Then
    cmdLandShip2.Enabled = True
End If


'put player number under owned planets
Dim Count As Integer
For Count = 0 To 49
    If Planet(Count).Owner = Current Then
        picGalaxy.CurrentX = picPlanet(Count).Left + (picPlanet(Count).Width / 2) - 25
        picGalaxy.CurrentY = picPlanet(Count).Top + picPlanet(Count).Height + 15
        picGalaxy.ForeColor = vbYellow
        picGalaxy.Print Str(Current + 1)
    ElseIf Planet(Count).Owner = Other Then
        picGalaxy.CurrentX = picPlanet(Count).Left + (picPlanet(Count).Width / 2) - 25
        picGalaxy.CurrentY = picPlanet(Count).Top + picPlanet(Count).Height + 15
        picGalaxy.ForeColor = vbRed
        picGalaxy.Print Str(Other + 1)
    End If
Next Count

End Sub

Public Sub UpdateAliens()
'set alien planet troop levels
'increase as game progresses

Randomize

Dim i
For i = 0 To 49
  If Planet(i).Owner = Alien Then
    Planet(i).Troops = Planet(i).Troops + Int(Rnd * (Int(TurnNumber / 4))) + 1
    
    'upper limit of troops tied to resources on planet
    Dim UpperLimit As Integer
    UpperLimit = Planet(i).Resources * 25
    If Planet(i).Troops > UpperLimit Then
        'vary amount +/- 5 troops
        Planet(i).Troops = UpperLimit + Int(Rnd * 10) - 5
        'MsgBox "resources = " + Str(Planet(i).Resources) + Chr(13) + "troops = " + Str(Planet(i).Troops), , Planet(i).Name
    End If
    
    'set mechs if enough troops
    If Planet(i).Troops > 50 Then
        Planet(i).AssaultTroops = Int(Rnd * 8) + 1
    ElseIf Planet(i).Troops > 30 Then
        Planet(i).AssaultTroops = Int(Rnd * 5) + 1
    ElseIf Planet(i).Troops > 20 Then
        Planet(i).AssaultTroops = Int(Rnd * 3) + 1
    End If
             
    SetCombatStrength (i)
  End If
Next i
End Sub

Public Sub LostInSpace()
'count number of planets available for landing
'if count is zero, tell player their ship is lost
'then reinitialize the ship etc.

If Player(Current).Ship(0).Launched Then
    Dim Count As Integer
    Dim x1, y1, X2, Y2
    Dim a As Integer
    Dim b As Integer
    Dim Distance
    Dim RangeLow, RangeHigh
    Dim xpos, ypos, radius
    
    'set activeship to appropriate ship number
    activeship = 0
    
    'clear the board of any inrange settings
    Dim z
    For z = 0 To 49
        Planet(z).InRange = False
    Next z

    'Check for UltraWarp and set ranges
    If Player(Current).UltraWarpResearched Then
        'increased range
        RangeLow = Player(Current).Ship(0).WarpPosition * 250
        RangeHigh = RangeLow + 700
    ElseIf Player(Current).UltraWarpResearched = False Then
        'lower ranges
        RangeLow = Player(Current).Ship(0).WarpPosition * 250
        RangeHigh = RangeLow + 350
    End If
  
    'ship's starting position - originating planet
    x1 = Player(Current).Ship(0).CenterX
    y1 = Player(Current).Ship(0).CenterY
    
    'check distance from home planet to each other planet
    'if within the range, set planet's InRange to true
    For Count = 0 To 49
        X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
        a = Abs(x1 - X2)
        b = Abs(y1 - Y2)
   
        Distance = Int(Sqr(a ^ 2 + b ^ 2))
        If Distance >= RangeLow And Distance <= RangeHigh And picPlanet(Count).Visible Then
            'planet is within range - add to list
            NumPlanets1 = NumPlanets1 + 1
        End If
    Next Count
Else
    NumPlanets1 = 999
End If

If Player(Current).Ship(1).Launched Then
    'set activeship to appropriate ship number
    activeship = 0
    
    'clear the board of any inrange settings
    For z = 0 To 49
        Planet(z).InRange = False
    Next z

    'Check for UltraWarp and set ranges
    If Player(Current).UltraWarpResearched Then
        'increased range
        RangeLow = Player(Current).Ship(1).WarpPosition * 250
        RangeHigh = RangeLow + 700
    ElseIf Player(Current).UltraWarpResearched = False Then
        'lower ranges
        RangeLow = Player(Current).Ship(1).WarpPosition * 250
        RangeHigh = RangeLow + 350
    End If
  
    'ship's starting position - originating planet
    x1 = Player(Current).Ship(1).CenterX
    y1 = Player(Current).Ship(1).CenterY
    
    'check distance from home planet to each other planet
    'if within the range, set planet's InRange to true
    For Count = 0 To 49
        X2 = picPlanet(Count).Left + (picPlanet(Count).Width / 2)
        Y2 = picPlanet(Count).Top + (picPlanet(Count).Height / 2)
    
        a = Abs(x1 - X2)
        b = Abs(y1 - Y2)
   
        Distance = Int(Sqr(a ^ 2 + b ^ 2))
        If Distance >= RangeLow And Distance <= RangeHigh And picPlanet(Count).Visible Then
            'planet is within range - add to list
            NumPlanets2 = NumPlanets2 + 1
        End If
    Next Count
Else
    NumPlanets2 = 999
End If
'MsgBox "Number of eligible planets:" + Str(NumPlanets1)

End Sub

Public Sub DrawGridLines()
'draw grid lines

Dim x1, X2, x3, x4
Dim y1, Y2, y3, y4
Dim linecolor, lineheight, linewidth

linecolor = &H808080
lineheight = picGalaxy.ScaleHeight
linewidth = picGalaxy.ScaleWidth

'vertical lines
x1 = picGalaxy.ScaleWidth / 5
picGalaxy.Line (x1, 0)-(x1, lineheight), linecolor
X2 = x1 * 2
picGalaxy.Line (X2, 0)-(X2, lineheight), linecolor
x3 = x1 * 3
picGalaxy.Line (x3, 0)-(x3, lineheight), linecolor
x4 = x1 * 4
picGalaxy.Line (x4, 0)-(x4, lineheight), linecolor

'horizontal lines
y1 = picGalaxy.ScaleHeight / 5
picGalaxy.Line (0, y1)-(linewidth, y1), linecolor
Y2 = y1 * 2
picGalaxy.Line (0, Y2)-(linewidth, Y2), linecolor
y3 = y1 * 3
picGalaxy.Line (0, y3)-(linewidth, y3), linecolor
y4 = y1 * 4
picGalaxy.Line (0, y4)-(linewidth, y4), linecolor

End Sub

Public Sub EraseGridLines()
'draw grid lines

Dim x1, X2, x3, x4
Dim y1, Y2, y3, y4
Dim lineheight, linewidth

lineheight = picGalaxy.ScaleHeight
linewidth = picGalaxy.ScaleWidth

'vertical lines
x1 = picGalaxy.ScaleWidth / 5
picGalaxy.Line (x1, 0)-(x1, lineheight), vbBlack
X2 = x1 * 2
picGalaxy.Line (X2, 0)-(X2, lineheight), vbBlack
x3 = x1 * 3
picGalaxy.Line (x3, 0)-(x3, lineheight), vbBlack
x4 = x1 * 4
picGalaxy.Line (x4, 0)-(x4, lineheight), vbBlack

'horizontal lines
y1 = picGalaxy.ScaleHeight / 5
picGalaxy.Line (0, y1)-(linewidth, y1), vbBlack
Y2 = y1 * 2
picGalaxy.Line (0, Y2)-(linewidth, Y2), vbBlack
y3 = y1 * 3
picGalaxy.Line (0, y3)-(linewidth, y3), vbBlack
y4 = y1 * 4
picGalaxy.Line (0, y4)-(linewidth, y4), vbBlack

End Sub

Public Sub AlienExpansion()
'Alien Expansion procedure
'***once aliens planets have at least 20 troops, then look at the 3 planets on either side.
'if they're neutral, there is a 5% + 1-TurnNumber chance of expanding

Dim X As Integer
Randomize

For X = 0 To 49
    If Planet(X).Owner = Alien And Planet(X).Troops > 20 Then
        'look at planets +/- 3 of the alien planet
        Dim CheckA As Integer
        Dim CheckZ As Integer
        Dim Result As Integer
        Dim Y As Integer
        Dim ChanceOfExpansion As Integer
        
        ChanceOfExpansion = 5 + (Int(Rnd * TurnNumber) + 1)
        
        CheckA = X - 3
        If CheckA < 0 Then CheckA = 0  'prevent error 9 - subscript out of range
        
        CheckZ = X + 3
        If CheckZ > 49 Then CheckZ = 49  'ditto for error 9
        
        For Y = CheckA To CheckZ
            If Planet(Y).Owner = Neutral Then
                Result = Int(Rnd * 100) + 1
                If Result <= ChanceOfExpansion Then
                    'aliens take planet
                    Dim Force As Integer        'how many aliens invade
                    Force = Int(Rnd * 5) + 5
                    Planet(Y).Owner = 3
                    Planet(Y).Troops = Force
                    Planet(X).Troops = Planet(X).Troops - Force
                    Exit Sub
                End If
            End If
        Next Y
    End If
Next X

End Sub
