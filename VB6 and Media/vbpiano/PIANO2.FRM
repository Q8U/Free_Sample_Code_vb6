VERSION 4.00
Begin VB.Form Piano 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piano"
   ClientHeight    =   2385
   ClientLeft      =   75
   ClientTop       =   1845
   ClientWidth     =   9510
   ClipControls    =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   3120
   Icon            =   "piano2.frx":0000
   Left            =   15
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   9510
   Top             =   1170
   Width           =   9630
   Begin Threed.SSPanel Panel3D3 
      Height          =   975
      Left            =   90
      TabIndex        =   20
      Top             =   1290
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   1720
      _StockProps     =   15
      ForeColor       =   8421504
      BevelWidth      =   3
      BorderWidth     =   0
      Outline         =   -1  'True
      Font3D          =   3
      Alignment       =   6
      Autosize        =   3
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":030A
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   61
         Left            =   8640
         TabIndex        =   82
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":0614
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   63
         Left            =   8910
         TabIndex        =   81
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":091E
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   1
         Left            =   240
         TabIndex        =   73
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":0C28
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   3
         Left            =   480
         TabIndex        =   72
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":0F32
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   6
         Left            =   960
         TabIndex        =   71
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":123C
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   8
         Left            =   1200
         TabIndex        =   70
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":1546
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   10
         Left            =   1440
         TabIndex        =   69
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":1850
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   13
         Left            =   1920
         TabIndex        =   40
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":1B5A
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   15
         Left            =   2160
         TabIndex        =   39
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":1E64
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   20
         Left            =   2880
         TabIndex        =   38
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":216E
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   22
         Left            =   3120
         TabIndex        =   37
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":2478
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   25
         Left            =   3600
         TabIndex        =   36
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":2782
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   27
         Left            =   3840
         TabIndex        =   35
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":2A8C
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   32
         Left            =   4560
         TabIndex        =   34
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":2D96
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   34
         Left            =   4800
         TabIndex        =   33
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":30A0
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   37
         Left            =   5280
         TabIndex        =   32
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":33AA
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   39
         Left            =   5520
         TabIndex        =   31
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":36B4
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   44
         Left            =   6240
         TabIndex        =   30
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":39BE
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   46
         Left            =   6480
         TabIndex        =   29
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":3CC8
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   49
         Left            =   6960
         TabIndex        =   28
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":3FD2
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   51
         Left            =   7200
         TabIndex        =   27
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":42DC
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   56
         Left            =   7920
         TabIndex        =   26
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":45E6
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   58
         Left            =   8160
         TabIndex        =   25
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":48F0
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   18
         Left            =   2640
         TabIndex        =   24
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":4BFA
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   30
         Left            =   4320
         TabIndex        =   23
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":4F04
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   42
         Left            =   6000
         TabIndex        =   22
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":520E
         DragMode        =   1  'Automatic
         Height          =   585
         Index           =   54
         Left            =   7680
         TabIndex        =   21
         Top             =   90
         Width           =   165
         _Version        =   65536
         _ExtentX        =   291
         _ExtentY        =   1032
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   0
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":5518
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   60
         Left            =   8490
         TabIndex        =   85
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":5822
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   62
         Left            =   8730
         TabIndex        =   84
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":5B2C
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   64
         Left            =   8970
         TabIndex        =   83
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":5E36
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   7
         Left            =   1050
         TabIndex        =   80
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":6140
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   9
         Left            =   1290
         TabIndex        =   79
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":644A
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   11
         Left            =   1530
         TabIndex        =   78
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":6754
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   4
         Left            =   570
         TabIndex        =   77
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":6A5E
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   5
         Left            =   810
         TabIndex        =   76
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":6D68
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   2
         Left            =   330
         TabIndex        =   75
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":7072
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   0
         Left            =   90
         TabIndex        =   74
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":737C
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   12
         Left            =   1770
         TabIndex        =   68
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":7686
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   14
         Left            =   2010
         TabIndex        =   67
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":7990
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   16
         Left            =   2250
         TabIndex        =   66
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":7C9A
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   17
         Left            =   2490
         TabIndex        =   65
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":7FA4
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   19
         Left            =   2730
         TabIndex        =   64
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":82AE
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   21
         Left            =   2970
         TabIndex        =   63
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":85B8
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   23
         Left            =   3210
         TabIndex        =   62
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":88C2
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   24
         Left            =   3450
         TabIndex        =   61
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":8BCC
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   26
         Left            =   3690
         TabIndex        =   60
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":8ED6
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   28
         Left            =   3930
         TabIndex        =   59
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":91E0
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   29
         Left            =   4170
         TabIndex        =   58
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":94EA
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   31
         Left            =   4410
         TabIndex        =   57
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":97F4
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   33
         Left            =   4650
         TabIndex        =   56
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":9AFE
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   35
         Left            =   4890
         TabIndex        =   55
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":9E08
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   36
         Left            =   5130
         TabIndex        =   54
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":A112
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   38
         Left            =   5370
         TabIndex        =   53
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":A41C
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   40
         Left            =   5610
         TabIndex        =   52
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":A726
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   41
         Left            =   5850
         TabIndex        =   51
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":AA30
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   43
         Left            =   6090
         TabIndex        =   50
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":AD3A
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   45
         Left            =   6330
         TabIndex        =   49
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":B044
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   47
         Left            =   6570
         TabIndex        =   48
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":B34E
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   48
         Left            =   6810
         TabIndex        =   47
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":B658
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   50
         Left            =   7050
         TabIndex        =   46
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":B962
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   52
         Left            =   7290
         TabIndex        =   45
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":BC6C
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   53
         Left            =   7530
         TabIndex        =   44
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":BF76
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   55
         Left            =   7770
         TabIndex        =   43
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":C280
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   57
         Left            =   8010
         TabIndex        =   42
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
      Begin Threed.SSPanel PanelWhite 
         DragIcon        =   "piano2.frx":C58A
         DragMode        =   1  'Automatic
         Height          =   855
         Index           =   59
         Left            =   8250
         TabIndex        =   41
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   1508
         _StockProps     =   15
         ForeColor       =   8421504
         BackColor       =   16777215
         BevelWidth      =   2
         BorderWidth     =   2
         Outline         =   -1  'True
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1245
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
      _ExtentY        =   2196
      _StockProps     =   14
      Caption         =   "Piano Settings"
      ForeColor       =   0
      Alignment       =   2
      ShadowColor     =   1
      Font3D          =   3
      ShadowStyle     =   1
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Left            =   2070
         TabIndex        =   18
         Top             =   720
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SSPanel6"
         ForeColor       =   8421504
         BorderWidth     =   2
         BevelInner      =   1
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
         Begin VB.HScrollBar HScrollVolume 
            Height          =   255
            LargeChange     =   10
            Left            =   60
            Max             =   127
            TabIndex        =   19
            Top             =   60
            Value           =   50
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   5520
         TabIndex        =   16
         Top             =   720
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SSPanel5"
         ForeColor       =   8421504
         BorderWidth     =   1
         BevelInner      =   1
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
         Begin VB.HScrollBar HScrollPatch 
            Height          =   285
            LargeChange     =   10
            Left            =   45
            Max             =   127
            TabIndex        =   17
            Top             =   45
            Value           =   1
            Width           =   2325
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SSPanel2"
         ForeColor       =   8421504
         BorderWidth     =   1
         BevelInner      =   1
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
         Begin VB.HScrollBar HScrollMIDIChannel 
            Height          =   285
            Left            =   45
            Max             =   15
            TabIndex        =   15
            Top             =   45
            Value           =   1
            Width           =   1590
         End
      End
      Begin Threed.SSPanel VolumeLabel 
         Height          =   225
         Left            =   3345
         TabIndex        =   13
         Top             =   390
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "100"
         ForeColor       =   0
         BorderWidth     =   1
         BevelOuter      =   0
         Font3D          =   3
         Alignment       =   4
         Autosize        =   3
      End
      Begin Threed.SSPanel MidiChannelOutLabel 
         Height          =   225
         Left            =   1575
         TabIndex        =   12
         Top             =   390
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "1"
         ForeColor       =   0
         BorderWidth     =   1
         BevelOuter      =   0
         Font3D          =   3
         Alignment       =   4
         Autosize        =   3
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   375
         Left            =   8100
         TabIndex        =   11
         Top             =   300
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "L -  Pan  - R"
         ForeColor       =   0
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         FloodShowPct    =   0   'False
         Font3D          =   3
         Autosize        =   3
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   375
         Index           =   0
         Left            =   2070
         TabIndex        =   10
         Top             =   300
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Volume"
         ForeColor       =   0
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         FloodShowPct    =   0   'False
         Font3D          =   3
         Alignment       =   1
         Autosize        =   3
      End
      Begin Threed.SSPanel PatchLabel 
         Height          =   375
         Left            =   6150
         TabIndex        =   9
         Top             =   300
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Electric Piano 2 "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            name            =   "Small Fonts"
            charset         =   0
            weight          =   700
            size            =   6
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         Font3D          =   3
         Autosize        =   3
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   375
         Left            =   5535
         TabIndex        =   8
         Top             =   300
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Patch"
         ForeColor       =   0
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         Font3D          =   3
         Alignment       =   1
         Autosize        =   3
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "MIDI Channel "
         ForeColor       =   0
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         Font3D          =   3
         Alignment       =   1
         Autosize        =   3
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   375
         Left            =   8100
         TabIndex        =   5
         Top             =   720
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SSPanel6"
         ForeColor       =   8421504
         BorderWidth     =   2
         BevelInner      =   1
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
         Begin VB.HScrollBar HScrollPan 
            Height          =   255
            LargeChange     =   10
            Left            =   60
            Max             =   127
            TabIndex        =   6
            Top             =   60
            Width           =   1095
         End
      End
      Begin Threed.SSPanel Panel3D2 
         Height          =   375
         Left            =   4050
         TabIndex        =   3
         Top             =   720
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SSPanel6"
         ForeColor       =   8421504
         BorderWidth     =   2
         BevelInner      =   1
         Font3D          =   3
         Alignment       =   6
         Autosize        =   3
         Begin VB.HScrollBar HScrollOctave 
            Height          =   255
            Left            =   60
            Max             =   4
            TabIndex        =   4
            Top             =   60
            Value           =   2
            Width           =   1080
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   375
         Index           =   1
         Left            =   4050
         TabIndex        =   2
         Top             =   300
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Octave"
         ForeColor       =   0
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
         FloodShowPct    =   0   'False
         Font3D          =   3
         Alignment       =   1
         Autosize        =   3
      End
      Begin Threed.SSPanel LabelOctave 
         Height          =   225
         Left            =   4950
         TabIndex        =   1
         Top             =   390
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "1"
         ForeColor       =   0
         BorderWidth     =   1
         BevelOuter      =   0
         Font3D          =   3
         Alignment       =   4
         Autosize        =   3
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Piano"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim NoteCatchCount As Integer
Dim NoteOnCatcher(1024) As Integer


Private Sub About_Click()
    AboutBox1.Show MODAL
End Sub

Private Sub Exit_Click()
   X% = midiOutClose(hMidiOutCopy)
   End
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Piano.Left = 0
   Piano.Top = 0
    
   ' Open Midi Driver
   MidiOutOpenPort
    
   HScrollMIDIChannel.Value = 13
   HScrollPatch.Value = 0
   HScrollVolume.Value = 100
   HScrollPan.Value = 64
   HScrollOctave.Value = 2
   Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   X% = midiOutClose(hMidiOutCopy)
   End
End Sub

Private Sub HScrollMIDIChannel_Change()
   ' Change Midi Channel to Vscroll1 value
   MidiChannelOut = HScrollMIDIChannel.Value

   ' Display new channel
   MidiChannelOutLabel.Caption = Str$(MidiChannelOut + 1)

   ' Sets the Patch & Volume for the current Midi Channel Out
   HScrollPatch.Value = MidiPatch(MidiChannelOut)
   HScrollVolume.Value = MidiVolume(MidiChannelOut)
   HScrollPan.Value = MidiPan(MidiChannelOut)
   HScrollOctave.Value = Octave(MidiChannelOut) / 12
End Sub

Private Sub HScrollOctave_Change()
   LabelOctave.Caption = Str$(HScrollOctave.Value)
   Octave(MidiChannelOut) = (HScrollOctave.Value * 12)
End Sub

Private Sub HScrollPan_Change()
   MidiPan(MidiChannelOut) = HScrollPan.Value

   ' 05-16-92 Pan Midi Out routine
   MidiEventOut = 176 + MidiChannelOut
   MidiNoteOut = 10
   MidiVelOut = MidiPan(MidiChannelOut)
   SendMidiOut
End Sub

Private Sub HScrollPatch_Change()
   ' Sets the Patch for the current Midi Channel Out
   MidiPatch(MidiChannelOut) = HScrollPatch.Value
   ReadPatch

   ' 05-15-92 Patch Midi Out routine
   MidiEventOut = &HC0 + MidiChannelOut
   MidiNoteOut = MidiPatch(MidiChannelOut)
   MidiVelOut = 0
   SendMidiOut
End Sub

Private Sub HScrollVolume_Change()
   MidiVelocity = HScrollVolume.Value
   MidiVolume(MidiChannelOut) = HScrollVolume.Value
   VolumeLabel.Caption = Str$(MidiVelocity)
End Sub

Private Sub PanelWhite_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   For nn = 0 To NoteCatchCount - 1
      MidiEventOut = 144 + MidiChannelOut
      MidiVelOut = 0
      MidiNoteOut = NoteOnCatcher(nn)
      SendMidiOut
      Piano.PanelWhite(NoteOnCatcher(nn) - Octave(MidiChannelOut)).BevelOuter = 2
   Next nn
   NoteCatchCount = 0
End Sub

Private Sub PanelWhite_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
   'If still on same note, discard
   If NoteCatchCount > 0 Then
      If NoteOnCatcher(NoteCatchCount - 1) = Index + Octave(MidiChannelOut) Then
         Exit Sub
      End If
   End If

   Piano.PanelWhite(Index).BevelOuter = 0
    
   MidiEventOut = 144 + MidiChannelOut
   MidiVelOut = MidiVelocity
   MidiNoteOut = Index + Octave(MidiChannelOut)
   SendMidiOut

   'Since drag/drop is being used, we must keep track of the note being played.
   NoteOnCatcher(NoteCatchCount) = MidiNoteOut
   If NoteCatchCount < 750 Then 'Don't let array get out of range
      NoteCatchCount = NoteCatchCount + 1
   End If
End Sub

