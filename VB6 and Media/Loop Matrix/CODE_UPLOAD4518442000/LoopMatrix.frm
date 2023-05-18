VERSION 5.00
Begin VB.Form Matrix 
   BackColor       =   &H80000012&
   Caption         =   "Loop Matrix"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   Icon            =   "LoopMatrix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   1185
      LargeChange     =   10
      Left            =   9765
      Max             =   0
      Min             =   100
      TabIndex        =   323
      Top             =   6255
      Width           =   240
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Normal Pan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6660
      Picture         =   "LoopMatrix.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   322
      ToolTipText     =   "Sets all pan levels to 50%."
      Top             =   6525
      Width           =   1410
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Random Pan"
      DownPicture     =   "LoopMatrix.frx":0E01
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6660
      Picture         =   "LoopMatrix.frx":18F8
      Style           =   1  'Graphical
      TabIndex        =   321
      ToolTipText     =   "Random panning for Channel strips with samples assighned."
      Top             =   6210
      Width           =   1410
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   4005
      TabIndex        =   190
      ToolTipText     =   "Loaded Samples."
      Top             =   6210
      Width           =   2580
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2790
      Top             =   6210
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00000000&
      Height          =   1545
      Left            =   90
      Picture         =   "LoopMatrix.frx":23EF
      ScaleHeight     =   1485
      ScaleWidth      =   11520
      TabIndex        =   29
      Top             =   6165
      Width           =   11580
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0C000&
         Caption         =   "Reset Matrix"
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
         Left            =   10035
         Picture         =   "LoopMatrix.frx":3BB6
         Style           =   1  'Graphical
         TabIndex        =   319
         ToolTipText     =   "Resets the Matrix"
         Top             =   1080
         Width           =   1410
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         Columns         =   2
         ForeColor       =   &H0000FFFF&
         Height          =   1425
         Index           =   4
         ItemData        =   "LoopMatrix.frx":46AD
         Left            =   45
         List            =   "LoopMatrix.frx":46BD
         TabIndex        =   191
         ToolTipText     =   "Applications directory"
         Top             =   45
         Width           =   3795
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   10
         Left            =   10080
         Max             =   60
         Min             =   180
         TabIndex        =   188
         Top             =   405
         Value           =   60
         Width           =   1365
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0C000&
         Caption         =   "Stop"
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
         Left            =   10755
         Picture         =   "LoopMatrix.frx":46EE
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Stops the Matrix"
         Top             =   720
         Width           =   690
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C000&
         Caption         =   "Play"
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
         Left            =   10035
         Picture         =   "LoopMatrix.frx":51E5
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Plays the Matrix"
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Master"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   0
         Left            =   9495
         TabIndex        =   324
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tempo"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   10080
         TabIndex        =   189
         ToolTipText     =   "Loop Tempo"
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":5CDC
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   3
      Top             =   3105
      Width           =   11580
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   465
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   135
         Width           =   240
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   8
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   89
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   8
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   88
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   8
         ItemData        =   "LoopMatrix.frx":6A6C
         Left            =   0
         List            =   "LoopMatrix.frx":6A7C
         TabIndex        =   87
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1485
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   8
         Left            =   1530
         Picture         =   "LoopMatrix.frx":6AAD
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1845
         TabIndex        =   17
         Top             =   405
         Width           =   375
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1845
         TabIndex        =   16
         Top             =   90
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":75A4
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   4
      Top             =   3870
      Width           =   11580
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   420
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   135
         Width           =   240
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   9
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   65
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   9
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   64
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   9
         ItemData        =   "LoopMatrix.frx":8334
         Left            =   45
         List            =   "LoopMatrix.frx":8344
         TabIndex        =   63
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   345
         Index           =   9
         Left            =   1530
         Picture         =   "LoopMatrix.frx":8375
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1845
         TabIndex        =   15
         Top             =   405
         Width           =   375
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1845
         TabIndex        =   14
         Top             =   90
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":8E6C
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   5
      Top             =   4635
      Width           =   11580
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   420
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   135
         Width           =   240
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   10
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   69
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   10
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   68
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   345
         Index           =   10
         Left            =   1530
         Picture         =   "LoopMatrix.frx":9BFC
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   270
         Width           =   240
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   10
         ItemData        =   "LoopMatrix.frx":A6F3
         Left            =   45
         List            =   "LoopMatrix.frx":A703
         TabIndex        =   66
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   232
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   135
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   1845
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   1845
         TabIndex        =   12
         Top             =   90
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":A734
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   6
      Top             =   5400
      Width           =   11580
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   11
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   28
         Top             =   90
         Value           =   50
         Width           =   1740
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   11
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   27
         Top             =   405
         Value           =   75
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   11
         Left            =   1530
         Picture         =   "LoopMatrix.frx":B4C4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   11
         ItemData        =   "LoopMatrix.frx":BFBB
         Left            =   45
         List            =   "LoopMatrix.frx":BFCB
         TabIndex        =   25
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   420
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   320
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   242
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   244
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   247
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   248
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   249
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   250
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   252
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   135
         Width           =   240
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   1845
         TabIndex        =   11
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   1845
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture0 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   45
      Picture         =   "LoopMatrix.frx":BFFC
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   7
      Top             =   45
      Width           =   11580
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   318
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   317
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   316
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   315
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   314
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   313
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   312
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   311
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   310
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   309
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   308
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   307
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   306
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   305
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   304
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   303
         Top             =   135
         Width           =   230
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   0
         ItemData        =   "LoopMatrix.frx":CD8C
         Left            =   45
         List            =   "LoopMatrix.frx":CD9C
         TabIndex        =   169
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1470
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   0
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   168
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   0
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   167
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   0
         Left            =   1575
         Picture         =   "LoopMatrix.frx":CDCD
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   135
         Width           =   230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   1845
         TabIndex        =   9
         Top             =   405
         Width           =   375
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   1845
         TabIndex        =   8
         Top             =   90
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":D8C4
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   2
      Top             =   2340
      Width           =   11580
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   269
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   135
         Width           =   225
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   268
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   266
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   265
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   263
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   262
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   261
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   260
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   259
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   258
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   257
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   256
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   135
         Width           =   225
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   3
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   109
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   3
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   108
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   3
         Left            =   1530
         Picture         =   "LoopMatrix.frx":E654
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   3
         ItemData        =   "LoopMatrix.frx":F14B
         Left            =   0
         List            =   "LoopMatrix.frx":F15B
         TabIndex        =   106
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   270
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   135
         Width           =   230
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1845
         TabIndex        =   24
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1845
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":F18C
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   18
      Top             =   810
      Width           =   11580
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   301
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   300
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   298
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   297
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   296
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   295
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   294
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   293
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   291
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   290
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   289
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   288
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   287
         Top             =   135
         Width           =   230
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   1
         ItemData        =   "LoopMatrix.frx":FF1C
         Left            =   45
         List            =   "LoopMatrix.frx":FF2C
         TabIndex        =   149
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1440
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   1
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   148
         Top             =   45
         Value           =   75
         Width           =   1740
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   1
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   147
         Top             =   360
         Value           =   50
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   1
         Left            =   1530
         Picture         =   "LoopMatrix.frx":FF5D
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   135
         Width           =   230
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   1845
         TabIndex        =   20
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   1845
         TabIndex        =   19
         Top             =   405
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C000C0&
      Height          =   735
      Left            =   90
      Picture         =   "LoopMatrix.frx":10A54
      ScaleHeight     =   675
      ScaleWidth      =   11520
      TabIndex        =   1
      Top             =   1575
      Width           =   11580
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   31
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   286
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   30
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   285
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   29
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   284
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   28
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   283
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   27
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   282
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   281
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   280
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   24
         Left            =   9675
         Style           =   1  'Graphical
         TabIndex        =   279
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   9450
         Style           =   1  'Graphical
         TabIndex        =   278
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   277
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   276
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   20
         Left            =   8775
         Style           =   1  'Graphical
         TabIndex        =   275
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   274
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   8325
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   7875
         Style           =   1  'Graphical
         TabIndex        =   271
         Top             =   135
         Width           =   230
      End
      Begin VB.HScrollBar hsbPan 
         Height          =   195
         Index           =   2
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   129
         Top             =   405
         Value           =   50
         Width           =   1740
      End
      Begin VB.CommandButton cmdBuffer 
         Caption         =   "T"
         Height          =   300
         Index           =   2
         Left            =   1530
         Picture         =   "LoopMatrix.frx":117E4
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Listen to Sample in the buffer"
         Top             =   315
         Width           =   240
      End
      Begin VB.HScrollBar hsbVolume 
         Height          =   195
         Index           =   2
         LargeChange     =   5
         Left            =   2430
         Max             =   100
         TabIndex        =   127
         Top             =   90
         Value           =   75
         Width           =   1740
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   645
         Index           =   2
         ItemData        =   "LoopMatrix.frx":122DB
         Left            =   0
         List            =   "LoopMatrix.frx":122EB
         TabIndex        =   126
         ToolTipText     =   "Window for selecting Loaded Samples."
         Top             =   0
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   5175
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   6525
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   135
         Width           =   230
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   135
         Width           =   230
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   1845
         TabIndex        =   22
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pan"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   1845
         TabIndex        =   21
         Top             =   405
         Width           =   375
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   31
      Visible         =   0   'False
      X1              =   11475
      X2              =   11475
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   30
      Visible         =   0   'False
      X1              =   11250
      X2              =   11250
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   29
      Visible         =   0   'False
      X1              =   11025
      X2              =   11025
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   28
      Visible         =   0   'False
      X1              =   10800
      X2              =   10800
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   27
      Visible         =   0   'False
      X1              =   10575
      X2              =   10575
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   26
      Visible         =   0   'False
      X1              =   10350
      X2              =   10350
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   25
      Visible         =   0   'False
      X1              =   10125
      X2              =   10125
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   24
      Visible         =   0   'False
      X1              =   9900
      X2              =   9900
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   23
      Visible         =   0   'False
      X1              =   9675
      X2              =   9675
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   22
      Visible         =   0   'False
      X1              =   9450
      X2              =   9450
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   20
      Visible         =   0   'False
      X1              =   9000
      X2              =   9000
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   19
      Visible         =   0   'False
      X1              =   8775
      X2              =   8775
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   18
      Visible         =   0   'False
      X1              =   8550
      X2              =   8550
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   17
      Visible         =   0   'False
      X1              =   8325
      X2              =   8325
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   16
      Visible         =   0   'False
      X1              =   8100
      X2              =   8100
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   15
      Visible         =   0   'False
      X1              =   7875
      X2              =   7875
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   14
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   13
      Visible         =   0   'False
      X1              =   7425
      X2              =   7425
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   12
      Visible         =   0   'False
      X1              =   7200
      X2              =   7200
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   11
      Visible         =   0   'False
      X1              =   6975
      X2              =   6975
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   10
      Visible         =   0   'False
      X1              =   6750
      X2              =   6750
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   9
      Visible         =   0   'False
      X1              =   6525
      X2              =   6525
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   8
      Visible         =   0   'False
      X1              =   6300
      X2              =   6300
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   7
      Visible         =   0   'False
      X1              =   6075
      X2              =   6075
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   6
      Visible         =   0   'False
      X1              =   5850
      X2              =   5850
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   5
      Visible         =   0   'False
      X1              =   5625
      X2              =   5625
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   4
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   3
      Visible         =   0   'False
      X1              =   5175
      X2              =   5175
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   2
      Visible         =   0   'False
      X1              =   4950
      X2              =   4950
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   1
      Visible         =   0   'False
      X1              =   4725
      X2              =   4725
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   0
      X1              =   4500
      X2              =   4500
      Y1              =   225
      Y2              =   5985
   End
   Begin VB.Label Label2 
      Caption         =   "RS = Random Sound :      RB  = Random Buffer :     RP = Random Pan"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   9675
      Width           =   8775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      DrawMode        =   12  'Nop
      Index           =   21
      Visible         =   0   'False
      X1              =   9270
      X2              =   9225
      Y1              =   180
      Y2              =   5985
   End
End
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const cSoundBuffers = 30

Private Sub cmdBuffer_Click(index As Integer)
If Option1.Value = True Then
  hsbPan(index).Value = (Rnd * 100) '+ 25
  PlaySoundWithPan index, List1(index), hsbVolume(index).Value, hsbPan(index).Value ' Play the buffer for this index
Else
PlaySoundWithPan index, List1(index), hsbVolume(index).Value, hsbPan(index).Value
End If
End Sub

Private Sub Command1_Click(index As Integer)
If Command1(index).BackColor = vbYellow Then
Command1(index).BackColor = vbBlue
Else
Command1(index).BackColor = vbYellow
End If
End Sub

Private Sub Command10_Click(index As Integer)
If Command10(index).BackColor = vbYellow Then
Command10(index).BackColor = vbBlue
Else
Command10(index).BackColor = vbYellow
End If
End Sub

Private Sub Command11_Click(index As Integer)
If Command11(index).BackColor = vbYellow Then
Command11(index).BackColor = vbBlue
Else
Command11(index).BackColor = vbYellow
End If
End Sub

Private Sub Command12_Click()
Timer1.Interval = HScroll1.Value - 59
Timer1.Enabled = True
End Sub

Private Sub Command13_Click()
Timer1.Enabled = False
End Sub



Private Sub Command15_Click()
Dim index As Integer
For index = 0 To 31
Command1(index).BackColor = vbBlue
Command2(index).BackColor = vbBlue
Command3(index).BackColor = vbBlue
Command4(index).BackColor = vbBlue
Command8(index).BackColor = vbBlue
Command9(index).BackColor = vbBlue
Command10(index).BackColor = vbBlue
Command11(index).BackColor = vbBlue
Next index
End Sub

Private Sub Command2_Click(index As Integer)
If Command2(index).BackColor = vbYellow Then
Command2(index).BackColor = vbBlue
Else
Command2(index).BackColor = vbYellow
End If
End Sub

Private Sub Command3_Click(index As Integer)
If Command3(index).BackColor = vbYellow Then
Command3(index).BackColor = vbBlue
Else
Command3(index).BackColor = vbYellow
End If
End Sub

Private Sub Command4_Click(index As Integer)
If Command4(index).BackColor = vbYellow Then
Command4(index).BackColor = vbBlue
Else
Command4(index).BackColor = vbYellow
End If
End Sub





Private Sub Command8_Click(index As Integer)
If Command8(index).BackColor = vbYellow Then
Command8(index).BackColor = vbBlue
Else
Command8(index).BackColor = vbYellow
End If
End Sub

Private Sub Command9_Click(index As Integer)
If Command9(index).BackColor = vbYellow Then
Command9(index).BackColor = vbBlue
Else
Command9(index).BackColor = vbYellow
End If
End Sub

Private Sub File1_Click()
PlaySoundAnyBuffer File1.Filename, 90
End Sub

Private Sub File1_DblClick()
Dim index As Integer
List1(index).AddItem File1.Filename
  List1(0).AddItem File1.Filename
  List1(1).AddItem File1.Filename
  List1(2).AddItem File1.Filename
  List1(3).AddItem File1.Filename
  List1(4).AddItem File1.Filename
  List1(8).AddItem File1.Filename
  List1(9).AddItem File1.Filename
  List1(10).AddItem File1.Filename
  List1(11).AddItem File1.Filename
List1(index).Refresh
End Sub

Private Sub Form_Load()
frmSplash.Show vbModal
File1.Path = App.Path & "\Sound"
File1.Pattern = "*.wav"
HScroll1.Value = 100
'This Code was created by D.R Hall
'For more Information and latest version
'E-mail me, derek.hall@virgin.net
'To set up the DX7 sound module, just call these routines 3 routines
  
  SetupDX7Sound Me              ' Assign DX7 to this Application
  SoundDir App.Path & "\Sound"  'Where is the applications sound stored

  CreateBuffers cSoundBuffers, "Beat1.wav" ' How many Channels/buffers (I used 10)
                                              'and assign a default sound,
                                              'change sound later.
                                              'To stop errors you must supply
                                              'a default wave to set up a buffer
                                              'select your smallest wave file
'**** Make a selection for each listbox on this form
  List1(0).ListIndex = 1
  List1(1).ListIndex = 3
  List1(2).ListIndex = 2
  List1(3).ListIndex = 2
  List1(8).ListIndex = 0
  List1(9).ListIndex = 1
  List1(10).ListIndex = 3
  List1(11).ListIndex = 0
'**************************************
End Sub

Private Sub hsbPan_Change(index As Integer)
  PanSound index, hsbPan(index).Value 'value must be 0 to 100, 50 is centered
End Sub

Private Sub hsbVolume_Change(index As Integer)
  VolumeLevel index, hsbVolume(index).Value 'value must be 0 to 100, 0 no sound,
End Sub

Private Sub HScroll1_Change()
Label1(0).Caption = "Tempo : " & HScroll1.Min - HScroll1.Value + 60
Dim a As Integer
a = HScroll1.Value - 59
Timer1.Interval = a
End Sub

Private Sub HScroll1_Scroll()
Label1(0).Caption = "Tempo : " & HScroll1.Min - HScroll1.Value + 60
Dim a As Integer
a = HScroll1.Value - 59
Timer1.Interval = a
End Sub

Private Sub Option2_Click()
hsbPan(0).Value = 50
hsbPan(1).Value = 50
hsbPan(2).Value = 50
hsbPan(3).Value = 50
hsbPan(8).Value = 50
hsbPan(9).Value = 50
hsbPan(10).Value = 50
hsbPan(11).Value = 50
End Sub

Private Sub Timer1_Timer()
Select Case True
Case Line1(0).Visible = True
Line1(0).Visible = False
Line1(1).Visible = True
highlight (0)
Case Line1(1).Visible = True
Line1(1).Visible = False
Line1(2).Visible = True
highlight (1)
Case Line1(2).Visible = True
Line1(2).Visible = False
Line1(3).Visible = True
highlight (2)
Case Line1(3).Visible = True
Line1(3).Visible = False
Line1(4).Visible = True
highlight (3)
Case Line1(4).Visible = True
Line1(4).Visible = False
Line1(5).Visible = True
highlight (4)
Case Line1(5).Visible = True
Line1(5).Visible = False
Line1(6).Visible = True
highlight (5)
Case Line1(6).Visible = True
Line1(6).Visible = False
Line1(7).Visible = True
highlight (6)
Case Line1(7).Visible = True
Line1(7).Visible = False
Line1(8).Visible = True
highlight (7)
Case Line1(8).Visible = True
Line1(8).Visible = False
Line1(9).Visible = True
highlight (8)
Case Line1(9).Visible = True
Line1(9).Visible = False
Line1(10).Visible = True
highlight (9)
Case Line1(10).Visible = True
Line1(10).Visible = False
Line1(11).Visible = True
highlight (10)
Case Line1(11).Visible = True
Line1(11).Visible = False
Line1(12).Visible = True
highlight (11)
Case Line1(12).Visible = True
Line1(12).Visible = False
Line1(13).Visible = True
highlight (12)
Case Line1(13).Visible = True
Line1(13).Visible = False
Line1(14).Visible = True
highlight (13)
Case Line1(14).Visible = True
Line1(14).Visible = False
Line1(15).Visible = True
highlight (14)
Case Line1(15).Visible = True
Line1(15).Visible = False
Line1(16).Visible = True
highlight (15)
Case Line1(16).Visible = True
Line1(16).Visible = False
Line1(17).Visible = True
highlight (16)
Case Line1(17).Visible = True
Line1(17).Visible = False
Line1(18).Visible = True
highlight (17)
Case Line1(18).Visible = True
Line1(18).Visible = False
Line1(19).Visible = True
highlight (18)
Case Line1(19).Visible = True
Line1(19).Visible = False
Line1(20).Visible = True
highlight (19)
Case Line1(20).Visible = True
Line1(20).Visible = False
Line1(21).Visible = True
highlight (20)
Case Line1(21).Visible = True
Line1(21).Visible = False
Line1(22).Visible = True
highlight (21)
Case Line1(22).Visible = True
Line1(22).Visible = False
Line1(23).Visible = True
highlight (22)
Case Line1(23).Visible = True
Line1(23).Visible = False
Line1(24).Visible = True
highlight (23)
Case Line1(24).Visible = True
Line1(24).Visible = False
Line1(25).Visible = True
highlight (24)
Case Line1(25).Visible = True
Line1(25).Visible = False
Line1(26).Visible = True
highlight (25)
Case Line1(26).Visible = True
Line1(26).Visible = False
Line1(27).Visible = True
highlight (26)
Case Line1(27).Visible = True
Line1(27).Visible = False
Line1(28).Visible = True
highlight (27)
Case Line1(28).Visible = True
Line1(28).Visible = False
Line1(29).Visible = True
highlight (28)
Case Line1(29).Visible = True
Line1(29).Visible = False
Line1(30).Visible = True
highlight (29)
Case Line1(30).Visible = True
Line1(30).Visible = False
Line1(31).Visible = True
highlight (30)
Case Line1(31).Visible = True
Line1(31).Visible = False
Line1(0).Visible = True
highlight (31)
End Select
End Sub

Function highlight(index As Integer)
If Command1(index).BackColor = vbYellow Then
cmdBuffer_Click (0)
End If
If Command2(index).BackColor = vbYellow Then
cmdBuffer_Click (1)
End If
If Command3(index).BackColor = vbYellow Then
cmdBuffer_Click (2)
End If
If Command4(index).BackColor = vbYellow Then
cmdBuffer_Click (3)
End If
If Command8(index).BackColor = vbYellow Then
cmdBuffer_Click (8)
End If
If Command9(index).BackColor = vbYellow Then
cmdBuffer_Click (9)
End If
If Command10(index).BackColor = vbYellow Then
cmdBuffer_Click (10)
End If
If Command11(index).BackColor = vbYellow Then
cmdBuffer_Click (11)
End If
End Function

Private Sub VScroll1_Change()
Dim index As Integer
hsbVolume(0) = VScroll1.Value
hsbVolume(1) = VScroll1.Value
hsbVolume(2) = VScroll1.Value
hsbVolume(3) = VScroll1.Value
hsbVolume(8) = VScroll1.Value
hsbVolume(9) = VScroll1.Value
hsbVolume(10) = VScroll1.Value
hsbVolume(11) = VScroll1.Value
End Sub
