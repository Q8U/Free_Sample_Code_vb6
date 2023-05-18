VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTagger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volure's ID3 Machine"
   ClientHeight    =   7725
   ClientLeft      =   1605
   ClientTop       =   -1740
   ClientWidth     =   7335
   LinkMode        =   1  'Source
   LinkTopic       =   "Window"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Choose Folder"
      Height          =   375
      Left            =   5760
      TabIndex        =   74
      Top             =   0
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3195
      Left            =   15
      TabIndex        =   50
      Top             =   4425
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   5636
      SortKey         =   1
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   7937
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4830
      OleObjectBlob   =   "frmMain.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel FolderPath 
      Height          =   255
      Left            =   60
      OleObjectBlob   =   "frmMain.frx":0234
      TabIndex        =   79
      Top             =   60
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Parse Files"
      Height          =   240
      Left            =   2305
      TabIndex        =   47
      Top             =   4155
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Select all"
      Height          =   240
      Left            =   6360
      TabIndex        =   49
      Top             =   4155
      Width           =   960
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Multiple"
      Height          =   240
      Left            =   3330
      TabIndex        =   48
      Top             =   4155
      Width           =   3030
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All SubFolders (Slower)"
      Height          =   240
      Left            =   30
      TabIndex        =   46
      Top             =   4155
      Width           =   2280
   End
   Begin VB.ListBox DirList1 
      Height          =   1860
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":02AE
      Left            =   7185
      List            =   "frmMain.frx":02B0
      TabIndex        =   37
      Top             =   5595
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox DirList2 
      Height          =   1860
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":02B2
      Left            =   8505
      List            =   "frmMain.frx":02B4
      TabIndex        =   38
      Top             =   5595
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9465
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "frmMain.frx":02B6
      Left            =   8505
      List            =   "frmMain.frx":02B8
      TabIndex        =   52
      Top             =   5115
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   255
      ItemData        =   "frmMain.frx":02BA
      Left            =   7185
      List            =   "frmMain.frx":02BC
      TabIndex        =   53
      Top             =   5115
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Respond 
      Interval        =   15000
      Left            =   7185
      Top             =   4635
   End
   Begin VB.Frame frmedia 
      Height          =   220
      Left            =   968
      TabIndex        =   77
      Top             =   5445
      Visible         =   0   'False
      Width           =   5880
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         Height          =   495
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Width           =   5896
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   0
         WindowlessVideo =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   35
      Top             =   7410
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "1:15 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frTab5 
      Caption         =   "Player Control"
      Height          =   3225
      Left            =   120
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   7095
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":02BE
         TabIndex        =   80
         Top             =   420
         Width           =   795
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   6
         Left            =   6480
         TabIndex        =   16
         ToolTipText     =   "Mute"
         Top             =   440
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   5
         Left            =   3120
         TabIndex        =   14
         ToolTipText     =   "Mute"
         Top             =   2840
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "Mute"
         Top             =   2360
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   10
         ToolTipText     =   "Mute"
         Top             =   1880
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Mute"
         Top             =   1380
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   6
         ToolTipText     =   "Mute"
         Top             =   900
         Width           =   200
      End
      Begin VB.CheckBox Mute 
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         ToolTipText     =   "Mute"
         Top             =   440
         Width           =   200
      End
      Begin VB.Frame Frame1 
         Caption         =   "Internal Player"
         Height          =   2295
         Left            =   3600
         TabIndex        =   51
         Top             =   840
         Width           =   3375
         Begin VB.OptionButton PMode 
            Caption         =   "Stop"
            Height          =   390
            Index           =   1
            Left            =   1755
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   810
            Width           =   780
         End
         Begin VB.OptionButton PMode 
            Caption         =   "Play"
            Height          =   390
            Index           =   2
            Left            =   945
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   810
            Width           =   780
         End
         Begin VB.CommandButton Stop 
            Caption         =   "Stop"
            Height          =   390
            Left            =   150
            TabIndex        =   113
            Top             =   810
            Width           =   780
         End
         Begin VB.CheckBox Mute 
            Caption         =   "Mute"
            Height          =   375
            Index           =   7
            Left            =   2565
            TabIndex        =   18
            ToolTipText     =   "Mute"
            Top             =   840
            Width           =   750
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   510
            Index           =   7
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   900
            _Version        =   393216
            LargeChange     =   10
            Min             =   -1000
            Max             =   1000
            TickStyle       =   2
            TickFrequency   =   25
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   7
            Left            =   180
            OleObjectBlob   =   "frmMain.frx":0328
            TabIndex        =   87
            Top             =   1320
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   9
            Left            =   360
            OleObjectBlob   =   "frmMain.frx":0394
            TabIndex        =   88
            Top             =   1560
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   10
            Left            =   300
            OleObjectBlob   =   "frmMain.frx":03FE
            TabIndex        =   89
            Top             =   1800
            Width           =   435
         End
         Begin ACTIVESKINLibCtl.SkinLabel SongArtist 
            Height          =   195
            Left            =   780
            OleObjectBlob   =   "frmMain.frx":046A
            TabIndex        =   90
            Top             =   1800
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SongTitle 
            Height          =   195
            Left            =   780
            OleObjectBlob   =   "frmMain.frx":04C8
            TabIndex        =   91
            Top             =   1560
            Width           =   2535
         End
      End
      Begin MSComctlLib.Slider Slider1 
         CausesValidation=   0   'False
         Height          =   510
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   6
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   3
         Left            =   960
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   4
         Left            =   960
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   510
         Index           =   5
         Left            =   960
         TabIndex        =   13
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   2
         TickFrequency   =   2
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":0526
         TabIndex        =   81
         Top             =   900
         Width           =   795
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":0588
         TabIndex        =   82
         Top             =   1380
         Width           =   795
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":05EE
         TabIndex        =   83
         Top             =   1860
         Width           =   915
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   4
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":0662
         TabIndex        =   84
         Top             =   2340
         Width           =   795
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   5
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":06D4
         TabIndex        =   85
         Top             =   2820
         Width           =   795
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   6
         Left            =   3660
         OleObjectBlob   =   "frmMain.frx":0740
         TabIndex        =   86
         Top             =   420
         Width           =   795
      End
      Begin Project1.Sound2 Sound11 
         Index           =   6
         Left            =   6413
         Top             =   242
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         DeviceToControl =   6
      End
      Begin Project1.Sound2 Sound11 
         Index           =   5
         Left            =   3146
         Top             =   2662
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   81
         mute            =   -1  'True
         DeviceToControl =   1
      End
      Begin Project1.Sound2 Sound11 
         Index           =   4
         Left            =   3146
         Top             =   2178
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   50
         mute            =   -1  'True
         DeviceToControl =   2
      End
      Begin Project1.Sound2 Sound11 
         Index           =   3
         Left            =   3146
         Top             =   1694
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   50
         DeviceToControl =   3
      End
      Begin Project1.Sound2 Sound11 
         Index           =   2
         Left            =   3146
         Top             =   1210
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   100
      End
      Begin Project1.Sound2 Sound11 
         Index           =   1
         Left            =   3146
         Top             =   726
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   76
         DeviceToControl =   4
      End
      Begin Project1.Sound2 Sound11 
         Index           =   0
         Left            =   3146
         Top             =   242
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   0   'False
         Volume          =   82
         DeviceToControl =   0
      End
   End
   Begin VB.Frame frTab1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3225
      Left            =   120
      TabIndex        =   39
      Top             =   720
      Width           =   7095
      Begin VB.Frame Frame3 
         Caption         =   "MP3 Info"
         Height          =   3225
         Left            =   4200
         TabIndex        =   41
         Top             =   0
         Width           =   2885
         Begin ACTIVESKINLibCtl.SkinLabel mp3info 
            Height          =   2910
            Left            =   105
            OleObjectBlob   =   "frmMain.frx":07B0
            TabIndex        =   98
            Top             =   225
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ID3 Tags"
         Height          =   3225
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   4125
         Begin VB.CheckBox cVer 
            Caption         =   "Check5"
            Height          =   210
            Index           =   2
            Left            =   3375
            TabIndex        =   120
            Top             =   1830
            Width           =   210
         End
         Begin VB.CheckBox cVer 
            Caption         =   "Check4"
            Height          =   210
            Index           =   1
            Left            =   2670
            TabIndex        =   119
            Top             =   1830
            Width           =   195
         End
         Begin VB.CommandButton cmdVer2 
            Caption         =   "2"
            Height          =   360
            Left            =   3345
            TabIndex        =   118
            Top             =   1755
            Width           =   660
         End
         Begin VB.CommandButton cmdVer1 
            Caption         =   "1"
            Height          =   360
            Left            =   2640
            TabIndex        =   117
            Top             =   1755
            Width           =   675
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   210
            Left            =   1995
            OleObjectBlob   =   "frmMain.frx":080E
            TabIndex        =   116
            Top             =   1830
            Width           =   585
         End
         Begin VB.CommandButton cReplace 
            Caption         =   "Find/Replace"
            Height          =   345
            Left            =   120
            TabIndex        =   31
            Top             =   2680
            Width           =   1215
         End
         Begin VB.TextBox Album 
            Height          =   330
            Left            =   1140
            TabIndex        =   25
            Top             =   210
            Width           =   2865
         End
         Begin VB.TextBox Artist 
            Height          =   330
            Left            =   1140
            TabIndex        =   26
            Top             =   585
            Width           =   2865
         End
         Begin VB.TextBox Title 
            Height          =   330
            Left            =   1140
            TabIndex        =   27
            Top             =   990
            Width           =   2865
         End
         Begin VB.TextBox Comment 
            Height          =   330
            Left            =   1140
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   2190
            Width           =   2865
         End
         Begin VB.TextBox Year 
            Height          =   330
            Left            =   1140
            TabIndex        =   29
            Top             =   1800
            Width           =   810
         End
         Begin VB.ComboBox Genre 
            Height          =   264
            Left            =   1140
            TabIndex        =   28
            Text            =   "Genre"
            Top             =   1395
            Width           =   2865
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear"
            Height          =   345
            Left            =   1440
            TabIndex        =   32
            Top             =   2680
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Save"
            Height          =   345
            Left            =   2250
            TabIndex        =   33
            Top             =   2680
            Width           =   810
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next »»"
            Default         =   -1  'True
            Height          =   345
            Left            =   3150
            TabIndex        =   34
            Top             =   2680
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   210
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   20
            Top             =   585
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   21
            Top             =   990
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   3
            Left            =   90
            TabIndex        =   22
            Top             =   1395
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   4
            Left            =   90
            TabIndex        =   23
            Top             =   1800
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   5
            Left            =   90
            TabIndex        =   24
            Top             =   2190
            Visible         =   0   'False
            Width           =   210
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   11
            Left            =   405
            OleObjectBlob   =   "frmMain.frx":0878
            TabIndex        =   92
            Top             =   240
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   12
            Left            =   405
            OleObjectBlob   =   "frmMain.frx":08E0
            TabIndex        =   93
            Top             =   600
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   13
            Left            =   360
            OleObjectBlob   =   "frmMain.frx":094A
            TabIndex        =   94
            Top             =   1005
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   14
            Left            =   345
            OleObjectBlob   =   "frmMain.frx":09B2
            TabIndex        =   95
            Top             =   1410
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   15
            Left            =   360
            OleObjectBlob   =   "frmMain.frx":0A1A
            TabIndex        =   96
            Top             =   1815
            Width           =   795
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   16
            Left            =   360
            OleObjectBlob   =   "frmMain.frx":0A80
            TabIndex        =   97
            Top             =   2220
            Width           =   795
         End
      End
   End
   Begin VB.Frame frTab2 
      Caption         =   "Rename"
      Height          =   3225
      Left            =   120
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton Command8 
         Caption         =   "Preview"
         Height          =   285
         Left            =   5160
         TabIndex        =   76
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Rename"
         Height          =   285
         Left            =   6120
         TabIndex        =   75
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtReName 
         Height          =   285
         Left            =   195
         TabIndex        =   57
         Text            =   "(%Artist%) - %Album% - %Title%"
         Top             =   360
         Width           =   4860
      End
      Begin VB.CheckBox ValidateDir 
         Caption         =   "Validate Existing Directories"
         Height          =   210
         Left            =   3840
         TabIndex        =   55
         Top             =   170
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chPresent 
         Caption         =   "Rename only if Tags are Present"
         Height          =   195
         Left            =   1005
         TabIndex        =   56
         Top             =   170
         Width           =   3615
      End
      Begin VB.Frame Frame5 
         Caption         =   "Options"
         Height          =   2535
         Left            =   45
         TabIndex        =   58
         Top             =   645
         Width           =   6990
         Begin VB.Frame Frame6 
            Caption         =   "Replace"
            Height          =   2355
            Left            =   4590
            TabIndex        =   59
            Top             =   135
            Width           =   2355
            Begin VB.CheckBox Own2 
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   1080
               Width           =   255
            End
            Begin VB.CheckBox Own1 
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   660
               Width           =   255
            End
            Begin VB.TextBox Rep1 
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
               Index           =   0
               Left            =   375
               TabIndex        =   71
               Top             =   600
               Width           =   510
            End
            Begin VB.TextBox Rep1 
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
               Index           =   1
               Left            =   1440
               TabIndex        =   70
               Top             =   615
               Width           =   600
            End
            Begin VB.TextBox Rep2 
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
               Index           =   0
               Left            =   375
               TabIndex        =   69
               Top             =   1050
               Width           =   525
            End
            Begin VB.TextBox Rep2 
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
               Index           =   1
               Left            =   1440
               TabIndex        =   68
               Top             =   1050
               Width           =   600
            End
            Begin VB.OptionButton chBefore 
               Caption         =   "Before"
               Height          =   210
               Left            =   120
               TabIndex        =   67
               Top             =   225
               Width           =   765
            End
            Begin VB.OptionButton chAfter 
               Caption         =   "After Rename"
               Height          =   210
               Left            =   885
               TabIndex        =   66
               Top             =   225
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.TextBox Rep4 
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
               Index           =   1
               Left            =   1440
               TabIndex        =   65
               Top             =   1890
               Width           =   600
            End
            Begin VB.TextBox Rep4 
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
               Index           =   0
               Left            =   375
               TabIndex        =   64
               Top             =   1890
               Width           =   525
            End
            Begin VB.CheckBox Own4 
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   1950
               Width           =   255
            End
            Begin VB.TextBox Rep3 
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
               Index           =   1
               Left            =   1440
               TabIndex        =   62
               Top             =   1455
               Width           =   600
            End
            Begin VB.TextBox Rep3 
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
               Index           =   0
               Left            =   375
               TabIndex        =   61
               Top             =   1440
               Width           =   510
            End
            Begin VB.CheckBox Own3 
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   1500
               Width           =   255
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   195
               Index           =   24
               Left            =   930
               OleObjectBlob   =   "frmMain.frx":0AEC
               TabIndex        =   106
               Top             =   660
               Width           =   465
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   195
               Index           =   25
               Left            =   930
               OleObjectBlob   =   "frmMain.frx":0B52
               TabIndex        =   107
               Top             =   1125
               Width           =   465
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   195
               Index           =   26
               Left            =   930
               OleObjectBlob   =   "frmMain.frx":0BB8
               TabIndex        =   108
               Top             =   1515
               Width           =   465
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   195
               Index           =   27
               Left            =   930
               OleObjectBlob   =   "frmMain.frx":0C1E
               TabIndex        =   109
               Top             =   1935
               Width           =   465
            End
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   23
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0C84
            TabIndex        =   99
            Top             =   345
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   17
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0D4A
            TabIndex        =   100
            Top             =   570
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   18
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0E08
            TabIndex        =   101
            Top             =   1500
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   19
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0EC4
            TabIndex        =   102
            Top             =   1245
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   20
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0F74
            TabIndex        =   103
            Top             =   1005
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   195
            Index           =   21
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":1028
            TabIndex        =   104
            Top             =   780
            Width           =   4365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   630
            Index           =   22
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":10E6
            TabIndex        =   105
            Top             =   1725
            Width           =   4365
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Index           =   28
         Left            =   60
         OleObjectBlob   =   "frmMain.frx":128C
         TabIndex        =   110
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.Frame frTab4 
      Caption         =   "Help"
      Height          =   3225
      Left            =   120
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   7095
      Begin VB.ListBox List2 
         Height          =   1740
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":12F6
         Left            =   120
         List            =   "frmMain.frx":12F8
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   2805
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   4815
      End
      Begin VB.ListBox List1 
         Height          =   1020
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":12FA
         Left            =   120
         List            =   "frmMain.frx":1310
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frTab3 
      Caption         =   "About Master LoKis ID3 Machine"
      Height          =   3225
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   7095
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   705
         Index           =   29
         Left            =   3690
         OleObjectBlob   =   "frmMain.frx":134C
         TabIndex        =   112
         Top             =   2010
         Width           =   2925
      End
      Begin VB.CommandButton Command7 
         Caption         =   "http://volure.zapto.org"
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   2760
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   1485
         Index           =   8
         Left            =   660
         OleObjectBlob   =   "frmMain.frx":143E
         TabIndex        =   111
         Top             =   525
         Width           =   5775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   0
      TabIndex        =   36
      Top             =   360
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6588
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Retag"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Press F1 For Help"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuListView 
      Caption         =   "Options"
      Begin VB.Menu mnuPlayInternally 
         Caption         =   "Play Internally"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop Internal Player"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open MP3"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open Folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu Splitter 
         Caption         =   "-"
      End
      Begin VB.Menu MP3Copy 
         Caption         =   "Copy MP3 Info"
         Visible         =   0   'False
         Begin VB.Menu CopyInfo 
            Caption         =   "Copy MP3 File Info"
         End
         Begin VB.Menu CopyRenameTag 
            Caption         =   "Copy ###"
         End
         Begin VB.Menu CopyAll 
            Caption         =   "Copy All info"
         End
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename MP3"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete MP3"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu Sett 
      Caption         =   "Settings"
      Begin VB.Menu DebugOn 
         Caption         =   "Error Output"
      End
   End
   Begin VB.Menu basemnu 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHlp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmTagger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Skin As String
Const MF_BYPOSITION = &H400&
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Const cHeight As Long = 8655
Const cWidth As Long = 7545
Const HKEY_CLASSES_ROOT = &H80000000
Public CancelIt As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Function IsDir(Directory As String) As Boolean
    'Is this actually a directory or a File named the same as a Directory
    If GetFileAttributes(Directory) And FILE_ATTRIBUTE_DIRECTORY Then
        IsDir = True
    Else
        IsDir = False
    End If
End Function

Private Sub Check1_Click()
    'Save this Setting
    SaveSetting Me.Caption, "Settings", "SubDirectories", Check1.Value
End Sub

Private Sub Check3_Click()
    'Save this Setting
    SaveSetting Me.Caption, "Settings", "Multiple", Check3.Value
    'Edit the Form activities to accomidate
    If Check3.Value = vbChecked Then
        ListView1.MultiSelect = True    'Listview is able to have more then one file selected
        cmdNext.Enabled = False         'Next is Disabled
        Check2(0).Visible = True        'Show the Check boxes by the Tag Properties
        Check2(1).Visible = True
        Check2(2).Visible = True
        Check2(3).Visible = True
        Check2(4).Visible = True
        Check2(5).Visible = True
    Else
        ListView1.MultiSelect = False   'Listview cannot select more then one file
        cmdNext.Enabled = True          'Next is Enabled (Can be accessed by the [Enter] key}
        Check2(0).Visible = False       'Hide the Check boxes by the Tag Properties
        Check2(1).Visible = False
        Check2(2).Visible = False
        Check2(3).Visible = False
        Check2(4).Visible = False
        Check2(5).Visible = False
    End If
        If ListView1.ListItems.Count <> 0 Then 'Enable these Menu Items if its not Multiselect
            mnuPlayInternally.Enabled = Not ListView1.MultiSelect
            mnuDelete.Enabled = Not ListView1.MultiSelect
            mnuRename.Enabled = Not ListView1.MultiSelect
            mnuOpen.Enabled = Not ListView1.MultiSelect
            mnuOpenFolder.Enabled = Not ListView1.MultiSelect
        Else    'Disable all menus cause theres no Files
            mnuPlayInternally.Enabled = False
            mnuDelete.Enabled = False
            mnuRename.Enabled = False
            mnuOpen.Enabled = False
            mnuOpenFolder.Enabled = False
        End If
End Sub

Private Sub cmdNext_Click()
    'Next is clicked
    If ListView1.ListItems.Count = 0 Then Exit Sub 'Just in case its not disabled
    Command3_Click 'Save the File first with the Save Button
    Album.SetFocus 'Set the focus to the Album Textbox
    If ListView1.MultiSelect = False Then 'Just in case its not Disabled
        If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
            If ListView1.ListItems.Count <> 1 Then
                ListView1.ListItems(1).Selected = True
                ListView1_ItemClick ListView1.SelectedItem
            End If
        Else
            ListView1.ListItems(ListView1.SelectedItem.Index + 1).Selected = True
            ListView1_ItemClick ListView1.SelectedItem
            
        End If
        ListView1.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub cmdVer1_Click()
    v2Last = False
    With MP3Tag
        .ID3v2 = False
        .ReadTag
        Artist = .Artist
        Title = .Title
        Album = .Album
        Genre.ListIndex = .GenreV1
        Year = .Year
        Comment = .Comments
    End With
End Sub

Private Sub cmdVer2_Click()
    v2Last = True
    With MP3Tag
        .ID3v2 = True
        .ReadTag
        Artist = .Artist
        Title = .Title
        Album = .Album
        Genre.Text = .GenreV2
        Year = .Year
        Comment = .Comments
    End With
End Sub

Private Sub Command1_Click()
    Dim XX As Long
    PreemptFolder = FolderPath.Caption
    Folder.Show vbModal, Me
    If Folder.ChoseFolder <> "" Then
        FolderPath.Caption = Folder.ChoseFolder
        SaveSetting Me.Caption, "Settings", "LastPath", Folder.ChoseFolder
        Command4_Click
    End If
    Unload Folder
    'Stop
    'Open "G:\Documents and Settings\BlueAdept\Desktop\test music\Alan_Jackson_-_Chattahoochie.mp3" For Binary Access Read As #1
    '    PreemptFolder = String(5000, " ")
    '    Get #1, , PreemptFolder
    '    'PreemptFolder = Mid(PreemptFolder, 1, InStr(1, PreemptFolder, "TIT2" & String(2, Chr(0)), vbBinaryCompare) + 250)
    'Close #1
    'Open "C:\Tmp.txt" For Output As #1
    '    For XX = 1 To Len(PreemptFolder)
    '        Print #1, "Chr(" & Asc(Mid(PreemptFolder, XX, 1)) & ") = " & Mid(PreemptFolder, XX, 1)
    '    Next XX
    'Close #1
    'Stop
End Sub

Private Sub Command2_Click()
    Album = ""
    Artist = ""
    Comment = ""
    Genre.ListIndex = -1
    Title = ""
    Year = ""
End Sub

Private Sub Command3_Click()
Dim I As Integer
'
On Error GoTo DebugErr
If ListView1.ListItems.Count = 0 Then Exit Sub
With MP3Tag
    If ListView1.MultiSelect = True Then
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Selected Then
                StatBar.Panels(1).Text = "Saving " & ListView1.ListItems(I).Text
                If cVer(1).Value = vbChecked Then
                    .ID3v2 = True
                    .Filename = ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text
                    .ReadTag
                    If Check2(0).Value = vbChecked Then .Album = Album
                    If Check2(1).Value = vbChecked Then .Artist = Artist
                    If Check2(5).Value = vbChecked Then .Comments = Comment
                    If Check2(3).Value = vbChecked Then .GenreV2 = Genre.Text
                    If Check2(2).Value = vbChecked Then .Title = Title
                    If Check2(4).Value = vbChecked Then .Year = Year
                    .WriteTag
                End If
                If cVer(1).Value = vbChecked Then
                    .ID3v2 = False
                    .Filename = ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text
                    .ReadTag
                    If Check2(0).Value = vbChecked Then .Album = Album
                    If Check2(1).Value = vbChecked Then .Artist = Artist
                    If Check2(5).Value = vbChecked Then .Comments = Comment
                    If Check2(3).Value = vbChecked Then .GenreV1 = Genre.ListIndex
                    If Check2(2).Value = vbChecked Then .Title = Title
                    If Check2(4).Value = vbChecked Then .Year = Year
                    .WriteTag
                End If
            End If
        Next I
        StatBar.Panels(1).Text = "Saved All"
    Else
        If cVer(2).Value = vbChecked Then
            .ID3v2 = True
            .Filename = ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
            .ReadTag
            .Album = Album
            .Artist = Artist
            .Comments = Comment
            .GenreV2 = Genre.Text
            .Title = Title
            .Year = Year
            .WriteTag
        End If
        If cVer(1).Value = vbChecked Then
            .ID3v2 = False
            .Filename = ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
            .ReadTag
            .Album = Album
            .Artist = Artist
            .Comments = Comment
            .GenreV1 = Genre.ListIndex
            .Title = Title
            .Year = Year
            .WriteTag
        End If
    End If
End With
If writeError Then Err_Show Me
Exit Sub
DebugErr:
Err.Raise Err.Number
End Sub

Private Sub Command4_Click()
    Me.Visible = False
    Parse.Visible = True
    ParseFiles
    Unload Parse
    Me.Visible = True
End Sub

Private Sub Command5_Click()
If ListView1.ListItems.Count = 0 Then Exit Sub
Dim I As Integer
Dim Tmp As String
On Error Resume Next
Err.Clear
If ListView1.MultiSelect Then
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Selected = True Then
            If TagsArePresent(ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text, txtReName) Then
                Tmp = TagRename(ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text)
                If ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text = ListView1.ListItems(I).SubItems(1) & Tmp Then
                    'Do Nothing
                ElseIf Tmp <> "" Then
                    Name ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text As ListView1.ListItems(I).SubItems(1) & Tmp
                    If Err.Number <> 0 Then
                        ERR_Add "Err:" & Err.Number & " " & Err.Description
                        ERR_Add "  New Name " & Tmp
                        ERR_Add "  Old Name " & ListView1.ListItems(I).Text
                        Err.Clear
                    Else
                        ListView1.ListItems(I).Text = Tmp
                    End If
                End If
            End If
        End If
    Next I
Else
    If TagsArePresent(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text, txtReName) Then
        Tmp = TagRename(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text)
        If Tmp <> "" Then
            Name ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text As ListView1.SelectedItem.SubItems(1) & TagRename(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text)
            If Err.Number <> 0 Then
                ERR_Add "Err:" & Err.Number & " " & Err.Description
                ERR_Add "  New Name " & Tmp
                ERR_Add "  Old Name " & ListView1.SelectedItem.Text
                Err.Clear
            Else
                ListView1.SelectedItem.Text = Tmp
            End If
        End If
    End If
End If
If DebugOn.Checked Then Err_Show Me
Err_Clear
End Sub

Private Function TagRename(Filename As String) As String
    Dim Tmp As String
    Dim TmpMP3 As cIDV3
    Set TmpMP3 = New cIDV3
    With TmpMP3
        .Filename = Filename
        .ReadTag
        Tmp = txtReName
        If chBefore.Value = True Then
            If Own1.Value = vbChecked Then Tmp = Replace(Tmp, Rep1(0), Rep1(1))
            If Own2.Value = vbChecked Then Tmp = Replace(Tmp, Rep2(0), Rep2(1))
            If Own3.Value = vbChecked Then Tmp = Replace(Tmp, Rep3(0), Rep3(1))
            If Own4.Value = vbChecked Then Tmp = Replace(Tmp, Rep4(0), Rep4(1))
        End If
            Tmp = Replace(Tmp, "%Artist%", .Artist, , , vbTextCompare)
            Tmp = Replace(Tmp, "%Album%", .Album, , , vbTextCompare)
            Tmp = Replace(Tmp, "%Title%", .Title, , , vbTextCompare)
            Tmp = Replace(Tmp, "%Year%", .Year, , , vbTextCompare)
            Tmp = Replace(Tmp, "%Genre%", .GenreText(.GenreV1), , , vbTextCompare)
            Tmp = Replace(Tmp, "%Comment%", .Comments, , , vbTextCompare)
            '/\:*"<>?|
            Tmp = Replace(Tmp, "\", "")
            Tmp = Replace(Tmp, "/", "")
            Tmp = Replace(Tmp, ":", "")
            Tmp = Replace(Tmp, "*", "")
            Tmp = Replace(Tmp, """", "")
            Tmp = Replace(Tmp, "<", "")
            Tmp = Replace(Tmp, ">", "")
            Tmp = Replace(Tmp, "?", "")
            Tmp = Replace(Tmp, "|", "")
        If chAfter.Value = True Then
            If Own1.Value = vbChecked Then Tmp = Replace(Tmp, Rep1(0), Rep1(1))
            If Own2.Value = vbChecked Then Tmp = Replace(Tmp, Rep2(0), Rep2(1))
            If Own3.Value = vbChecked Then Tmp = Replace(Tmp, Rep3(0), Rep3(1))
            If Own4.Value = vbChecked Then Tmp = Replace(Tmp, Rep4(0), Rep4(1))
        End If
    End With
    Tmp = Replace(Tmp, Chr(0), "")
    If LCase(Right(Tmp, 4)) <> ".mp3" Then Tmp = Tmp & ".mp3"
    TagRename = Tmp
End Function

Private Function TagsArePresent(Filename As String, TagString As String) As Boolean
    TagsArePresent = True
    Dim TmpMP3 As cIDV3
    Set TmpMP3 = New cIDV3
    With TmpMP3
        .Filename = Filename
        .ReadTag
        If chPresent.Value = vbUnchecked Then Exit Function
        If InStr(1, TagString, "%Album%", vbTextCompare) > 0 And Replace(.Album, Chr(0), "") = "" Then TagsArePresent = False
        If InStr(1, TagString, "%Artist%", vbTextCompare) > 0 And Replace(.Artist, Chr(0), "") = "" Then TagsArePresent = False
        If InStr(1, TagString, "%Title%", vbTextCompare) > 0 And Replace(.Title, Chr(0), "") = "" Then TagsArePresent = False
        If InStr(1, TagString, "%Genre%", vbTextCompare) > 0 And Replace(.GenreText(.GenreV1), Chr(0), "") = "" Then TagsArePresent = False
        If InStr(1, TagString, "%Year%", vbTextCompare) > 0 And Replace(.Year, Chr(0), "") = "" Then TagsArePresent = False
        If InStr(1, TagString, "%Comments%", vbTextCompare) > 0 And Replace(.Comments, Chr(0), "") = "" Then TagsArePresent = False
    End With
End Function

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim I As Integer
    Dim TmpMP3 As cIDV3
    Set TmpMP3 = New cIDV3
    Command5.ToolTipText = "Hover for Rename Preview"
    If ListView1.ListItems.Count = 0 Then Exit Sub
End Sub

Private Sub Command6_Click()
    Dim I As Integer
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.MultiSelect = False Then Exit Sub
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Selected = True
    Next I
End Sub

Private Sub Command7_Click()
    ShellExecute Me.hwnd, vbNullString, Command7.Caption, vbNullString, "C:\", 1
End Sub

Private Sub Command8_Click()
    Dim I As Integer
    Dim I2 As Integer
    Dim Tmp As String
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.MultiSelect Then
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Selected = True Then
                If TagsArePresent(ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text, txtReName) Then
                    Tmp = "These are the First MP3's" & vbCrLf & vbCrLf
                    I2 = I2 + 1
                    Tmp = Tmp & TagRename(ListView1.ListItems(I).SubItems(1) & ListView1.ListItems(I).Text) & vbCrLf
                End If
                If I2 = 3 Then Exit For
            End If
        Next I
        If Tmp = "" Then Tmp = "No MP3's Were Capable of Being Renamed" & Chr(0)
    Else
        I2 = I2 + 1
        If TagsArePresent(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text, txtReName) Then
            Tmp = Tmp & TagRename(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text) & vbCrLf
        End If
    End If
    MsgBox Tmp & vbCrLf & "Press Rename to Apply " & IIf(I2 > 1, "these names", "this name")
End Sub

Private Sub CopyAll_Click()
'Dim TmpStr As String
'With MP3Tag
'    If chBefore.Value = True Then
'        If Own1.Value = vbChecked Then TmpName = Replace(TmpName, Rep1(0).Text, Rep1(1).Text, , , vbTextCompare)
'        If Own2.Value = vbChecked Then TmpName = Replace(TmpName, Rep2(0).Text, Rep2(1).Text, , , vbTextCompare)
'        If Own3.Value = vbChecked Then TmpName = Replace(TmpName, Rep3(0).Text, Rep3(1).Text, , , vbTextCompare)
'        If Own4.Value = vbChecked Then TmpName = Replace(TmpName, Rep4(0).Text, Rep4(1).Text, , , vbTextCompare)
'    End If
'    TmpName = Replace(txtReName, "%Artist%", .Artist, , , vbBinaryCompare)
'    TmpName = Replace(TmpName, "%Album%", .Album, , , vbBinaryCompare)
'    TmpName = Replace(TmpName, "%Title%", .Title, , , vbBinaryCompare)
'    TmpName = Replace(TmpName, "%Genre%", .GenreText(.Genre), , , vbBinaryCompare)
'    TmpName = Replace(TmpName, "%Year%", .Year, , , vbBinaryCompare)
'    TmpName = Replace(TmpName, "%Comment%", .Comments, , , vbBinaryCompare)
'    If chAfter.Value = True Then
'        If Own1.Value = vbChecked Then TmpName = Replace(TmpName, Rep1(0).Text, Rep1(1).Text)
'        If Own2.Value = vbChecked Then TmpName = Replace(TmpName, Rep2(0).Text, Rep2(1).Text)
'        If Own3.Value = vbChecked Then TmpName = Replace(TmpName, Rep3(0).Text, Rep3(1).Text)
'        If Own4.Value = vbChecked Then TmpName = Replace(TmpName, Rep4(0).Text, Rep4(1).Text)
'    End If
'    Do
'        TmpName = Replace(TmpName, ">", "")
'        TmpName = Replace(TmpName, "<", "")
'        TmpName = Replace(TmpName, "/", "")
'        TmpName = Replace(TmpName, "?", "")
'        TmpName = Replace(TmpName, "*", "")
'        TmpName = Replace(TmpName, ":", "")
'        TmpName = Replace(TmpName, "\\", "")
'        TmpName = Replace(TmpName, "|", "")
'        TmpName = Replace(TmpName, ">", "")
'        TmpName = Replace(TmpName, "<", "")
'        TmpName = Replace(TmpName, "/", "")
'        TmpName = Replace(TmpName, "?", "")
'        TmpName = Replace(TmpName, "*", "")
'        TmpName = Replace(TmpName, ":", "")
'        TmpName = Replace(TmpName, "\\", "")
'        TmpName = Replace(TmpName, "|", "")
'
End Sub

Private Sub DebugOn_Click()
    DebugOn.Checked = Not DebugOn.Checked 'Allow Debug
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    '
    List4.AddItem CmdStr 'Remote Commands
    
End Sub

Private Sub Form_Load()
Dim hMenu As Long
Dim hSubMenu As Long
Skin1.Empty 'Empty the Skin
If FileExists(App.Path & "\Skins\frmTagger.Skn") Then 'Check for Local skin
    Skin1.LoadSkin App.Path & "\Skins\frmTagger.Skn" 'Apply Skin if Exists
End If
Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor 'Add Version to the Caption
Dim I As Integer
Dim FF As Integer
Dim Tmp As String
FF = FreeFile 'Set a Free File Variable Preventing Old file access Errors
For I = 0 To 6 'Activate Sound Objects and Set Volumes
    Sound11(I).Enabled = True
    Slider1(I).Value = Sound11(I).Volume
    If Sound11(I).Mute = True Then
        Mute(I).Value = vbChecked
    Else
        Mute(I).Value = vbUnchecked
    End If
    Slider1(I).Enabled = Sound11(I).Enabled
    Mute(I).Enabled = Sound11(I).Enabled
Next I
    Set MP3Tag = New cIDV3 'Set the MP3Tag Object
    Genre.Clear
    'Add all the genre's
    Genre.AddItem "[Blues]"
    Genre.AddItem "[Classic Rock]"
    Genre.AddItem "[Country]"
    Genre.AddItem "[Dance]"
    Genre.AddItem "[Disco]"
    Genre.AddItem "[Funk]"
    Genre.AddItem "[Grunge]"
    Genre.AddItem "[Hip-Hop]"
    Genre.AddItem "[Jazz]"
    Genre.AddItem "[Metal]"
    Genre.AddItem "[New Age]"
    Genre.AddItem "[Oldies]"
    Genre.AddItem "[Other]"
    Genre.AddItem "[Pop]"
    Genre.AddItem "[R&B]"
    Genre.AddItem "[Rap]"
    Genre.AddItem "[Reggae]"
    Genre.AddItem "[Rock]"
    Genre.AddItem "[Techno]"
    Genre.AddItem "[Industrial]"
    Genre.AddItem "[Alternative]"
    Genre.AddItem "[Ska]"
    Genre.AddItem "[Death Metal]"
    Genre.AddItem "[Pranks]"
    Genre.AddItem "[Soundtrack]"
    Genre.AddItem "[Euro-Techno]"
    Genre.AddItem "[Ambient]"
    Genre.AddItem "[Trip-Hop]"
    Genre.AddItem "[Vocal]"
    Genre.AddItem "[Jazz+Funk]"
    Genre.AddItem "[Fusion]"
    Genre.AddItem "[Trance]"
    Genre.AddItem "[Classical]"
    Genre.AddItem "[Instrumental]"
    Genre.AddItem "[Acid]"
    Genre.AddItem "[House]"
    Genre.AddItem "[Game]"
    Genre.AddItem "[Sound Clip]"
    Genre.AddItem "[Gospel]"
    Genre.AddItem "[Noise]"
    Genre.AddItem "[Alt. Rock]"
    Genre.AddItem "[Bass]"
    Genre.AddItem "[Soul]"
    Genre.AddItem "[Punk]"
    Genre.AddItem "[Space]"
    Genre.AddItem "[Meditative]"
    Genre.AddItem "[Instrumental Pop]"
    Genre.AddItem "[Instrumental Rock]"
    Genre.AddItem "[Ethnic]"
    Genre.AddItem "[Gothic]"
    Genre.AddItem "[Darkwave]"
    Genre.AddItem "[Techno-Industrial]"
    Genre.AddItem "[Electronic]"
    Genre.AddItem "[Pop-Folk]"
    Genre.AddItem "[Eurodance]"
    Genre.AddItem "[Dream]"
    Genre.AddItem "[Southern Rock]"
    Genre.AddItem "[Comedy]"
    Genre.AddItem "[Cult]"
    Genre.AddItem "[Gangsta Rap]"
    Genre.AddItem "[Top 40]"
    Genre.AddItem "[Christian Rap]"
    Genre.AddItem "[Pop/Funk]"
    Genre.AddItem "[Jungle]"
    Genre.AddItem "[Native American]"
    Genre.AddItem "[Cabaret]"
    Genre.AddItem "[New Wave]"
    Genre.AddItem "[Phychedelic]"
    Genre.AddItem "[Rave]"
    Genre.AddItem "[Showtunes]"
    Genre.AddItem "[Trailer]"
    Genre.AddItem "[Lo-Fi]"
    Genre.AddItem "[Tribal]"
    Genre.AddItem "[Acid Punk]"
    Genre.AddItem "[Acid Jazz]"
    Genre.AddItem "[Polka]"
    Genre.AddItem "[Retro]"
    Genre.AddItem "[Musical]"
    Genre.AddItem "[Rock & Roll]"
    Genre.AddItem "[Hard Rock]"
    Genre.AddItem "[Folk]"
    Genre.AddItem "[Folk/Rock]"
    Genre.AddItem "[National Folk]"
    Genre.AddItem "[Swing]"
    Genre.AddItem "[Fast-Fusion]"
    Genre.AddItem "[Bebob]"
    Genre.AddItem "[Latin]"
    Genre.AddItem "[Revival]"
    Genre.AddItem "[Celtic]"
    Genre.AddItem "[Bluegrass]"
    Genre.AddItem "[Avantegarde]"
    Genre.AddItem "[Gothic Rock]"
    Genre.AddItem "[Progressive Rock]"
    Genre.AddItem "[Psychedelic Rock]"
    Genre.AddItem "[Symphonic Rock]"
    Genre.AddItem "[Slow Rock]"
    Genre.AddItem "[Big Band]"
    Genre.AddItem "[Chorus]"
    Genre.AddItem "[Easy Listening]"
    Genre.AddItem "[Acoustic]"
    Genre.AddItem "[Humour]"
    Genre.AddItem "[Speech]"
    Genre.AddItem "[Chanson]"
    Genre.AddItem "[Opera]"
    Genre.AddItem "[Chamber Music]"
    Genre.AddItem "[Sonata]"
    Genre.AddItem "[Symphony]"
    Genre.AddItem "[Booty Bass]"
    Genre.AddItem "[Primus]"
    Genre.AddItem "[Porn Groove]"
    Genre.AddItem "[Satire]"
    Genre.AddItem "[Slow Jam]"
    Genre.AddItem "[Club]"
    Genre.AddItem "[Tango]"
    Genre.AddItem "[Samba]"
    Genre.AddItem "[Folklore]"
    Genre.AddItem "[Ballad]"
    Genre.AddItem "[Power Ballad]"
    Genre.AddItem "[Rhythmic Soul]"
    Genre.AddItem "[Freestyle]"
    Genre.AddItem "[Duet]"
    Genre.AddItem "[Punk Rock]"
    Genre.AddItem "[Drum Solo]"
    Genre.AddItem "[A Capella]"
    Genre.AddItem "[Euro-House]"
    Genre.AddItem "[Dance Hall]"
    Genre.AddItem "[Goa]"
    Genre.AddItem "[Drum & Bass]"
    Genre.AddItem "[Club-House]"
    Genre.AddItem "[Hardcore]"
    Genre.AddItem "[Terror]"
    Genre.AddItem "[Indie]"
    Genre.AddItem "[BritPop]"
    Genre.AddItem "[Negerpunk]"
    Genre.AddItem "[Polsk Punk]"
    Genre.AddItem "[Beat]"
    Genre.AddItem "[Christian Gangsta Rap]"
    Genre.AddItem "[Heavy Metal]"
    Genre.AddItem "[Black Metal]"
    Genre.AddItem "[Crossover]"
    Genre.AddItem "[Contemporary Christian]"
    Genre.AddItem "[Christian Rock]"
    Genre.AddItem "[Merengue]"
    Genre.AddItem "[Salsa]"
    Genre.AddItem "[Trash Metal]"
    Genre.AddItem "[Anime]"
    Genre.AddItem "[JPop]"
    Genre.AddItem "[Synthpop]"
    'Load settings from Registry
    txtReName = GetSetting(Me.Caption, "Settings", "RenameFormat", txtReName)
    If GetSetting(Me.Caption, "Settings", "RepBA", 1) = 0 Then chBefore.Value = True
    If GetSetting(Me.Caption, "Settings\Own", "S1a", False) = True Then Own1.Value = vbChecked
    Rep1(0).Text = GetSetting(Me.Caption, "Settings\Own", "S1b", "")
    Rep1(1).Text = GetSetting(Me.Caption, "Settings\Own", "S1c", "")
    If GetSetting(Me.Caption, "Settings\Own", "S2a", False) = True Then Own2.Value = vbChecked
    Rep2(0).Text = GetSetting(Me.Caption, "Settings\Own", "S2b", "")
    Rep2(1).Text = GetSetting(Me.Caption, "Settings\Own", "S2c", "")
    If GetSetting(Me.Caption, "Settings\Own", "S3a", False) = True Then Own3.Value = vbChecked
    Rep3(0).Text = GetSetting(Me.Caption, "Settings\Own", "S3b", "")
    Rep3(1).Text = GetSetting(Me.Caption, "Settings\Own", "S3c", "")
    If GetSetting(Me.Caption, "Settings\Own", "S4a", False) = True Then Own4.Value = vbChecked
    Rep4(0).Text = GetSetting(Me.Caption, "Settings\Own", "S4b", "")
    Rep4(1).Text = GetSetting(Me.Caption, "Settings\Own", "S4c", "")
    FolderPath.Caption = GetSetting(Me.Caption, "Settings", "Lastpath", "C:\")
    v2Last = GetSetting(Me.Caption, "Settings", "v2Last", False)
    MP3Tag.ID3v2 = v2Last
    
    If GetSetting(Me.Caption, "Settings", "DebugOn", False) = True Then DebugOn.Checked = True
    
    If GetSetting(Me.Caption, "Settings", "V1", False) = True Then cVer(1).Value = vbChecked
    If GetSetting(Me.Caption, "Settings", "V2", False) = True Then cVer(2).Value = vbChecked
    If Val(GetSetting(Me.Caption, "Settings", "SubDirectories", vbUnchecked)) = vbChecked Then Check1.Value = vbChecked
    If Val(GetSetting(Me.Caption, "Settings", "CheckTags", vbUnchecked)) = vbChecked Then chPresent.Value = vbChecked
    If Val(GetSetting(Me.Caption, "Settings", "ValidateDir", vbUnchecked)) = vbChecked Then ValidateDir.Value = vbChecked
    If Val(GetSetting(Me.Caption, "Settings", "Multiple", vbUnchecked)) = vbChecked Then
        Check3.Value = vbChecked
        ListView1.MultiSelect = True
    Else
        ListView1.MultiSelect = False
    End If
    'Add Parsing Comments
    With List3
        .AddItem "Lets Get Busy and Load this Crap Up"
        .AddItem "I cant tell How long this is gonna take. You got too many Danged MP3's"
        .AddItem "I think You should Get a Faster CPU or Lose some Music. But thats just My opinion" & vbCrLf & _
                 "I could be wrong"
        .AddItem "Somehow I cant see the Light." & vbCrLf & _
                 "*Looks Down the Tunnel*"
        .AddItem "Well If you think hard enough You can Smell Smoke"
    End With
    Me.Visible = False 'Hide me
    Parse.Visible = True 'Show Parsing
    ParseFiles 'Activate Parsing
    Unload Parse 'Hide Parsing
    Slider1(7).Value = MediaPlayer1.Volume + 1000 'Set mediaplayer Volume
    'Reset parsing Comments from the Comments File
    With List3
        .Clear
        .AddItem "You can Add more Random Comments to this Window. Check out " & vbCrLf & App.Path & App.EXEName & ".cmt" & vbCrLf & " and add one for each line"
        If LCase(Dir(App.Path & "\" & App.EXEName & ".cmt")) = LCase(App.EXEName & ".cmt") Then
            FF = FreeFile
            Open App.Path & "\" & App.EXEName & ".cmt" For Input As #FF
                Do Until EOF(FF)
                    Line Input #FF, Tmp
                    .AddItem Replace(Tmp, "%crlf%", vbCrLf, , , vbTextCompare)
                Loop
            Close #FF
        End If
    End With
    'Apply the Skin
    Skin1.ApplySkin Me.hwnd
    
    Skin1.ApplySkinByName PMode(1).hwnd, "Button"
    Skin1.ApplySkinByName PMode(2).hwnd, "Button"
    frmTagger.Visible = True
End Sub


Private Sub Form_Resize()
    'Auto Size when Form is loaded (Accomidate for the Skin)
    With frmedia
        .Top = StatBar.Top + 20
        .Left = StatBar.Left + 20
        .Width = StatBar.Panels(1).Width
        MediaPlayer1.Width = .Width
    End With
    ListView1.Left = TabStrip1.Left
    ListView1.Width = Me.Width - (ListView1.Left * 2)
    ListView1.Height = StatBar.Top - ListView1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Frame1.Enabled = False Then Cancel = True: Exit Sub
    'Save Settings
    SaveSetting Me.Caption, "Settings", "DebugOn", DebugOn.Checked
    SaveSetting Me.Caption, "Settings", "RenameFormat", txtReName
    If chBefore.Value = True Then
        SaveSetting Me.Caption, "Settings", "RepBA", 0
    Else
        SaveSetting Me.Caption, "Settings", "RepBA", 1
    End If
    SaveSetting Me.Caption, "Settings", "CheckTags", chPresent.Value
    SaveSetting Me.Caption, "Settings", "ValidateDir", ValidateDir.Value
    SaveSetting Me.Caption, "Settings\Own", "S1a", Own1.Value
    SaveSetting Me.Caption, "Settings\Own", "S1b", Rep1(0).Text
    SaveSetting Me.Caption, "Settings\Own", "S1c", Rep1(1).Text
    SaveSetting Me.Caption, "Settings\Own", "S2a", Own2.Value
    SaveSetting Me.Caption, "Settings\Own", "S2b", Rep2(0).Text
    SaveSetting Me.Caption, "Settings\Own", "S2c", Rep2(1).Text
    SaveSetting Me.Caption, "Settings\Own", "S3a", Own3.Value
    SaveSetting Me.Caption, "Settings\Own", "S3b", Rep3(0).Text
    SaveSetting Me.Caption, "Settings\Own", "S3c", Rep3(1).Text
    SaveSetting Me.Caption, "Settings\Own", "S4a", Own4.Value
    SaveSetting Me.Caption, "Settings\Own", "S4b", Rep4(0).Text
    SaveSetting Me.Caption, "Settings\Own", "S4c", Rep4(1).Text
    SaveSetting Me.Caption, "Settings", "v2Last", v2Last
    SaveSetting Me.Caption, "Settings", "V1", cVer(1).Value
    SaveSetting Me.Caption, "Settings", "V2", cVer(2).Value
    Unload Me 'Unload
    
End Sub

Private Sub List1_Click()
Dim Tmp As String
With List2
    .Clear
    'Change the Sub Help Items when the Main item is clicked
    'Also load help info when needed
    Select Case List1.List(List1.ListIndex)
        Case "Extra"
            Tmp = "Extra Tweaks" & vbCrLf & _
                  "" & vbCrLf & _
                  "         These things arent Really Nessecary to learn. Im working on a couple things. they should appear in order in the bottom list to the Left" & vbCrLf & _
                  ""
            .AddItem "Parsing"
            
        Case "Retag"
            Tmp = "Retag" & vbCrLf & _
                  "" & vbCrLf & _
                  "         The Retag options are listed to the left. is theres Something you Still dont understand Please Contact me at MrLoKi@lvcm.com" & vbCrLf & _
                  ""
            .AddItem "Multiple Checkbox"
            .AddItem "SubFolders Checkbox"
            .AddItem "Parse Files"
            .AddItem "Select All"
            .AddItem "Find/Replace"
            .AddItem "Clear"
            .AddItem "Save"
            .AddItem "Next »»"
            .AddItem "MP3 Info"
        Case "Rename"
            Tmp = "Rename" & vbCrLf & _
                  "" & vbCrLf & _
                  "         The Rename options are listed to the left. is theres Something you Still dont understand Please Contact me at MrLoKi@lvcm.com" & vbCrLf & _
                  ""
            .AddItem "Format"
            .AddItem "%Tags%"
            .AddItem """Rename Only if"" Check"
            .AddItem "Rename"
            .AddItem "Preview"
            .AddItem "Replace"
        Case "Player"
            .AddItem "Volume Controls"
            .AddItem "Checkboxes"
            .AddItem "Internal Player"
        Case "Shortcuts"
            Tmp = "Shortcuts" & vbCrLf & _
                  "" & vbCrLf & _
                  "         The Shortcuts are listed to the left. is theres Something you Still dont understand Please Contact me at MrLoKi@lvcm.com" & vbCrLf & _
                  ""
            .AddItem "Enter"
            .AddItem "Ctrl+P"
            .AddItem "Ctrl+O"
            .AddItem "Ctrl+F"
            .AddItem "F2"
            .AddItem "Del"
            .AddItem "Drag & Drop"
        Case "Intended Uses"
            Tmp = "Intended Uses" & vbCrLf & _
                  "" & vbCrLf & _
                  "         I have not listed any Intended uses. If you think of a good Idea for using this Tagger Contact me at MrLoKi@lvcm.com and Ill try to list them here" & vbCrLf & _
                  ""
            
    End Select
        Text1.Text = Tmp
End With
End Sub

Private Sub List2_Click()
Dim Tmp As String
Select Case List2.List(List2.ListIndex)
    'Change the Sub Help Items when the Main item is clicked
    'Also load help info when needed
    Case "Parsing"
        Tmp = "Parsing" & vbCrLf & _
              "" & vbCrLf & _
              "         In the App path """ & App.Path & """ You can open or create a file named """ & App.EXEName & ".cmt"" with Notepad. Each Line is a new comment" & vbCrLf & _
              ""
    Case "Volume Controls"
        Tmp = "Volume Controls" & vbCrLf & _
              "" & vbCrLf & _
              "         The Volume controls should be labeled as they are in the Windows Volume Controls. They should also work accordingly." & vbCrLf & _
              "" & vbCrLf & _
              "   Master: Volume and Mute Controls the General overall volume and Mute" & vbCrLf & _
              "   CD: Volume and Mute Controls the CD Rom Audio" & vbCrLf & _
              "   Wave: Volume and Mute Controls the Wave and MP3 Players" & vbCrLf & _
              "   Synthesizer: Volume and Mute Controls the MIDI Players" & vbCrLf & _
              "   Microphone: Volume and Mute Controls the Microphone" & vbCrLf & _
              "   Line in: Volume and Mute Controls the Line in" & vbCrLf & _
              "   Auxiliary: Volume and Mute Controls the Auxiliary" & vbCrLf & _
              ""
    Case "Checkboxes"
        Tmp = "Checkboxes" & vbCrLf & _
              "" & vbCrLf & _
              "         The checkboxes next to each of the Variable Sliders Shuts off the Volume to that Specific Volume control" & vbCrLf & _
              ""
    Case "Internal Player"
        Tmp = "Internal Player" & vbCrLf & _
              "" & vbCrLf & _
              "         Here is the basic controls for the Internal Player. You can select the file you want to play with the Right click on the File list" & vbCrLf & _
              "" & vbCrLf & _
              "         The buttons only work if theres something they can do. Otherwise stop will be selected." & vbCrLf & _
              "" & vbCrLf & _
              "         The Mute button Mutes the Internal Player Only" & vbCrLf & _
              "" & vbCrLf & _
              "         The Variable Slider is for Setting the Volume of the Internal Player it Ranges from -1000 to 1000 and can be set by Scrolling the Slider" & vbCrLf & _
              "" & vbCrLf & _
              "     HINT: It should show you the Volume while you Scroll" & vbCrLf & _
              ""
    Case "Multiple Checkbox"
        Tmp = "Multiple Checkbox" & vbCrLf & _
              "" & vbCrLf & _
              "         This program is intended for Single and Multiple MP3's." & vbCrLf & _
              "" & vbCrLf & _
              "         In Single mode you can edit the Artist and Title Comments or Whatever then Press Enter this will save the File and open the next one. (Enter is not Available if Multiple is Checked)" & vbCrLf & _
              "" & vbCrLf & _
              "         In Multiple Mode you can Check the Boxes next to the Tags you want to apply to Your Selected MP3's and when Saving it will only Affect the Checked Tags" & vbCrLf & _
              "         (In Single mode these Checkboxes are not Present)" & vbCrLf & _
              "" & vbCrLf & _
              "the only Files that Will be affected in Multiple Mode are the ones selected in the List below (See Select All)"
              
    Case "SubFolders Checkbox"
        Tmp = "Subfolders Checkbox" & vbCrLf & _
              "" & vbCrLf & _
              "         Subfolders are Directories inside Directories. Checking this Option before Parsing will Check all the Sibdirectoried in the Root Directory Shown Above." & vbCrLf & _
              "" & vbCrLf & _
              "     HINT: Selecting C:\ as Your root Directory and Parsing with SubFolders Checked. It will search the entire Hard Drive" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: This can slow it down Tremendously if you have a lot of SubFolders"
    Case "Parse Files"
        Tmp = "Parse Files" & vbCrLf & _
              "" & vbCrLf & _
              "         Parse Files will Clear the MP3 Files List and Re-Search the Folder for MP3's" & vbCrLf & _
              "" & vbCrLf & _
              "         This Task is Executed when the Application is Started"
    Case "Select All"
        Tmp = "Select All" & vbCrLf & _
              "" & vbCrLf & _
              "         This will Select all the MP3's in the List below. This is useful for Applying Specific Tags all the MP3's or Find/Replacing all the MP3's just as _'s in the Artist Tag or something like that" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: If you have a lot of Files and all of them selected the taging and Find/Replacing will take a little Longer" & vbCrLf & _
              ""
    Case "Find/Replace"
        Tmp = "Find/Replace" & vbCrLf & _
              "" & vbCrLf & _
              "         A Dialog will pop up allowing you to search through the selected MP3's. When you find the Desired MP3 you can drag the Music Note to Drop it where you want it." & vbCrLf & _
              "" & vbCrLf & _
              "(See Drag & Drop)" & vbCrLf & _
              "" & vbCrLf & _
              "     HINT:Check the Checkboxes to Search only those fields and Uncheck them to Search nothing" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: There is no UNDO" & vbCrLf & _
              ""
    Case "Clear"
        Tmp = "Clear" & vbCrLf & _
              "" & vbCrLf & _
              "         This will clear all the tabs for you to Change them or save them erased" & vbCrLf & _
              ""
    Case "Save"
        Tmp = "Save" & vbCrLf & _
              "" & vbCrLf & _
              "         This will save the Selected Tags" & vbCrLf & _
              "" & vbCrLf & _
              "(On single it will save all the tags and on Multiple it will Save only the ones with the Check boxes Checked)" & vbCrLf & _
              ""
    Case "Next »»"
        Tmp = "Next »»" & vbCrLf & _
              "" & vbCrLf & _
              "         The next button is Triggered when Enter is Pressed." & vbCrLf & _
              "" & vbCrLf & _
              "         It works like Pressing Save and Selecting the next file in the list" & vbCrLf & _
              ""
    Case "MP3 Info"
        Tmp = "MP3 Info" & vbCrLf & _
              "" & vbCrLf & _
              "         This has all the info Pertaining to the MP3. Im working on a Copy Procedure" & vbCrLf & _
              ""
    Case "Format"
        Tmp = "Format" & vbCrLf & _
              "" & vbCrLf & _
              "         This will define the Format of the Renaming" & vbCrLf & _
              "" & vbCrLf & _
              "(See %Tags%)" & vbCrLf & _
              ""
    Case "%Tags%"
        Tmp = "Tags" & vbCrLf & _
              "" & vbCrLf & _
              "         These tags will be replaced with the Information from the ID3 Tags" & vbCrLf & _
              "" & vbCrLf & _
              "     %Album% will be replaced with the Album tag" & vbCrLf & _
              "     %Artist% will be replaced with the Artist tag" & vbCrLf & _
              "     %Title% will be replaced with the Title tag" & vbCrLf & _
              "     %Genre% will be replaced with the Genre tag" & vbCrLf & _
              "     %Year% will be replaced with the Year tag" & vbCrLf & _
              "     %Comment% will be replaced with the comment tag" & vbCrLf & _
              "" & vbCrLf & _
              ""
    Case """Rename Only if"" Check"
        Tmp = """Rename Only if"" Check" & vbCrLf & _
              "" & vbCrLf & _
              "         If this Checkbox is checked the Rename operation will only occur if the tags in the Format box are Present on the MP3. Otherwise it will be left as is" & vbCrLf & _
              ""
    Case "Rename"
        Tmp = "Rename" & vbCrLf & _
              "" & vbCrLf & _
              "         Pressing this button will use the Format tage to Rename all the Files Selected below" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: There is no UNDO" & vbCrLf & _
              ""
    Case "Replace"
        Tmp = "Replace" & vbCrLf & _
              "" & vbCrLf & _
              "         With this option you can select Before or After tag Rename Meaning Replacing before tagging or replacing whats in the tags too" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: There is no UNDO" & vbCrLf & _
              ""
    Case "Enter"
        Tmp = "Enter" & vbCrLf & _
              "" & vbCrLf & _
              "         This option Executed the Next»» Button. Thus will save then select the Next MP3 in the list" & vbCrLf & _
              "" & vbCrLf & _
              "(This shortcut only works when you are ReTagging One file at a time)" & vbCrLf & _
              ""
    Case "Ctrl+P", "Play Internally"
        Tmp = "Ctrl+P or Play Internally" & vbCrLf & _
              "" & vbCrLf & _
              "         I used Windows Media player to make a player contained in this Program." & vbCrLf & _
              "" & vbCrLf & _
              "         Useful only cause It wont open another application to play the song" & vbCrLf & _
              ""
    Case "Ctrl+O", "Open MP3"
        Tmp = "Ctrl+O or Open MP3" & vbCrLf & _
              "" & vbCrLf & _
              "         This opens the MP3 with the Associated MP3 Player" & vbCrLf & _
              ""

    Case "Ctrl+F", "Open Folder"
        Tmp = "Ctrl+F or Open Folder" & vbCrLf & _
              "" & vbCrLf & _
              "         This will open the Folder of the MP3 you have selected" & vbCrLf & _
              ""
    Case "F2", "Rename MP3"
        Tmp = "F2 or Rename MP3" & vbCrLf & _
              "" & vbCrLf & _
              "         This will allow you to Rename the Currently Selected MP3 Manually" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: There is no UNDO" & vbCrLf & _
              ""
    Case "Del", "Delete"
        Tmp = "Del or Delete" & vbCrLf & _
              "" & vbCrLf & _
              "         This will delete the Currently Selected MP3" & vbCrLf & _
              "" & vbCrLf & _
              "     WARNING: There is no UNDO" & vbCrLf & _
              ""
    Case "Drag & Drop"
        Tmp = "Drag & Drop" & vbCrLf & _
              "" & vbCrLf & _
              "         This function is Mostly used for Dragging to CD Writing Software. The only Function allowed with Drag & Drop is Copy so as not to lose the file by accidentally Dropping it into an Explorer Folder." & vbCrLf & _
              "" & vbCrLf & _
              "     HINT: Use the Find/Replace Dialog to find songs then Drag the Music Note to the Burning Software" & vbCrLf & _
              ""
End Select
Text1.Text = Tmp
End Sub


Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error Resume Next
'Rername the File Manually
Open ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text For Binary Access Write As #1
If Err.Number = 70 Then
    MsgBox "File in use"
    Cancel = True
    Err.Clear
    Close #1
Else
    Close #1
    Name ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text As _
         ListView1.SelectedItem.SubItems(1) & NewString
    If Err.Number <> 0 Then
        ERR_Add "Err:" & Err.Number & " " & Err.Description
        ERR_Add "  New Name " & ListView1.SelectedItem.Text
        ERR_Add "  Old Name " & NewString
        Err.Clear
        Cancel = True
    Else
        If Dir(ListView1.SelectedItem.SubItems(1) & NewString) = NewString Then
            'Stop
        Else
            Cancel = True
        End If
    End If
End If
    If DebugOn.Checked Then Err_Show Me
    Err_Clear
End Sub

Public Sub ParseFiles()
    'Parsing Files
    'All Manual Parsing
    ListView1.Sorted = False
    Dim Count As Double
    Dim I As Integer
    Dim TmpStr As String
    Dim TmpFile As String
    CancelIt = False
    ListView1.ListItems.Clear
    DirList1.Clear
    DirList2.Clear
    If Right(FolderPath.Caption, 1) <> "\" Then FolderPath.Caption = FolderPath.Caption & "\"
    If xMain.ValidateDir(FolderPath.Caption & "") = False Then
        FolderPath.Caption = App.Path
    End If
    TmpStr = Dir(FolderPath.Caption, vbNormal Or vbDirectory)
    Do Until TmpStr = ""
        DoEvents
If CancelIt Then Exit Sub
        StatBar.Panels(1).Text = "Parsing " & FolderPath.Caption
        DoEvents
        DirList1.AddItem TmpStr
        TmpStr = Dir()
    Loop
    For I = 0 To DirList1.ListCount - 1
        DoEvents
If CancelIt Then Exit Sub
        StatBar.Panels(1).Text = "Parsing " & FolderPath.Caption
        If DirList1.List(I) = ".." Or DirList1.List(I) = "." Then
            'Do Nothing
        ElseIf IsDir(FolderPath.Caption & DirList1.List(I)) Then
            If LCase(Right(DirList1.List(I), 4)) = ".mp3" Then Stop
            DirList2.AddItem DirList1.List(I) & "\"
        Else
            If LCase(Right(DirList1.List(I), 4)) = ".mp3" Then
                DoEvents
                ListView1.ListItems.Add , , DirList1.List(I)
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = FolderPath.Caption & TmpStr
                Count = Count + 1
            End If
        End If
    Next I
    If Check1.Value = vbChecked Then
        Do Until DirList2.ListCount = 0
            DoEvents
If CancelIt Then Exit Sub
            StatBar.Panels(1).Text = "Parsing " & FolderPath.Caption & TmpFile
            TmpFile = DirList2.List(0)
            DirList1.Clear
            TmpStr = Dir(FolderPath.Caption & TmpFile, vbNormal Or vbDirectory)
            Do Until TmpStr = ""
                DoEvents
If CancelIt Then Exit Sub
                StatBar.Panels(1).Text = "Parsing " & FolderPath.Caption & TmpFile
                DirList1.AddItem TmpStr
                TmpStr = Dir()
            Loop
            For I = 0 To DirList1.ListCount - 1
                DoEvents
If CancelIt Then Exit Sub
                StatBar.Panels(1).Text = "Parsing " & FolderPath.Caption & TmpFile
                If DirList1.List(I) = "." Or DirList1.List(I) = ".." Then
                    'Do Nothing
                ElseIf IsDir(FolderPath.Caption & TmpFile & DirList1.List(I)) Then
                    DirList2.AddItem TmpFile & DirList1.List(I) & "\"
                ElseIf LCase(Right(DirList1.List(I), 4)) = ".mp3" Then
                    DoEvents
                    ListView1.ListItems.Add , , DirList1.List(I)
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = FolderPath.Caption & TmpFile & TmpStr
                    Count = Count + 1
                End If
            Next I
            DirList2.RemoveItem 0
            If frmTagger.Visible = False Then
                If Parse.Visible = True Then
                    If Val(Left(Parse.FileCount.Caption, InStr(1, Parse.FileCount.Caption, " "))) <> Count Then
                        Parse.FileCount.Caption = Count & " Files found so far"
                    End If
                End If
            End If
        Loop
    Else
        
    End If
    StatBar.Panels(1).Text = "Total Files Parsed = " & ListView1.ListItems.Count
    ListView1.Sorted = True
    If ListView1.ListItems.Count > 0 Then ListView1_ItemClick ListView1.SelectedItem
End Sub

Private Sub ListView1_DblClick()
    'Play Internally
    If ListView1.ListItems.Count <> 0 Then
        If ListView1.MultiSelect = False Then
            mnuPlayInternally_Click
        End If
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    If ListView1.MultiSelect Then Exit Sub 'If Multi Do Nothing
    With MP3Tag
        'Load ID3 Info
        .Filename = item.SubItems(1) & item.Text
        .ReadTag
        Album = .Album
        Artist = .Artist
        Title = .Title
        If .ID3v2 Then
            Genre.Text = .GenreV2
        Else
            If .GenreV1 <= Genre.ListCount - 1 Then
                Genre.ListIndex = .GenreV1
            Else
                Genre.ListIndex = -1
            End If
        End If
        Year = .Year
        Comment = .Comments
        mp3info.Caption = "Channels :" & .Channels & vbCrLf
        mp3info.Caption = mp3info.Caption & "CopyWrighted :" & .Copyrighted & vbCrLf
        mp3info.Caption = mp3info.Caption & "Crc's :" & .CRCs & vbCrLf
        mp3info.Caption = mp3info.Caption & "Emphasis :" & .Emphasis & vbCrLf
        mp3info.Caption = mp3info.Caption & "Bytes :" & .FileBytes & vbCrLf
        mp3info.Caption = mp3info.Caption & "Frames :" & .Frames & vbCrLf
        mp3info.Caption = mp3info.Caption & "Hertz :" & .Hz & vbCrLf
        'MP3Info.Caption = MP3Info.Caption & "Info :" & .InfoString & vbCrLf
        mp3info.Caption = mp3info.Caption & "KBits :" & .Kbits & vbCrLf
        mp3info.Caption = mp3info.Caption & "Layer Version :" & .LayerVersion & vbCrLf
        mp3info.Caption = mp3info.Caption & "Mode :" & .Mode & vbCrLf
        mp3info.Caption = mp3info.Caption & "Mpeg Version :" & .MpegVersion & vbCrLf
        mp3info.Caption = mp3info.Caption & "Original :" & .Original & vbCrLf
        mp3info.Caption = mp3info.Caption & "Seconds :" & .Seconds & vbCrLf
        mp3info.Caption = mp3info.Caption & "Time :" & Int(0.5 + (.Seconds / 60)) & ":" & (.Seconds Mod 60) & vbCrLf
    End With
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sMenu As Long
    Dim hMenu As Long
    Dim hMenuID As Long
    'Been trying to Skin the Menu with no Luck
    sMenu = GetMenu(Me.hwnd)
    hMenu = GetSubMenu(sMenu, 1)
    hMenuID = GetMenuItemID(hMenu, 1)
    
    If Button = 2 Then 'If Right click
            PopupMenu mnuListView, , , , mnuPlayInternally 'Popup Menu
    End If
End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Dim DD As Integer
    'Allow Drag and Drop (good for Dropping into Other Programs)
    If ListView1.ListItems.Count = 0 Then Exit Sub
    AllowedEffects = vbDropEffectCopy
    Data.SetData , vbCFFiles
    If Check3.Value = vbChecked Then 'If Multi
        For DD = 1 To ListView1.ListItems.Count
            'Add Selected
            If ListView1.ListItems(DD).Selected Then Data.Files.Add ListView1.ListItems(DD).SubItems(1) & ListView1.ListItems(DD).Text
        Next DD
    Else 'If Not
        Data.Files.Add ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text 'Add Selected
    End If
End Sub

Private Sub cReplace_Click()
    frmReplace.Show vbModal, Me 'Show Replace Dialog
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    Dim MP3Tmp As cIDV3
    Set MP3Tmp = New cIDV3
    Select Case NewState
        Case 0 'Stop
            frmedia.Visible = False
            MediaPlayer1.Filename = ""
            mnuStop.Enabled = False
            Me.Stop.Value = True
        Case 1 'Pause
            PMode(NewState).Value = True
        Case 2 'Play
            With MP3Tmp
                .Filename = MediaPlayer1.Filename
                .ReadTag
                SongTitle = .Title
                SongArtist = .Artist
            End With
            PMode(NewState).Value = True
        Case Else
            Debug.Assert False
            mnuStop.Enabled = True
    End Select
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal 'Show About Dialog
End Sub

Private Sub mnuDelete_Click()
Dim TmpMP3 As cIDV3
Set TmpMP3 = New cIDV3
'Make sure
If ListView1.MultiSelect = True Or ListView1.ListItems.Count = 0 Then Exit Sub
With TmpMP3
    .Filename = ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
    .ReadTag
    'Verify with User
    Select Case MsgBox(Replace("Deleting " & ListView1.SelectedItem.Text & vbCrLf & _
                       "" & vbCrLf & _
                       "Album: " & .Album & vbCrLf & _
                       "Artist: " & .Artist & vbCrLf & _
                       "Title: " & .Title & vbCrLf & _
                       "Genre: " & IIf(.ID3v2, .GenreV2, .GenreText(.GenreV1)) & vbCrLf & _
                       "Year: " & .Year & vbCrLf & _
                       "Comments: " & .Comments & vbCrLf & _
                       "" & vbCrLf & _
                       "Are you sure you want to Delete this ?", Chr(0), ""), vbYesNo, "Confirm")
        Case vbYes
            'Kill the File till its gone
            Do Until Dir(ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text) = ""
                Kill ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
            Loop
            'Remove Deleted Item from the Listview
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
            If ListView1.ListItems.Count <> 0 Then
                ListView1.SelectedItem.Selected = True
            End If
    End Select
End With
End Sub

Private Sub mnuHlp_Click()
Dim I As Integer
For I = 1 To TabStrip1.Tabs.Count 'Look through the tabs
    If TabStrip1.Tabs(I).Caption = "Help" Then 'Help Tab
        TabStrip1.Tabs(I).Selected = True 'Select
        TabStrip1_Click 'Click Event
        Exit For
    End If
Next I
End Sub

Private Sub mnuOpen_Click()
    'Just to make sure
    If ListView1.MultiSelect = True Then Exit Sub
    If ListView1.ListItems.Count <> 0 Then
        'Open with Default Program
        ShellExecute Me.hwnd, "open", ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text, vbNullString, "C:\", SW_SHOWNORMAL
    End If
End Sub

Private Sub mnuOpenFolder_Click()
    'Just to make sure
    If ListView1.MultiSelect = True Then Exit Sub
    If ListView1.ListItems.Count <> 0 Then
        'Open Folder
        ShellExecute Me.hwnd, "open", ListView1.SelectedItem.SubItems(1), vbNullString, "C:\", SW_SHOWNORMAL
    End If
End Sub

Private Sub mnuPlayInternally_Click()
On Error Resume Next
    'Just to make sure
    If ListView1.MultiSelect = False Then
        'Adding the Filename Plays the Song
        MediaPlayer1.Filename = ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
        frmedia.Visible = True
    End If
End Sub

Private Sub mnuRename_Click()
    'Just to make sure
    If ListView1.MultiSelect = True Or ListView1.ListItems.Count = 0 Then Exit Sub
    'Allow Label Edit
    ListView1.StartLabelEdit
End Sub

Private Sub mnuStop_Click()
    'Stop
    MediaPlayer1.Stop
    MediaPlayer1.Filename = ""
End Sub

Private Sub Mute_Click(Index As Integer)
    'Mute Checkboxes
    If Index = 7 Then
        MediaPlayer1.Mute = Mute(Index).Value
    Else
        If Me.Visible = True Then Sound11(Index).Mute = Mute(Index).Value
    End If
End Sub

Private Sub PMode_Click(Index As Integer)
    'Play Buttons
    Select Case Index
        Case 0 'Stop
            
            MediaPlayer1.Stop
            MediaPlayer1.Filename = ""
            frmedia.Visible = False
        Case 1 'Pause
            If MediaPlayer1.Filename = "" Or Dir(MediaPlayer1.Filename) = "" Then
                PMode(0).Value = True
            Else
                MediaPlayer1.Pause
            End If
        Case 2 'Play
            If MediaPlayer1.Filename = "" Or Dir(MediaPlayer1.Filename) = "" Then
                If ListView1.ListItems.Count <> 0 Then
                    MediaPlayer1.Filename = ListView1.SelectedItem.SubItems(1) & ListView1.SelectedItem.Text
                    MediaPlayer1.Play
                    frmedia.Visible = True
                Else
                    PMode(0).Value = True
                End If
            Else
                MediaPlayer1.Play
                frmedia.Visible = True
            End If
    End Select
End Sub

Private Sub Respond_Timer()
    'Remote Command Timer
    Dim Tmp As String
    Respond.Enabled = False
    Dim sTmp() As String
    Dim I As Integer
    Dim I2 As Integer
    If List4.ListCount = 0 Then Exit Sub
    For I = 0 To List4.ListCount
        If List4.List(I) <> "" Then
            sTmp = Split(List4.List(I), ":")
            I2 = UBound(sTmp)
            Select Case UCase(sTmp(0))
                Case "COMM"
                    'Select Case UCase(sTmp(1))
                    
            End Select
        End If
    Next I
End Sub

Private Sub SkinForm1_OnSkinNotify(ByVal SkinClass As String, ByVal SkinEvent As String)
    'Print Skin Info
    Debug.Print "OnSkinNotify Class:" & SkinClass & " Event:" & SkinEvent
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    'Volume Scrolling
    Select Case Index
        Case 7
            'MediaPlayer1
            MediaPlayer1.Volume = Slider1(Index).Value - 1000
        Case Else
            Sound11(Index).Volume = Slider1(Index).Value
    End Select
End Sub

Private Sub Sound11_MuteChanged(Index As Integer, NewMute As Boolean)
    'Mute Changed update the Mute Checkboxes
    If Index > 6 Then Exit Sub
    Select Case NewMute
        Case True
            Mute(Index).Value = vbChecked
        Case Else
            Mute(Index).Value = vbUnchecked
    End Select
End Sub

Private Sub Sound11_VolumeChanged(Index As Integer, NewVolume As Long)
    'Volume changed. Update Sliders
    Slider1(Index).Value = NewVolume
End Sub

Private Sub Stop_Click()
    'Stop
    MediaPlayer1.Stop
    MediaPlayer1.Filename = ""
    'Hide media Player
    frmedia.Visible = False
    PMode(1).Value = False
    PMode(2).Value = False
    
End Sub

Private Sub TabStrip1_Click()
'Tabs Info
frTab1.Visible = False
frTab2.Visible = False
frTab3.Visible = False
frTab4.Visible = False
frTab5.Visible = False
    'Show Per tab
    Select Case TabStrip1.SelectedItem.Caption
        Case "Retag"
            frTab1.Visible = True
        Case "Rename"
            frTab2.Visible = True
        Case "About"
            frTab3.Visible = True
        Case "Help"
            frTab4.Visible = True
        Case "Player"
            frTab5.Visible = True
    End Select
End Sub

Private Sub txtReName_Change()
    'Nothing to Validate Directory yet
    If InStr(1, txtReName, "\", vbTextCompare) > 0 Then
        'ValidateDir.Visible = True
    Else
        'ValidateDir.Visible = False
    End If
End Sub

Private Sub txtReName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("\")
            MsgBox "Directories in Renaming is Not Correctly Implimented" & vbCrLf & _
                   "This means it May not Work Correctly At this time." & vbCrLf & _
                   "I have Disabled it For Now" & vbCrLf & _
                   "Please Check Later Versions for This Feature" _
                   , vbInformation, "Warning"
            KeyAscii = 0
    End Select
End Sub

Public Sub HideAllSecrets()
    'Hiding Everything Here
    'No Secrets to hide
End Sub

