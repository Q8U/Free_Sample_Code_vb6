VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Icon Editor"
   ClientHeight    =   6990
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbProperties 
      Height          =   405
      Left            =   4950
      TabIndex        =   29
      Top             =   405
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbSolid"
            Object.ToolTipText     =   "Solid"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbDash"
            Object.ToolTipText     =   "Dash"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbDot"
            Object.ToolTipText     =   "Dot"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbDash-Dot"
            Object.ToolTipText     =   "Dash-Dot"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbDash-Dot-Dot"
            Object.ToolTipText     =   "Dash-Dot-Dot"
            ImageIndex      =   31
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1440
      TabIndex        =   27
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tlbTools 
      Height          =   405
      Left            =   0
      TabIndex        =   26
      Top             =   405
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbPencil"
            Object.ToolTipText     =   "Pencil"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbEraser"
            Object.ToolTipText     =   "Eraser"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tlbFillColor"
            Object.ToolTipText     =   "Fill Color"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbLine"
            Object.ToolTipText     =   "Line"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbSquare"
            Object.ToolTipText     =   "Square"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbSquareFill"
            Object.ToolTipText     =   "Square Fill"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbRectangle"
            Object.ToolTipText     =   "Rectangle"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbRectangleFill"
            Object.ToolTipText     =   "Rectangle Fill"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbCircle"
            Object.ToolTipText     =   "Circle"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbCircleFill"
            Object.ToolTipText     =   "Circle Fill"
            ImageIndex      =   26
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picForSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2040
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog cdbDialog 
      Left            =   2040
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSave 
      Left            =   2040
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   15
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   24
      Top             =   5985
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   14
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   23
      Top             =   5760
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   13
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   22
      Top             =   5535
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   21
      Top             =   5310
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   11
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   20
      Top             =   5085
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   10
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   19
      Top             =   4860
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   18
      Top             =   4635
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   17
      Top             =   4410
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   16
      Top             =   4185
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   15
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   14
      Top             =   3735
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   13
      Top             =   3510
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   12
      Top             =   3285
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   11
      Top             =   3060
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   10
      Top             =   2835
      Width           =   1335
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1305
      TabIndex        =   9
      Top             =   2610
      Width           =   1335
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      Height          =   450
      Left            =   480
      ScaleHeight     =   390
      ScaleWidth      =   840
      TabIndex        =   8
      ToolTipText     =   "Click to Change Fore Color"
      Top             =   2085
      Width           =   900
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Shadow"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1785
      Width           =   1215
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Back Color"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1485
      Width           =   1215
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Fore Color"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1185
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame fraColors 
      Caption         =   "Colors"
      Height          =   5505
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Colors (Fore Color, Background Color and Shadow Color)"
      Top             =   960
      Width           =   1575
      Begin VB.PictureBox picColors 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   16
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   1305
         TabIndex        =   30
         Top             =   5250
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbNew"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbOpen"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbSave"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbUndo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbLeft"
            Object.ToolTipText     =   "Left"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbUp"
            Object.ToolTipText     =   "Up"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbDown"
            Object.ToolTipText     =   "Down"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbRight"
            Object.ToolTipText     =   "Right"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbRotateCW"
            Object.ToolTipText     =   "Rotate Clockwise"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbRotateUCW"
            Object.ToolTipText     =   "Rotate UnClockwise"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbShadow"
            Object.ToolTipText     =   "Shadow"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbChangeColor"
            Object.ToolTipText     =   "Change Color"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbChangeColors"
            Object.ToolTipText     =   "Change Colors"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tlbInverseColors"
            Object.ToolTipText     =   "Inverse Colors"
            ImageIndex      =   24
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   750
      Left            =   1920
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   1
      ToolTipText     =   "Icon Preview"
      Top             =   1080
      Width           =   750
   End
   Begin VB.PictureBox picIconArea 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   2880
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":162A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":194E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":250A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3096
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4386
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5352
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5896
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":631E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":760E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8096
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9062
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPosition 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo 0 time(s)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "&Colors"
      Begin VB.Menu mnuColorsColor 
         Caption         =   "&Fore Color"
         Index           =   0
      End
      Begin VB.Menu mnuColorsColor 
         Caption         =   "&Back Color"
         Index           =   1
      End
      Begin VB.Menu mnuColorsColor 
         Caption         =   "&Shadow Color"
         Index           =   2
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTools 
         Caption         =   "&Tools"
         Begin VB.Menu mnuViewToolsUp 
            Caption         =   "&Up"
         End
         Begin VB.Menu mnuViewToolsDown 
            Caption         =   "&Down"
         End
         Begin VB.Menu mnuViewToolsLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuViewToolsRight 
            Caption         =   "&Right"
         End
         Begin VB.Menu mnuViewToolsBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolsRotateCW 
            Caption         =   "R&otate Clockwise"
         End
         Begin VB.Menu mnuViewToolsRotateUCW 
            Caption         =   "Ro&tate UnClockwise"
         End
         Begin VB.Menu mnuViewToolsBar2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolsShaodw 
            Caption         =   "&Shadow"
         End
         Begin VB.Menu mnuViewToolsBar3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolsChangeColor 
            Caption         =   "C&hange Color"
         End
         Begin VB.Menu mnuViewToolsChangeColors 
            Caption         =   "Ch&ange Colors"
         End
         Begin VB.Menu mnuViewToolsInverseColors 
            Caption         =   "In&verse Colors"
         End
      End
      Begin VB.Menu mnuViewDrawTools 
         Caption         =   "&Draw Tools"
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Pencil"
            Index           =   1
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Eraser"
            Index           =   2
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Color Fill"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Line"
            Index           =   5
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Square"
            Index           =   6
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "S&quare Fill"
            Index           =   7
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "&Rectangle"
            Index           =   8
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "Rec&tangle Fill"
            Index           =   9
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "C&ircle"
            Index           =   10
         End
         Begin VB.Menu mnuViewDrawToolsTool 
            Caption         =   "Circle &Fill"
            Index           =   11
         End
      End
      Begin VB.Menu mnuViewLineProperties 
         Caption         =   "&Line Properties"
         Begin VB.Menu mnuViewLinePropertiesTool 
            Caption         =   "&Solid"
            Index           =   1
         End
         Begin VB.Menu mnuViewLinePropertiesTool 
            Caption         =   "&Dash"
            Index           =   2
         End
         Begin VB.Menu mnuViewLinePropertiesTool 
            Caption         =   "D&ot"
            Index           =   3
         End
         Begin VB.Menu mnuViewLinePropertiesTool 
            Caption         =   "D&ash-Dot"
            Index           =   4
         End
         Begin VB.Menu mnuViewLinePropertiesTool 
            Caption         =   "Da&sh-Dot-Dot"
            Index           =   5
         End
      End
      Begin VB.Menu mnuViewBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intMaxUndoNumber As Integer
Public blnKeepEdgeDate As Boolean
Public lngBackColor, lngForeColor, lngShadowColor, lngLineColor, lngMouseOnColor, lngSelectColor As Long
Public blnMakeSelection As Boolean
Public intScale As Integer
Public intMaxXY As Integer
Private intI, intJ As Integer
Private lngIcon(1 To 32, 1 To 32) As Long
Public blnIconChanged As Boolean
Public intRow, intCol As Integer
Public strFileName As String
Private lngIconTemp(1 To 32, 1 To 32) As Long
Public intShadowSpace As Integer
Public intToolNumber As Integer
Private intPointX1, intPointY1, intPointX2, intPointY2 As Integer
Public intDrawStyle As Integer
Private lngUndoIcon(1 To 10, 1 To 32, 1 To 32) As Long
Public intUndoNumber, intUndo As Integer
Public intLastRow, intLastCol As Integer

Private Sub CreateUndoIcon()
  Dim intI, intJ As Integer
  
  intUndo = intUndo + 1
  If (intUndo = intMaxUndoNumber + 1) Then intUndo = 1
  
  If (intUndoNumber < intMaxUndoNumber) Then intUndoNumber = intUndoNumber + 1
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngUndoIcon(intUndo, intI, intJ) = lngIcon(intI, intJ)
    Next intJ
  Next intI
  
  If (intUndoNumber = 2) Then
    tlbMain.Buttons(5).Enabled = True
    mnuEditUndo.Enabled = True
  End If
  mnuEditUndo.Caption = "&Undo " & intUndoNumber - 1 & " time(s)"
End Sub

Private Sub ViewOptions()
  frmOptions.hsbShadowSpace.Value = intShadowSpace
  frmOptions.lblShaodwSpaceNumber.Caption = intShadowSpace
  frmOptions.intButtonClicked = vbYes
  If (blnKeepEdgeDate = True) Then
    frmOptions.chkKeepData.Value = 1
  Else
    frmOptions.chkKeepData.Value = 0
  End If
  frmOptions.Show (1)
  
  If (frmOptions.intButtonClicked = vbOK) Then
    intShadowSpace = frmOptions.hsbShadowSpace.Value
    
    If (frmOptions.chkKeepData.Value = 1) Then
      blnKeepEdgeDate = True
    Else
      blnKeepEdgeDate = False
    End If
  End If
End Sub

Private Sub NewIcon()
  Dim intI As Integer
  Dim intResult As Integer

  If (blnIconChanged = True) Then
    intResult = MsgBox("Icon changed! Save it?", vbYesNoCancel + vbDefaultButton1 + vbQuestion, "Save Changes")
    If (intResult = vbCancel) Then
      Exit Sub
    End If
    If (intResult = vbYes) Then
      Call SaveIcon
    End If
  End If
  
  lngBackColor = QBColor(7)
  lngForeColor = QBColor(0)
  lngLineColor = vbBlack
  lngMouseOnColor = vbWhite
  lngSelectColor = vbRed
  lngShadowColor = QBColor(8)
  blnIconChanged = False
  blnMakeSelection = False
  strFileName = ""
  intToolNumber = 1
  intDrawStyle = 1
  intMaxUndoNumber = 10
  intUndo = 0
  intUndoNumber = 0
  intLastRow = 1
  intLastCol = 1
  blnKeepEdgeDate = True
  
  For Each btnTemp In tlbTools.Buttons
    btnTemp.Value = tbrPressed
    btnTemp.Value = tbrUnpressed
  Next btnTemp
  tlbTools.Buttons(intDrawStyle).Value = tbrPressed
  
  For intI = 1 To mnuViewDrawToolsTool.Count
    mnuViewDrawToolsTool(intI).Checked = False
  Next intI
  mnuViewDrawToolsTool(intToolNumber).Checked = True
  
  For Each btnTemp In tlbProperties.Buttons
    btnTemp.Value = tbrUnpressed
    btnTemp.Enabled = False
  Next btnTemp
  tlbProperties.Buttons(1).Value = tbrPressed
  
  For intI = 1 To mnuViewLinePropertiesTool.Count
    mnuViewLinePropertiesTool(intI).Checked = False
    mnuViewLinePropertiesTool(intI).Enabled = False
  Next intI
  mnuViewLinePropertiesTool(intDrawStyle).Checked = True
  
  mnuEditUndo.Enabled = False
  mnuEditCut.Enabled = False
  mnuEditCopy.Enabled = False
  mnuEditPaste.Enabled = False
  
  picIconArea.BackColor = lngBackColor
  picIcon.BackColor = lngBackColor
  
  For intI = 0 To intMaxXY
    picIconArea.Line (intI * intScale, 0)-(intI * intScale, picIconArea.Height), lngLineColor
    picIconArea.Line (0, intI * intScale)-(picIconArea.Width, intI * intScale), lngLineColor
  Next intI
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngBackColor
    Next intJ
  Next intI
  
  picColor.BackColor = lngForeColor
  tlbMain.Buttons(5).Enabled = False
  CreateUndoIcon
End Sub

Private Sub cmdAbout_Click()
  Call mnuHelpAbout_Click
End Sub

Private Sub cmdExit_Click()
  Call mnuFileExit_Click
End Sub

Private Sub Form_Load()
  If (App.PrevInstance = True) Then
    Call MsgBox("Sorry! this program works only once each time", vbOKOnly + vbExclamation)
    Unload Me
    End
  End If
  
  intScale = 12
  intMaxXY = 32
  intShadowSpace = 2
  blnIconChanged = False
    
  For intI = 0 To 15
    picColors(intI).BackColor = QBColor(15 - intI)
    picColors(intI).ToolTipText = "Color (" & 15 - intI & ") :" & QBColor(15 - intI)
  Next intI
  picColors(16).BackColor = &H8000000F

  mnuColorsColor(0).Checked = True
  
  lblPosition.Caption = "Row: 00 Col:   00"
  Call NewIcon
  
  If (Command <> "") Then
    strFileName = Command
    OpenIcon (strFileName)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mnuFileExit_Click
End Sub

Private Sub mnuColorsColor_Click(Index As Integer)
  Call optColor_Click(Index)
End Sub

Private Sub mnuEditUndo_Click()
  Call UndoIcon
End Sub

Private Sub mnuFileExit_Click()
  Dim intResult As Integer
  
  If (blnIconChanged = True) Then
    intResult = MsgBox("Icon changed! Save it before exit?", vbYesNoCancel + vbDefaultButton1 + vbQuestion, "Save Changes")
  End If
  If (intResult = vbCancel) Then
    Exit Sub
  End If
  If (intResult = vbYes) Then
    Call SaveIcon
  End If
  blnIconChanged = False
  
  Unload Me
  End
End Sub

Private Sub mnuFileNew_Click()
  Call NewIcon
End Sub

Private Sub OpenIcon(ByVal strFileName As String)
  Dim intI, intJ As Integer

  'If FileLen(strFileName) <> 766 Then
  '  MsgBox "Invalid or unsupported file format.", vbCritical
  '  Exit Sub
 ' End If
    
  picForSave = LoadPicture(strFileName)
    
  For intI = 0 To intMaxXY - 1
    For intJ = 0 To intMaxXY - 1
      lngIcon(intI + 1, intJ + 1) = picForSave.Point(intI, intJ)
    Next intJ
  Next intI
    
  Call DrawChangedIcon
    
  blnIconChanged = False
End Sub

Private Sub mnuFileOpen_Click()
  Dim intResult As Integer

  If (blnIconChanged = True) Then
    intResult = MsgBox("Icon changed! Save it?", vbYesNoCancel + vbDefaultButton1 + vbQuestion, "Save Changes")
  End If
  If (intResult = vbCancel) Then
    Exit Sub
  End If
  If (intResult = vbYes) Then
    Call SaveIcon
  End If
  
  cdbDialog.DialogTitle = "Open File"
  cdbDialog.FileName = ""
  cdbDialog.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
  cdbDialog.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp"
  cdbDialog.ShowOpen
  
  If (cdbDialog.FileName <> "") Then
    strFileName = cdbDialog.FileName
    
    Call OpenIcon(strFileName)
  End If
End Sub

Private Sub mnuFileSave_Click()
  Call SaveIcon
End Sub

Private Sub mnuFileSaveAs_Click()
  Dim strNewFileName As String
  
  strNewFileName = strFileName
  strFileName = ""
  Call SaveAsIcon
  If (strFileName <> "") Then
    SaveIcon
  Else
    strFileName = strNewFileName
  End If
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.lblDescription = "This program can help you to create your own ICONS and save theme in standard ICO or BMP format for windows."
  frmAbout.lblDisclaimer = "Attention: Icons can be only 32x32 (16 Color)"
  frmAbout.lblProgrammer = "Ali (Maziar) Amirnezhad (amirnezhad@yahoo.com)"
  frmAbout.Show (1)
End Sub

Private Sub mnuViewDrawToolsTool_Click(Index As Integer)
  mnuViewDrawToolsTool(intToolNumber).Checked = False
  tlbTools.Buttons(intToolNumber).Value = tbrPressed
  tlbTools.Buttons(intToolNumber).Value = tbrUnpressed
  
  intToolNumber = Index
  
  mnuViewDrawToolsTool(intToolNumber).Checked = True
  tlbTools.Buttons(intToolNumber).Value = tbrPressed

  For Each btnTemp In tlbProperties.Buttons
    btnTemp.Enabled = False
    mnuViewLinePropertiesTool(btnTemp.Index).Enabled = False
  Next btnTemp
  
  If (intToolNumber >= 5) Then
    For Each btnTemp In tlbProperties.Buttons
      btnTemp.Enabled = True
      mnuViewLinePropertiesTool(btnTemp.Index).Enabled = True
    Next btnTemp
  End If
End Sub

Private Sub mnuViewLinePropertiesTool_Click(Index As Integer)
  mnuViewLinePropertiesTool(intDrawStyle).Checked = False
  tlbProperties.Buttons(intDrawStyle).Value = tbrPressed
  tlbProperties.Buttons(intDrawStyle).Value = tbrUnpressed
  
  intDrawStyle = Index
  
  mnuViewLinePropertiesTool(intDrawStyle).Checked = True
  tlbProperties.Buttons(intDrawStyle).Value = tbrPressed
End Sub

Private Sub mnuViewOptions_Click()
  Call ViewOptions
End Sub

Private Sub mnuViewToolsDown_Click()
  Call DownIcon
End Sub

Private Sub mnuViewToolsInverseColors_Click()
  Call InverseColorsIcon
End Sub

Private Sub mnuViewToolsLeft_Click()
  Call LeftIcon
End Sub

Private Sub mnuViewToolsRight_Click()
  Call RightIcon
End Sub

Private Sub mnuViewToolsRotateCW_Click()
  Call RotateCWIcon
End Sub

Private Sub mnuViewToolsRotateUCW_Click()
  Call RotateUCWIcon
End Sub

Private Sub mnuViewToolsChangeColor_Click()
  Call ChangeColorIcon
End Sub

Private Sub mnuViewToolsChangeColors_Click()
  Call ChangeColorsIcon
End Sub

Private Sub mnuViewToolsShaodw_Click()
  Call ShadowIcon
End Sub

Private Sub mnuViewToolsUp_Click()
  Call UpIcon
End Sub

Private Sub optColor_Click(Index As Integer)
  optColor(Index).Value = True
  
  If (Index = 0) Then
    picColor.BackColor = lngForeColor
    picColor.ToolTipText = "Fore Color"
  ElseIf (Index = 1) Then
    picColor.BackColor = lngBackColor
    picColor.ToolTipText = "Back Color"
  Else
    picColor.BackColor = lngShadowColor
    picColor.ToolTipText = "Shadow Color"
  End If
  
  For intI = 0 To 2
    mnuColorsColor(intI).Checked = False
  Next intI
  mnuColorsColor(Index).Checked = True
End Sub

Private Sub picColors_Click(Index As Integer)
  picColor.BackColor = picColors(Index).BackColor
  
  If (optColor(0).Value = True) Then
    lngForeColor = picColors(Index).BackColor
  ElseIf (optColor(1).Value = True) Then
    lngBackColor = picColors(Index).BackColor
  Else
    lngShadowColor = picColors(Index).BackColor
  End If
End Sub

Private Sub picIconArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (intToolNumber < 4) Then
    Call picIconArea_MouseMove(Button, Shift, X, Y)
    Exit Sub
  End If
  
  If (Button = vbLeftButton) Then
    intPointX1 = X \ intScale + 1
    intPointY1 = Y \ intScale + 1
  
    If (intPointX1 < 1) Then
      intPointX1 = 1
    End If
    If (intPointX1 > intMaxXY) Then
      intPointX1 = intMaxXY
    End If
    If (intPointY1 < 1) Then
      intPointY1 = 1
    End If
    If (intPointY1 > intMaxXY) Then
     intPointY1 = intMaxXY
    End If
    
    blnMakeSelection = True
  Else
    intPointX1 = 0
    intPointY1 = 0
  End If
End Sub

Private Sub picIconArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim strRow, strCol As String
  Dim lngColor As Long
  Dim intX1, intY1, intX2, intY2 As Integer

  intRow = X \ intScale + 1
  intCol = Y \ intScale + 1

  If (intRow < 1) Then
    intRow = 1
  End If
  If (intRow > intMaxXY) Then
    intRow = intMaxXY
  End If
  If (intCol < 1) Then
    intCol = 1
  End If
  If (intCol > intMaxXY) Then
    intCol = intMaxXY
  End If
  
  If (blnMakeSelection = True) Then
    picIconArea.FillStyle = 1
    If (intPointX1 > intLastRow) Then
      intX1 = intLastRow
      intX2 = intPointX1
    Else
      intX1 = intPointX1
      intX2 = intLastRow
    End If
    If (intPointY1 > intLastCol) Then
      intY1 = intLastCol
      intY2 = intPointY1
    Else
      intY1 = intPointY1
      intY2 = intLastCol
    End If
    picIconArea.Line ((intX1 - 1) * 12, (intY1 - 1) * 12)-(intX2 * 12, intY2 * 12), lngLineColor, B
  
    If (intPointX1 > intRow) Then
      intX1 = intRow
      intX2 = intPointX1
    Else
      intX1 = intPointX1
      intX2 = intRow
    End If
    If (intPointY1 > intCol) Then
      intY1 = intCol
      intY2 = intPointY1
    Else
      intY1 = intPointY1
      intY2 = intCol
    End If
    picIconArea.Line ((intX1 - 1) * 12, (intY1 - 1) * 12)-(intX2 * 12, intY2 * 12), lngSelectColor, B
  End If
  If ((intLastRow <> intRow) Or (intLastCol <> intCol)) Then
    picIconArea.FillStyle = 1
    picIconArea.Line ((intLastRow - 1) * 12, (intLastCol - 1) * 12)-(intLastRow * 12, intLastCol * 12), lngLineColor, B
    
    intLastRow = intRow
    intLastCol = intCol
    
    picIconArea.Line ((intLastRow - 1) * 12, (intLastCol - 1) * 12)-(intLastRow * 12, intLastCol * 12), lngMouseOnColor, B
  End If
  
  strRow = LTrim(Str(intCol))
  strCol = LTrim(Str(intRow))

  If (intCol < 10) Then
    strRow = "0" & strRow
  End If
  If (intRow < 10) Then
    strCol = "0" & strCol
  End If
  lblPosition.Caption = "Row: " & strRow & " Col:   " & strCol
  
  If (intToolNumber = 1) Then
    If ((Button = vbLeftButton) Or (Button = vbRightButton)) Then
      If (Button = vbLeftButton) Then
        lngColor = lngForeColor
      Else
        lngColor = lngBackColor
      End If
    
      picIcon.PSet (intRow + 7, intCol + 7), lngColor
      lngIcon(intRow, intCol) = lngColor
    
      intX1 = (intRow - 1) * intScale + 1
      intY1 = (intCol - 1) * intScale + 1
      intX2 = intRow * 12 - 1
      intY2 = intCol * 12 - 1
      picIconArea.FillColor = lngColor
      picIconArea.FillStyle = 0
      picIconArea.Line (intX1, intY1)-(intX2, intY2), lngColor, B
    
      blnIconChanged = True
    End If
  End If
  
  If (intToolNumber = 2) Then
    If (Button = vbLeftButton) Then
      picIcon.PSet (intRow + 7, intCol + 7), QBColor(7)
      lngIcon(intRow, intCol) = QBColor(7)
      
      intX1 = (intRow - 1) * intScale + 1
      intY1 = (intCol - 1) * intScale + 1
      intX2 = intRow * 12 - 1
      intY2 = intCol * 12 - 1
      picIconArea.FillColor = QBColor(7)
      picIconArea.FillStyle = 0
      picIconArea.Line (intX1, intY1)-(intX2, intY2), QBColor(7), B
    
      blnIconChanged = True
    End If
  End If
End Sub
Private Sub SaveAsIcon()
  cdbDialog.DialogTitle = "Save File"
  cdbDialog.FileName = ""
  cdbDialog.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
  cdbDialog.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp"
  cdbDialog.ShowSave
    
  strFileName = cdbDialog.FileName
End Sub

Private Sub SaveIcon()
  Dim intI, intJ As Integer

  If (strFileName = "") Then
    Call SaveAsIcon
  End If
  If (strFileName <> "") Then
    For intI = 1 To intMaxXY
      For intJ = 1 To intMaxXY
        picForSave.PSet (intI - 1, intJ - 1), lngIcon(intI, intJ)
      Next intJ
    Next intI
    
     
    If (cdbDialog.FilterIndex = 1) Then
      Dim imgX As ListImage
      Set imgX = imlSave.ListImages.Add(1, , picForSave.Image)
      
      Dim picX As Picture
      Set picX = imlSave.ListImages(1).ExtractIcon
      Call SavePicture(picX, strFileName)
    Else
      Call SavePicture(picForSave.Image, strFileName)
    End If
    blnIconChanged = False
  End If
End Sub

Private Sub DrawChangedIcon()
  Dim intI, intJ As Integer
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      picIcon.PSet (intI + 7, intJ + 7), lngIcon(intI, intJ)
      picIconArea.FillColor = lngIcon(intI, intJ)
      picIconArea.FillStyle = 0
      picIconArea.Line ((intI - 1) * 12 + 1, (intJ - 1) * 12 + 1)-(intI * 12 - 1, intJ * 12 - 1), lngIcon(intI, intJ), B
    Next intJ
  Next intI
End Sub

Private Sub CreateTempIcon()
  Dim intI, intJ As Integer
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIconTemp(intI, intJ) = lngIcon(intI, intJ)
    Next intJ
  Next intI
End Sub

Private Sub UpIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY - 1
      lngIcon(intI, intJ) = lngIconTemp(intI, intJ + 1)
    Next intJ
  Next intI
  For intI = 1 To intMaxXY
    If (blnKeepEdgeDate = True) Then
      lngIcon(intI, 32) = lngIconTemp(intI, 1)
    Else
      lngIcon(intI, 32) = QBColor(7)
    End If
  Next intI
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub
Private Sub DownIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY
    For intJ = 2 To intMaxXY
      lngIcon(intI, intJ) = lngIconTemp(intI, intJ - 1)
    Next intJ
  Next intI
  For intI = 1 To intMaxXY
    If (blnKeepEdgeDate = True) Then
      lngIcon(intI, 1) = lngIconTemp(intI, 32)
    Else
      lngIcon(intI, 1) = QBColor(7)
    End If
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub LeftIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY - 1
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngIconTemp(intI + 1, intJ)
    Next intJ
  Next intI
  For intI = 1 To intMaxXY
    If (blnKeepEdgeDate = True) Then
      lngIcon(32, intI) = lngIconTemp(1, intI)
    Else
      lngIcon(32, intI) = QBColor(7)
    End If
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub RightIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 2 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngIconTemp(intI - 1, intJ)
    Next intJ
  Next intI
  For intI = 1 To intMaxXY
    If (blnKeepEdgeDate = True) Then
      lngIcon(1, intI) = lngIconTemp(32, intI)
    Else
      lngIcon(1, intI) = QBColor(7)
    End If
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub RotateUCWIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngIconTemp(intMaxXY + 1 - intJ, intI)
    Next intJ
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub RotateCWIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngIconTemp(intJ, intMaxXY + 1 - intI)
    Next intJ
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub ShadowIcon()
  Dim intI, intJ As Integer
  
  Call CreateTempIcon
  
  For intI = 1 To intMaxXY - intShadowSpace
    For intJ = 1 To intMaxXY - intShadowSpace
      If (lngIconTemp(intI, intJ) <> QBColor(7)) Then
        lngIcon(intI + intShadowSpace, intJ + intShadowSpace) = lngShadowColor
      End If
    Next intJ
  Next intI
  
  For intI = 1 To intMaxXY - 1
    For intJ = 1 To intMaxXY - 1
      If (lngIconTemp(intI, intJ) <> QBColor(7)) Then
        lngIcon(intI, intJ) = lngIconTemp(intI, intJ)
      End If
    Next intJ
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub ChangeColorsIcon()
  frmSelectColors.lblStatus.Caption = "<=>"
  frmSelectColors.Show (1)
  
  Call CreateTempIcon
  
  If (frmSelectColors.intButtonClicked = vbOK) Then
    For intI = 1 To intMaxXY
      For intJ = 1 To intMaxXY
        If (lngIconTemp(intI, intJ) = frmSelectColors.lngColor1) Then
          lngIcon(intI, intJ) = frmSelectColors.lngColor2
        End If
        If (lngIconTemp(intI, intJ) = frmSelectColors.lngColor2) Then
          lngIcon(intI, intJ) = frmSelectColors.lngColor1
        End If
      Next intJ
    Next intI
    
    Call DrawChangedIcon
  
    blnIconChanged = True
    Call CreateUndoIcon
  End If
End Sub

Private Sub ChangeColorIcon()
  frmSelectColors.lblStatus.Caption = "==>"
  frmSelectColors.Show (1)
  
  If (frmSelectColors.intButtonClicked = vbOK) Then
    For intI = 1 To intMaxXY
      For intJ = 1 To intMaxXY
        If (lngIcon(intI, intJ) = frmSelectColors.lngColor1) Then
          lngIcon(intI, intJ) = frmSelectColors.lngColor2
        End If
      Next intJ
    Next intI
    
    Call DrawChangedIcon
  
    blnIconChanged = True
    Call CreateUndoIcon
  End If
End Sub

Private Sub picIconArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim intI, intJ As Integer
  Dim intX1, intY1, intX2, intY2 As Integer
  Dim intTemp As Integer
  
  If (intToolNumber < 4) Then
    CreateUndoIcon
    Exit Sub
  End If
  
  If ((intPointX1 = 0) Or (intPointY1 = 0)) Then
    Exit Sub
  End If
  
  intPointX2 = X \ intScale + 1
  intPointY2 = Y \ intScale + 1
  
  If (intPointX2 < 1) Then
    intPointX2 = 1
  End If
  If (intPointX2 > intMaxXY) Then
    intPointX2 = intMaxXY
  End If
  If (intPointY2 < 1) Then
    intPointY2 = 1
  End If
  If (intPointY2 > intMaxXY) Then
    intPointY2 = intMaxXY
  End If
  
  If (intPointX1 > intPointX2) Then
    intX1 = intPointX2
    intX2 = intPointX1
  Else
    intX1 = intPointX1
    intX2 = intPointX2
  End If
  If (intPointY1 > intPointY2) Then
    intY1 = intPointY2
    intY2 = intPointY1
  Else
    intY1 = intPointY1
    intY2 = intPointY2
  End If
  picIconArea.FillStyle = 1
  picIconArea.Line ((intX1 - 1) * 12, (intY1 - 1) * 12)-(intX2 * 12, intY2 * 12), lngLineColor, B
  
  If ((intPointX1 = intPointX2) And (intPointY1 = intPointY2)) Then
    blnMakeSelection = False
    Exit Sub
  End If
  
  If (intToolNumber = 5) Then
    picIcon.DrawStyle = intDrawStyle - 1
    picIcon.Line (intPointX1 + 7, intPointY1 + 7)-(intPointX2 + 7, intPointY2 + 7), lngForeColor
    picIcon.DrawStyle = 0
    
    picIcon.PSet (intPointX1 + 7, intPointY1 + 7), lngForeColor
    picIcon.PSet (intPointX2 + 7, intPointY2 + 7), lngForeColor
  End If
    
  If ((intToolNumber = 6) Or (intToolNumber = 7)) Then
    If (Abs(intPointX2 - intPointX1) < Abs(intPointY2 - intPointY1)) Then
      If (intPointY2 > intPointY1) Then
        intPointY2 = intPointY1 + (Abs(intPointX1 - intPointX2))
      Else
        intPointY2 = intPointY1 - (Abs(intPointX1 - intPointX2))
      End If
    Else
      If (intPointX2 > intPointX1) Then
        intPointX2 = intPointX1 + (Abs(intPointY1 - intPointY2))
      Else
        intPointX2 = intPointX1 - (Abs(intPointY1 - intPointY2))
      End If
    End If
  
    If (intToolNumber = 7) Then
      picIcon.FillColor = lngBackColor
      picIcon.FillStyle = 0
      picIcon.Line (intPointX1 + 7, intPointY1 + 7)-(intPointX2 + 7, intPointY2 + 7), lngBackColor, B
    End If
    
    picIcon.DrawStyle = intDrawStyle - 1
    picIcon.FillStyle = 1
    picIcon.FillColor = lngForeColor
    picIcon.Line (intPointX1 + 7, intPointY1 + 7)-(intPointX2 + 7, intPointY2 + 7), lngForeColor, B
    picIcon.DrawStyle = 0
  End If
  
  If ((intToolNumber = 8) Or (intToolNumber = 9)) Then
    If (intToolNumber = 9) Then
      picIcon.FillColor = lngBackColor
      picIcon.FillStyle = 0
      picIcon.Line (intPointX1 + 7, intPointY1 + 7)-(intPointX2 + 7, intPointY2 + 7), lngBackColor, B
    End If
    
    picIcon.DrawStyle = intDrawStyle - 1
    picIcon.FillColor = lngForeColor
    picIcon.FillStyle = 1
    picIcon.Line (intPointX1 + 7, intPointY1 + 7)-(intPointX2 + 7, intPointY2 + 7), lngForeColor, B
    picIcon.DrawStyle = 0
  End If
  
  If ((intToolNumber = 10) Or (intToolNumber = 11)) Then
    Dim intCenterX, intCenterY As Integer
    Dim intRX, intRY, intR As Integer
    
    intCenterX = Int((intPointX1 + intPointX2) / 2)
    intCenterY = Int((intPointY1 + intPointY2) / 2)
    intRX = Abs(intPointX1 - intCenterX)
    intRY = Abs(intPointY1 - intCenterY)
    intR = intRX
    If (intR > intRY) Then
      intR = intRY
    End If
    
    If Int(intToolNumber = 10) Then
      picIcon.FillStyle = 1
      picIcon.FillColor = QBColor(7)
    Else
      picIcon.FillStyle = 0
      picIcon.FillColor = lngBackColor
    End If
    
    picIcon.DrawStyle = intDrawStyle - 1
    picIcon.Circle (intCenterX + 7, intCenterY + 7), intR, lngForeColor
    picIcon.DrawStyle = 0
  End If
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = picIcon.Point(intI + 7, intJ + 7)
    Next intJ
  Next intI
    
  Call DrawChangedIcon
  
  CreateUndoIcon
  blnIconChanged = True
  blnMakeSelection = False
End Sub

Private Sub InverseColorsIcon()
  Dim intI, intJ, intColor As Integer

  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      For intColor = 0 To 15
        If (QBColor(intColor) = lngIcon(intI, intJ)) Then
          lngIcon(intI, intJ) = QBColor(15 - intColor)
          intColor = 20
        End If
      Next intColor
    Next intJ
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
  Call CreateUndoIcon
End Sub

Private Sub UndoIcon()
  Dim intI, intJ As Integer
  
  intUndo = intUndo - 1
  If (intUndo = 0) Then intUndo = intMaxUndoNumber
  
  intUndoNumber = intUndoNumber - 1
  mnuEditUndo.Caption = "&Undo " & intUndoNumber - 1 & " time(s)"
  If (intUndoNumber < 2) Then
    tlbMain.Buttons(5).Enabled = False
    mnuEditUndo.Enabled = False
  End If
  
  For intI = 1 To intMaxXY
    For intJ = 1 To intMaxXY
      lngIcon(intI, intJ) = lngUndoIcon(intUndo, intI, intJ)
    Next intJ
  Next intI
  
  Call DrawChangedIcon
  
  blnIconChanged = True
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case Is = "tlbNew"
      NewIcon
    Case Is = "tlbSave"
      SaveIcon
    Case Is = "tlbUp"
      UpIcon
    Case Is = "tlbDown"
      DownIcon
    Case Is = "tlbLeft"
      LeftIcon
    Case Is = "tlbRight"
      RightIcon
    Case Is = "tlbRotateUCW"
      RotateUCWIcon
    Case Is = "tlbRotateCW"
      RotateCWIcon
    Case Is = "tlbOpen"
      mnuFileOpen_Click
    Case Is = "tlbShadow"
      ShadowIcon
    Case Is = "tlbChangeColor"
      ChangeColorIcon
    Case Is = "tlbChangeColors"
      ChangeColorsIcon
    Case Is = "tlbInverseColors"
      InverseColorsIcon
    Case Is = "tlbUndo"
      UndoIcon
  End Select
End Sub

Private Sub tlbProperties_ButtonClick(ByVal Button As MSComctlLib.Button)
  tlbProperties.Buttons(intDrawStyle).Value = tbrPressed
  tlbProperties.Buttons(intDrawStyle).Value = tbrUnpressed
  mnuViewLinePropertiesTool(intDrawStyle).Checked = False
  
  intDrawStyle = Button.Index
  
  tlbProperties.Buttons(intDrawStyle).Value = tbrPressed
  mnuViewLinePropertiesTool(intDrawStyle).Checked = True
End Sub

Private Sub tlbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim btnTemp As Button
  tlbTools.Buttons(intToolNumber).Value = tbrPressed
  tlbTools.Buttons(intToolNumber).Value = tbrUnpressed
  mnuViewDrawToolsTool(intToolNumber).Checked = False
    
  intToolNumber = Button.Index
  
  tlbTools.Buttons(intToolNumber).Value = tbrPressed
  mnuViewDrawToolsTool(intToolNumber).Checked = True
  
  For Each btnTemp In tlbProperties.Buttons
    btnTemp.Enabled = False
    mnuViewLinePropertiesTool(btnTemp.Index).Enabled = False
  Next btnTemp
  
  If (intToolNumber >= 5) Then
    For Each btnTemp In tlbProperties.Buttons
      btnTemp.Enabled = True
      mnuViewLinePropertiesTool(btnTemp.Index).Enabled = True
    Next btnTemp
  End If
End Sub
