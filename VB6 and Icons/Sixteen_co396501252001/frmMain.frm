VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Editor"
   ClientHeight    =   5895
   ClientLeft      =   855
   ClientTop       =   1890
   ClientWidth     =   8970
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   100
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   Visible         =   0   'False
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   2292
      Left            =   5640
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2172
      Begin VB.Shape shpMain 
         Height          =   375
         Left            =   240
         Shape           =   2  'Oval
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   30
         X2              =   150
         Y1              =   50
         Y2              =   50
      End
   End
   Begin VB.PictureBox picIcon1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   432
      Left            =   2400
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   360
      Width           =   432
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      ToolTipText     =   "Re-draws icon with latest color changes"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkSolidBox 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solid"
      Enabled         =   0   'False
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   4800
      Width           =   780
   End
   Begin VB.CheckBox chkSolidCirc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solid"
      Enabled         =   0   'False
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   780
   End
   Begin VB.VScrollBar vsbRed 
      Height          =   3870
      LargeChange     =   10
      Left            =   3360
      Max             =   0
      Min             =   255
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar vsbGreen 
      Height          =   3870
      LargeChange     =   10
      Left            =   4080
      Max             =   0
      Min             =   255
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar vsbBlue 
      Height          =   3840
      LargeChange     =   10
      Left            =   4800
      Max             =   0
      Min             =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "UNTITLED"
      Top             =   0
      Width           =   4572
   End
   Begin VB.Image imgRtArrow 
      Height          =   210
      Left            =   1200
      Picture         =   "frmMain.frx":030A
      ToolTipText     =   "Indicates transparent Color"
      Top             =   1440
      Width           =   300
   End
   Begin VB.Image imgLtArrow 
      Height          =   210
      Left            =   2760
      Picture         =   "frmMain.frx":0434
      ToolTipText     =   "Indicates transparent Color"
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblButtonLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Box"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   37
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label lblButtonLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ellipse"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   36
      Top             =   3720
      Width           =   570
   End
   Begin VB.Label lblButtonLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Line"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   35
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblButtonLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fill"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   34
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblButtonLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   33
      Top             =   840
      Width           =   450
   End
   Begin VB.Image imgButtonOn 
      Height          =   675
      Index           =   4
      Left            =   120
      Picture         =   "frmMain.frx":055E
      Top             =   4320
      Width           =   675
   End
   Begin VB.Image imgButtonOn 
      Height          =   675
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":0A24
      Top             =   3480
      Width           =   675
   End
   Begin VB.Image imgButtonOn 
      Height          =   675
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":0EEA
      Top             =   2520
      Width           =   675
   End
   Begin VB.Image imgButtonOn 
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":13B0
      Top             =   1560
      Width           =   675
   End
   Begin VB.Image imgButtonOn 
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":1876
      Top             =   600
      Width           =   675
   End
   Begin VB.Image imgButtonOff 
      Height          =   675
      Index           =   4
      Left            =   120
      Picture         =   "frmMain.frx":1D3C
      Top             =   4320
      Width           =   675
   End
   Begin VB.Image imgButtonOff 
      Height          =   675
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":2202
      Top             =   3480
      Width           =   675
   End
   Begin VB.Image imgButtonOff 
      Height          =   675
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":26C8
      Top             =   2520
      Width           =   675
   End
   Begin VB.Image imgButtonOff 
      Height          =   675
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":2B8E
      Top             =   1560
      Width           =   675
   End
   Begin VB.Image imgButtonOff 
      Height          =   675
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":3048
      Top             =   600
      Width           =   675
   End
   Begin VB.Shape shpOutlineBox 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   450
      Left            =   2205
      Top             =   1860
      Width           =   450
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   15
      Left            =   2280
      TabIndex        =   32
      Top             =   4320
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   1800
      TabIndex        =   31
      Top             =   4320
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   13
      Left            =   2280
      TabIndex        =   30
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   12
      Left            =   1800
      TabIndex        =   29
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   11
      Left            =   2280
      TabIndex        =   28
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   1800
      TabIndex        =   27
      Top             =   3360
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   2280
      TabIndex        =   26
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   1800
      TabIndex        =   25
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   2280
      TabIndex        =   24
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   1800
      TabIndex        =   23
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   2280
      TabIndex        =   22
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   20
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   19
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   960
      Width           =   315
   End
   Begin VB.Label lblPalNum 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   17
      Top             =   960
      Width           =   315
   End
   Begin VB.Label lblRedSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblGreenSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      ForeColor       =   &H0000BF00&
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblBlueSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblRedVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblGreenVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H0000BF00&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblBlueVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblColr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   6720
      TabIndex        =   9
      Top             =   5280
      Width           =   84
   End
   Begin VB.Label lblXY 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   552
   End
   Begin VB.Label lblPal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palette"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFileHeading 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "S&ave as..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelpHeading 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Icon Editor..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRefresh_Click()
Call DrawIcon
Call DrawGrid
End Sub

Private Sub Form_Click()
Unload frmAbout
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Line1.Visible = False
  shpMain.Visible = False
  LineFlag = False
  CircFlag = False
  BoxFlag = False
End If
End Sub

Private Sub Form_Load()
On Error GoTo Oops
'Adjust size to my original
Width = Int(8556 * Screen.TwipsPerPixelX / 12 + 0.5)
Height = Int(5076 * Screen.TwipsPerPixelY / 12 + 0.5)
ScaleWidth = Int(Width / Screen.TwipsPerPixelX - 6 + 0.5)
ScaleHeight = Int(Height / Screen.TwipsPerPixelY - 44 + 0.5)

'Adjust font in palette boxes
If Screen.TwipsPerPixelX < 15 Then
  For n% = 0 To 15
  lblPalNum(n%).FontSize = 8
  Next n%
End If

'Center form
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

chkSolidCirc.Left = 65
chkSolidCirc.Top = 279
chkSolidBox.Left = 65
chkSolidBox.Top = 337

cmdRefresh.Top = 40
cmdRefresh.Left = 245
cmdRefresh.Height = 25
cmdRefresh.Width = 97

For n% = 0 To 4
  lblButtonLabel(n%).Left = 70
  lblButtonLabel(n%).Top = 85 + 58 * n%
  imgButtonOn(n%).Left = 10
  imgButtonOn(n%).Top = 70 + 58 * n%
  imgButtonOff(n%).Left = 10
  imgButtonOff(n%).Top = 70 + 58 * n%
Next n%

lblName.Left = 20
lblName.Top = 13

lblPal.Left = 143
lblPal.Top = 349

For A% = 0 To 15
  X = 147 + (A% And 1) * 30
  Y = 120 + Int(A% / 2) * 28
  lblPalNum(A%).Left = X
  lblPalNum(A%).Top = Y
  lblPalNum(A%).Width = 21
  lblPalNum(A%).Height = 21
Next A%

lblRedSB.Left = 231
lblRedSB.Top = 344
lblRedSB.Height = 18
lblGreenSB.Left = 279
lblGreenSB.Top = 344
lblGreenSB.Height = 18
lblBlueSB.Left = 327
lblBlueSB.Top = 344
lblBlueSB.Height = 18

lblRedVal.Top = 358
lblRedVal.Left = 515
lblRedVal.Height = 20
lblGreenVal.Top = 358
lblGreenVal.Left = 565
lblGreenVal.Height = 20
lblBlueVal.Top = 358
lblBlueVal.Left = 615
lblBlueVal.Height = 20

lblXY.Left = 495
lblXY.Top = 40
lblXY.Width = 145

lblColr.Left = 420
lblColr.Top = 358
'lblColr.Width = 145:'(Autosized)

picIcon1.Left = 155
picIcon1.Top = 62
picIcon1.Width = 36
picIcon1.Height = 36
picIcon1.ScaleWidth = 32
picIcon1.ScaleHeight = 32

'shpOutlineBox.Left = 134
shpOutlineBox.Height = 29
shpOutlineBox.Width = 29

Text1.Left = 90
Text1.Top = 10
Text1.Height = 20
Text1.Width = 560

vsbRed.Left = 239
vsbRed.Top = 80
vsbRed.Height = 248
vsbGreen.Left = 287
vsbGreen.Top = 80
vsbGreen.Height = 248
vsbBlue.Left = 335
vsbBlue.Top = 80
vsbBlue.Height = 248

picMain.Left = 390
picMain.Top = 60

'*******  Main edit window  *******
'Values on next 2 lines can be changed by 32's to resize main edit window
picMain.Width = 293
picMain.Height = 293

picMain.ScaleWidth = picMain.Width - 4
picMain.ScaleHeight = picMain.Height - 4
'Size of pixel boxes:
SqWidth = Int((picMain.ScaleWidth - 33) / 32)
SqHeight = Int((picMain.ScaleHeight - 33) / 32)
Call NewFile
Visible = True

'1st 62 bytes of icon file:
Hdr = "0000010001002020100000000000E802"
Hdr = Hdr + "00001600000028000000200000004000"
Hdr = Hdr + "00000100040000000000800200000000"
Hdr = Hdr + "0000000000000000000000000000"

For n% = 0 To 7
BitPos(n%) = 2 ^ (7 - n%)
Next n%

FileChangedFlag = False
cmdRefresh.Enabled = False
Open "c:\Icn$path.txt" For Binary As #1
If e% = 0 Then
  Line Input #1, t$
End If
Close
If t$ > "" Then
  ChDrive Left$(t$, 2)
  ChDir t$
End If

PathText = Command()
If PathText > "" Then
  If Left$(PathText, 1) = Chr$(34) Then PathText = Mid$(PathText, 2)
  If Right$(PathText, 1) = Chr$(34) Then PathText = Left$(PathText, Len(PathText) - 1)
  Call GetFileTextInfo
  If TextStatus > 0 Then
    If Mid$(TempDirPath, 2, 1) = ":" Then
      ChDrive Left$(TempDirPath, 2)
      ChDir TempDirPath
    End If
    LoadPath = TempDirPath
    If TempFileName > "" Then
      LoadName = TempFileName
      FormOpenLoadedFlag = False
      Call CheckFileFormat(LoadName)
      If FileFormat = 1 Then
        Call LoadIconFile
        Else: frmMain.Text1.Text = "Can't Open " & LoadName
      End If
    End If
  End If
End If
Call ButtonSelect(0)
Exit Sub

Oops:
e% = Err
Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblXY.Caption = ""
lblColr.Visible = False
lblBlueVal.Visible = False
lblGreenVal.Visible = False
lblRedVal.Visible = False
MousePointer = 0
End Sub

Private Sub imgButtonOff_Click(Index As Integer)
LineStartFlag = False
Line1.Visible = False
If (CircStartFlag = True) Or (BoxStartFlag = True) Then
  Call DrawIcon
  Call DrawGrid
End If
CircStartFlag = False
BoxStartFlag = False
Call ButtonSelect(Index)
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X < 0) Or (X > picMain.ScaleWidth + 2) Or (Y < 0) Or (Y > picMain.Height - 2) Then
  MousePointer = 0
  Exit Sub
End If
MousePointer = 2
ClickedPixelX = Int(X / (picMain.ScaleWidth + 2) * 32)
ClickedPixelY = Int(Y / (picMain.ScaleWidth + 2) * 32)
If PenFlag = True Then
  picMain.Line (ClickedPixelX * (SqWidth + 1) + 1, ClickedPixelY * (SqHeight + 1) + 1)-((ClickedPixelX + 1) * (SqWidth + 1) - 1, (ClickedPixelY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  picIcon1.PSet (ClickedPixelX, ClickedPixelY), Pal(SelectedColor)
  If SelectedColor = TransparentColor Then
    frmMain.picMain.Line (Int((ClickedPixelX + 0.5) * (SqWidth + 1)), Int((ClickedPixelY + 0.5) * (SqHeight + 1)))-(Int((ClickedPixelX + 0.5) * (SqWidth + 1)) + 1, Int((ClickedPixelY + 0.5) * (SqHeight + 1)) + 1), 16777215, BF
    'frmMain.picIcon1.PSet (ClickedPixelX, ClickedPixelY), 12632256
  End If
  PixArray(ClickedPixelX, ClickedPixelY) = SelectedColor
  FileChangedFlag = True: cmdRefresh.Enabled = True
  Exit Sub
End If
If FillFlag = True Then
  Call Fill
  FileChangedFlag = True
  cmdRefresh.Enabled = True
  Call DrawIcon
  Call DrawGrid
  Exit Sub
End If
If LineFlag = True Then
    LineStartX = ClickedPixelX
    LineStartY = ClickedPixelY
    Line1.X1 = Int((LineStartX + 0.5) * (SqWidth + 1))
    Line1.Y1 = Int((LineStartY + 0.5) * (SqWidth + 1))
    Line1.X2 = Int((LineStartX + 0.5) * (SqWidth + 1))
    Line1.Y2 = Int((LineStartY + 0.5) * (SqWidth + 1))
    Line1.BorderColor = Pal(SelectedColor)
    Line1.Visible = True
    Exit Sub
End If
If CircFlag = True Then
  If CircStartFlag = False Then
    CircStartFlag = True
    CircStartX = ClickedPixelX
    CircStartY = ClickedPixelY
    shpMain.Shape = 2 'Oval
    'shpMain.Left = CircStartX
    'shpMain.Top = CircStartY
    shpMain.Width = 0
    shpMain.Height = 0
    shpMain.FillColor = Pal(SelectedColor)
    If chkSolidCirc.Value = 1 Then shpMain.FillStyle = 0
    If chkSolidCirc.Value = 0 Then shpMain.FillStyle = 1
    shpMain.Visible = True
    FileChangedFlag = True
    cmdRefresh.Enabled = True
    Exit Sub
  End If
End If
If BoxFlag = True Then
  shpMain.Shape = 0 'Rectangle
  BoxStartX = ClickedPixelX
  BoxStartY = ClickedPixelY
  shpMain.Left = Int(BoxStartX * (picMain.ScaleWidth + 2) / 32 + SqWidth / 2)
  shpMain.Top = Int(BoxStartY * (picMain.ScaleWidth + 2) / 32 + SqWidth / 2)
  shpMain.Width = 0
  shpMain.Height = 0
  shpMain.BorderColor = Pal(SelectedColor)
  shpMain.FillColor = Pal(SelectedColor)
  shpMain.FillStyle = 1 - chkSolidBox.Value
  'If chkSolidBox.Value = 1 Then shpMain.FillStyle = 0
  'If chkSolidBox.Value = 0 Then shpMain.FillStyle = 1
  shpMain.Visible = True
  Exit Sub
End If
End Sub

Private Sub picmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickedPixelX = Int(X / (picMain.ScaleWidth + 2) * 32)
ClickedPixelY = Int(Y / (picMain.ScaleWidth + 2) * 32)
If ClickedPixelX > 31 Then ClickedPixelX = 31
If ClickedPixelY > 31 Then ClickedPixelY = 31
If ClickedPixelX < 0 Then ClickedPixelX = 0
If ClickedPixelY < 0 Then ClickedPixelY = 0
MousePointer = 2
lblXY.Caption = "X, Y:  " & ClickedPixelX & ", " & ClickedPixelY
Tmp% = PixArray(ClickedPixelX, ClickedPixelY)
B% = Int(Pal(Tmp%) / 65536) And 255
g% = Int(Pal(Tmp%) / 256) And 255
r% = Pal(Tmp%) And 255
lblColr.Caption = "Color# " & Tmp%
lblBlueVal.Caption = B%
lblGreenVal.Caption = g%
lblRedVal.Caption = r%
lblColr.Visible = True
lblRedVal.Visible = True
lblGreenVal.Visible = True
lblBlueVal.Visible = True

If Button = 0 Then Exit Sub

If PenFlag = True Then
  If (X < 1) Or (X > picMain.ScaleWidth + 2) Or (Y < 1) Or (Y > picMain.ScaleHeight + 2) Then Exit Sub
  frmMain.AutoRedraw = False
  picMain.Line (ClickedPixelX * (SqWidth + 1) + 1, ClickedPixelY * (SqHeight + 1) + 1)-((ClickedPixelX + 1) * (SqWidth + 1) - 1, (ClickedPixelY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  picIcon1.PSet (ClickedPixelX, ClickedPixelY), Pal(SelectedColor)
  If SelectedColor = TransparentColor Then
    frmMain.picMain.Line (Int((ClickedPixelX + 0.5) * (SqWidth + 1)), Int((ClickedPixelY + 0.5) * (SqHeight + 1)))-(Int((ClickedPixelX + 0.5) * (SqWidth + 1)) + 1, Int((ClickedPixelY + 0.5) * (SqHeight + 1)) + 1), 16777215, BF
    'frmMain.picIcon1.PSet (ClickedPixelX, ClickedPixelY), 12632256
  End If
  frmMain.AutoRedraw = True
  PixArray(ClickedPixelX, ClickedPixelY) = SelectedColor
  FileChangedFlag = True: cmdRefresh.Enabled = True
End If
If LineFlag = True Then
  LineEndX = ClickedPixelX
  LineEndY = ClickedPixelY
  Line1.X2 = Int((LineEndX + 0.5) * (SqWidth + 1))
  Line1.Y2 = Int((LineEndY + 0.5) * (SqWidth + 1))
  Exit Sub
End If

If CircStartFlag = True Then
  CircEndX = ClickedPixelX
  CircEndY = ClickedPixelY
  If CircStartX < CircEndX Then
    shpMain.Left = (CircStartX * picMain.ScaleWidth / 32) + SqWidth / 2
    Else
    shpMain.Left = (CircEndX * picMain.ScaleWidth / 32) + SqWidth / 2
  End If
  If CircStartY < CircEndY Then
    shpMain.Top = (CircStartY * picMain.ScaleHeight / 32) + SqHeight / 2
    Else
    shpMain.Top = (CircEndY * picMain.ScaleHeight / 32) + SqHeight / 2
  End If
  shpMain.Width = Abs((CircEndX - CircStartX) * picMain.ScaleWidth / 32)
  shpMain.Height = Abs((CircEndY - CircStartY) * picMain.ScaleHeight / 32)
  shpMain.BorderColor = Pal(SelectedColor)
  shpMain.FillColor = Pal(SelectedColor)
End If
If BoxFlag = True Then
  BoxEndX = ClickedPixelX
  BoxEndY = ClickedPixelY
  If BoxStartX < BoxEndX Then
    shpMain.Left = (BoxStartX * picMain.ScaleWidth / 32) + SqWidth / 2
    Else
    shpMain.Left = (BoxEndX * picMain.ScaleWidth / 32) + SqWidth / 2
  End If
  If BoxEndY < BoxStartY Then
    shpMain.Top = (BoxEndY * picMain.ScaleHeight / 32) + SqHeight / 2
    Else
    shpMain.Top = (BoxStartY * picMain.ScaleHeight / 32) + SqHeight / 2
  End If
  shpMain.Width = Int(Abs(BoxEndX - BoxStartX) * picMain.ScaleWidth / 32 + 1)
  shpMain.Height = Int(Abs(BoxEndY - BoxStartY) * picMain.ScaleHeight / 32 + 1)
  Exit Sub
End If
Exit Sub

If PrevButton = 1 Then PrevButton = 0: Exit Sub
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ClickedPixelX = Int(X / (picMain.ScaleWidth + 2) * 32)
  ClickedPixelY = Int(Y / (picMain.ScaleWidth + 2) * 32)
  If ClickedPixelX > 31 Then ClickedPixelX = 31
  If ClickedPixelY > 31 Then ClickedPixelY = 31
  If ClickedPixelX < 0 Then ClickedPixelX = 0
  If ClickedPixelY < 0 Then ClickedPixelY = 0
If LineFlag = True Then
  LineEndX = ClickedPixelX
  LineEndY = ClickedPixelY
  Line1.X2 = Int((LineEndX + 0.5) * (SqWidth + 1))
  Line1.Y2 = Int((LineEndY + 0.5) * (SqWidth + 1))
  'draw line!
  TmpX% = Abs(LineStartX - LineEndX)
  TmpY% = Abs(LineStartY - LineEndY)
  If TmpX% > TmpY% Then
    For XPlot = LineEndX To LineStartX Step Sgn(LineStartX - LineEndX)
    YPlot = LineStartY + Int(Abs(XPlot - LineStartX) * TmpY% / TmpX% + 0.5) * Sgn(LineEndY - LineStartY)
    'frmMain.picMain.Line (XPlot * (SqWidth + 1) + 1, YPlot * (SqHeight + 1) + 1)-((XPlot + 1) * (SqWidth + 1) - 1, (YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    PixArray(XPlot, YPlot) = SelectedColor
    Next XPlot
    Else
    For YPlot = LineEndY To LineStartY Step Sgn(LineStartY - LineEndY)
    XPlot = LineStartX + Int(Abs(YPlot - LineStartY) * TmpX% / TmpY% + 0.5) * Sgn(LineEndX - LineStartX)
    'frmMain.picMain.Line (XPlot * (SqWidth + 1) + 1, YPlot * (SqHeight + 1) + 1)-((XPlot + 1) * (SqWidth + 1) - 1, (YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    PixArray(XPlot, YPlot) = SelectedColor
    Next YPlot
  End If
  Line1.Visible = False
  FileChangedFlag = True
  cmdRefresh.Enabled = True
  Call DrawIcon
  Call DrawGrid
  Exit Sub
End If
If CircStartFlag = True Then
  CircStartFlag = False
  shpMain.Visible = False
  'CircStartX, CircStartY, CircEndX, CircEndY
  If CircStartX > CircEndX Then
    Tmp% = CircStartY
    CircStartY = CircEndY
    CircEndY = Tmp%
  End If
  If CircStartY > CircEndY Then
    Tmp% = CircStartX
    CircStartX = CircEndX
    CircEndX = Tmp%
  End If
  Call DrawCircle(CircStartX, CircStartY, CircEndX, CircEndY)
  FileChangedFlag = True
  cmdRefresh.Enabled = True
  Call DrawIcon
  Call DrawGrid
  Exit Sub
End If
If BoxFlag = True Then
  BoxEndX = ClickedPixelX
  BoxEndY = ClickedPixelY
  If BoxStartX > BoxEndX Then
    Tmp% = BoxStartX
    BoxStartX = BoxEndX
    BoxEndX = Tmp%
  End If
  If BoxStartY > BoxEndY Then
    Tmp% = BoxStartY
    BoxStartY = BoxEndY
    BoxEndY = Tmp%
  End If
  'Draw box!
  If chkSolidBox.Value = 0 Then
    For A% = BoxStartX To BoxEndX
    PixArray(A%, BoxStartY) = SelectedColor
    PixArray(A%, BoxEndY) = SelectedColor
    'frmMain.picMain.Line (A% * (SqWidth + 1) + 1, BoxStartY * (SqHeight + 1) + 1)-((A% + 1) * (SqWidth + 1) - 1, (BoxStartY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    'frmMain.picMain.Line (A% * (SqWidth + 1) + 1, BoxEndY * (SqHeight + 1) + 1)-((A% + 1) * (SqWidth + 1) - 1, (BoxEndY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    Next A%
    For A% = BoxStartY To BoxEndY
    PixArray(BoxStartX, A%) = SelectedColor
    PixArray(BoxEndX, A%) = SelectedColor
    'frmMain.picMain.Line (BoxStartX * (SqWidth + 1) + 1, A% * (SqHeight + 1) + 1)-((BoxStartX + 1) * (SqWidth + 1) - 1, (A% + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    'frmMain.picMain.Line (BoxEndX * (SqWidth + 1) + 1, A% * (SqHeight + 1) + 1)-((BoxEndX + 1) * (SqWidth + 1) - 1, (A% + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    Next A%
    Else
    For A% = BoxStartX To BoxEndX
    For B% = BoxStartY To BoxEndY
    PixArray(A%, B%) = SelectedColor
    'frmMain.picMain.Line (A% * (SqWidth + 1) + 1, BoxStartY * (SqHeight + 1) + 1)-((A% + 1) * (SqWidth + 1) - 1, (BoxStartY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    'frmMain.picMain.Line (A% * (SqWidth + 1) + 1, BoxEndY * (SqHeight + 1) + 1)-((A% + 1) * (SqWidth + 1) - 1, (BoxEndY + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
    Next B%, A%
  End If
  shpMain.Visible = False
  FileChangedFlag = True
  cmdRefresh.Enabled = True
  Call DrawIcon
  Call DrawGrid
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ExitAfterSaveFlag = True Then Exit Sub
If FileChangedFlag = True Then
  A% = MsgBox("Save Current Icon File?", 51, "File NOT Saved")
  If A% = 6 Then
    Rem Yes, save file
    ExitAfterSaveFlag = True
    Cancel = 1
    Load frmSaveAs
    Exit Sub
  End If
  If A% = 7 Then Cancel = 0: Exit Sub 'No
  If A% = 2 Then Cancel = 1: Exit Sub 'Cancel
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Oops
Unload frmOpenFile
Unload frmSaveAs
Unload frmAbout
Open "c:\Icn$path.txt" For Binary As #1
If e% = 0 Then
  t$ = CurDir$ + Chr$(13) + Chr$(10)
  Put #1, 1, t$
End If
Close
Exit Sub

Oops:
e% = Err
Resume Next
End Sub

Private Sub lblPalNum_Click(Index As Integer)
SelectedColor = Index
Call OutlineBox
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileNew_Click()
'New icon file
If FileChangedFlag = True Then
  A% = MsgBox("Save Current Icon File?", 51, "File NOT Saved")
  If A% = 6 Then
    'Yes
    NewAfterSaveFlag = True
    Call SaveIconFile
    Exit Sub
  End If
  'If a% = 7 Then Stop 'No
  If A% = 2 Then Exit Sub 'Cancel
End If
Call NewFile
End Sub

Private Sub mnuFileOpen_Click()
'Open file
Load frmOpenFile
End Sub

Private Sub mnuFileSave_Click()
'Save
If (Text1.Text = "UNTITLED.ico") Or (LCase$(Right$(LoadName, 4)) = ".bmp") Then
  Load frmSaveAs
  Exit Sub
End If
If FileChangedFlag = True Then
  Call SaveIconFile
End If
End Sub

Private Sub mnuFileSaveAs_Click()
'Save As...
Load frmSaveAs
End Sub

Private Sub mnuHelpAbout_Click()
Load frmAbout
End Sub

Private Sub vsbBlue_Change()
Call vsbBlue_Scroll
End Sub

Private Sub vsbBlue_Scroll()
If SkipFlag = False Then FileChangedFlag = True: cmdRefresh.Enabled = True
B% = vsbBlue.Value
lblBlueSB.Caption = B%
Pal(SelectedColor) = (Pal(SelectedColor) And 65535) Or (B% * 65536)
Call SetPalColors
End Sub

Private Sub vsbGreen_Change()
Call vsbGreen_Scroll
End Sub

Private Sub vsbGreen_Scroll()
If SkipFlag = False Then FileChangedFlag = True: cmdRefresh.Enabled = True
g% = vsbGreen.Value
lblGreenSB.Caption = g%
Pal(SelectedColor) = (Pal(SelectedColor) And 16711935) Or (g% * 256&)
Call SetPalColors
End Sub

Private Sub vsbRed_Change()
Call vsbRed_Scroll
End Sub

Private Sub vsbRed_Scroll()
If SkipFlag = False Then FileChangedFlag = True: cmdRefresh.Enabled = True
r% = vsbRed.Value
lblRedSB = r%
Pal(SelectedColor) = (Pal(SelectedColor) And 16776960) Or r%
Call SetPalColors
End Sub
