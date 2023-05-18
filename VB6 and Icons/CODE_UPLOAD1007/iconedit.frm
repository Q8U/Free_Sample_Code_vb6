VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Editor 
   Caption         =   "IconWorks Editor: 1 - [Untitled]"
   ClientHeight    =   4665
   ClientLeft      =   1290
   ClientTop       =   1830
   ClientWidth     =   6165
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
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ICONEDIT.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   Begin VB.PictureBox Pic_ToolPalette 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      Height          =   3840
      Left            =   0
      Picture         =   "ICONEDIT.frx":030A
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.VScrollBar Scrl_Zoom 
      Height          =   3870
      HelpContextID   =   1903
      LargeChange     =   32
      Left            =   480
      Max             =   258
      Min             =   34
      SmallChange     =   32
      TabIndex        =   4
      Tag             =   "9040"
      Top             =   -15
      Value           =   34
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Pic_Edit 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   735
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Tag             =   "9000"
      Top             =   -30
      Visible         =   0   'False
      Width           =   3870
      Begin VB.PictureBox Pic_Mask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1965
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_Image 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1380
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_Grid 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1380
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         Top             =   1515
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_Work 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1965
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1515
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_TempImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1380
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   2085
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_TempMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1965
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         Top             =   2085
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Pic_EditTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   1680
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   20
         Top             =   2745
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Pic_StatusArea 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   4680
      HelpContextID   =   1904
      Left            =   4590
      ScaleHeight     =   308
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "9010"
      Top             =   -30
      Visible         =   0   'False
      Width           =   1590
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":138C
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "ICONEDIT.frx":1696
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   180
         Width           =   480
      End
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":19A0
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   885
         Picture         =   "ICONEDIT.frx":1CAA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   180
         Width           =   480
      End
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":1FB4
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   210
         Picture         =   "ICONEDIT.frx":22BE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   840
         Width           =   480
      End
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":25C8
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   885
         Picture         =   "ICONEDIT.frx":28D2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   840
         Width           =   480
      End
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":2BDC
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   210
         Picture         =   "ICONEDIT.frx":2EE6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   1500
         Width           =   480
      End
      Begin VB.PictureBox Pic_Icons 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":31F0
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   885
         Picture         =   "ICONEDIT.frx":34FA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "9130"
         Top             =   1500
         Width           =   480
      End
      Begin VB.PictureBox Pic_Undo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "ICONEDIT.frx":3804
         ForeColor       =   &H80000008&
         Height          =   480
         HelpContextID   =   1902
         Left            =   1035
         Picture         =   "ICONEDIT.frx":3B0E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "9150"
         Top             =   2505
         Width           =   480
      End
      Begin VB.OptionButton Opt_Mouse 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Index           =   0
         Left            =   690
         TabIndex        =   17
         Tag             =   "9160"
         Top             =   3600
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Opt_Mouse 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         HelpContextID   =   1901
         Index           =   1
         Left            =   690
         TabIndex        =   1
         Tag             =   "9160"
         Top             =   4200
         Width           =   195
      End
   End
   Begin VB.PictureBox Pic_ColorPalette 
      Height          =   810
      HelpContextID   =   1905
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "9030"
      Top             =   3840
      Visible         =   0   'False
      Width           =   4605
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   15
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      HelpContextID   =   1100
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&New"
         HelpContextID   =   1101
         Index           =   0
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Open..."
         HelpContextID   =   1102
         Index           =   1
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Save"
         HelpContextID   =   1103
         Index           =   2
         Shortcut        =   +{F12}
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "Save &As..."
         HelpContextID   =   1104
         Index           =   3
         Shortcut        =   {F12}
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "E&xit"
         HelpContextID   =   1105
         Index           =   5
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
      HelpContextID   =   1200
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Undo"
         HelpContextID   =   1201
         Index           =   0
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Cu&t"
         HelpContextID   =   1202
         Index           =   2
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Copy"
         HelpContextID   =   1203
         Index           =   3
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Paste"
         HelpContextID   =   1204
         Index           =   4
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Paste &Opaque"
         HelpContextID   =   1205
         Index           =   5
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Delete"
         HelpContextID   =   1206
         Index           =   6
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Select All"
         HelpContextID   =   1207
         Index           =   7
         Shortcut        =   ^A
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Flip &Horizontal"
         HelpContextID   =   1208
         Index           =   9
         Shortcut        =   ^H
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Flip &Vertical"
         HelpContextID   =   1209
         Index           =   10
         Shortcut        =   ^V
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Rotate 90  &Right"
         HelpContextID   =   1210
         Index           =   12
         Shortcut        =   ^R
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "Rotate 90  &Left"
         HelpContextID   =   1211
         Index           =   13
         Shortcut        =   ^L
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu Menu_EditSelection 
         Caption         =   "&Invert"
         HelpContextID   =   1212
         Index           =   15
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu Menu_View 
      Caption         =   "&View"
      HelpContextID   =   1300
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Status Area"
         Checked         =   -1  'True
         HelpContextID   =   1301
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Tool Palette"
         Checked         =   -1  'True
         HelpContextID   =   1302
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Color Palette"
         Checked         =   -1  'True
         HelpContextID   =   1303
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Zooming Scroll bar"
         Checked         =   -1  'True
         HelpContextID   =   1304
         Index           =   3
         Shortcut        =   ^Z
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "Show &All"
         HelpContextID   =   1305
         Index           =   4
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Grid"
         HelpContextID   =   1306
         Index           =   6
         Shortcut        =   ^G
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Line Grid"
         HelpContextID   =   1307
         Index           =   7
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Dotted Grid"
         HelpContextID   =   1308
         Index           =   8
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "Zoom &In"
         HelpContextID   =   1309
         Index           =   10
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "Zoom &Out"
         HelpContextID   =   1310
         Index           =   11
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Focus to Editor on Drag&&Drop"
         Checked         =   -1  'True
         HelpContextID   =   1311
         Index           =   13
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "Display &Border on selected Icon"
         Checked         =   -1  'True
         HelpContextID   =   1312
         Index           =   14
      End
      Begin VB.Menu Menu_ViewSelection 
         Caption         =   "&Pop-up Tools menu on Form Click"
         Checked         =   -1  'True
         HelpContextID   =   1313
         Index           =   15
      End
   End
   Begin VB.Menu Menu_Tools 
      Caption         =   "&Tools"
      HelpContextID   =   1400
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Select"
         HelpContextID   =   1401
         Index           =   0
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Paint"
         HelpContextID   =   1402
         Index           =   1
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Fill"
         HelpContextID   =   1403
         Index           =   2
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Line"
         HelpContextID   =   1404
         Index           =   3
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Square"
         HelpContextID   =   1405
         Index           =   4
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "Filled S&quare"
         HelpContextID   =   1406
         Index           =   5
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "&Circle"
         HelpContextID   =   1407
         Index           =   6
      End
      Begin VB.Menu Menu_ToolsSelection 
         Caption         =   "Filled C&ircle"
         HelpContextID   =   1408
         Index           =   7
      End
   End
   Begin VB.Menu Menu_Icons 
      Caption         =   "&Icons"
      HelpContextID   =   1500
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&1 - [Untitled]"
         Index           =   0
      End
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&2 - [Untitled]"
         Index           =   1
      End
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&3 - [Untitled]"
         Index           =   2
      End
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&4 - [Untitled]"
         Index           =   3
      End
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&5 - [Untitled]"
         Index           =   4
      End
      Begin VB.Menu Menu_IconsSelection 
         Caption         =   "&6 - [Untitled]"
         Index           =   5
      End
   End
   Begin VB.Menu Menu_Color 
      Caption         =   "&Color"
      HelpContextID   =   1600
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "&Colors..."
         HelpContextID   =   1601
         Index           =   0
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "&Default Color Palette"
         Checked         =   -1  'True
         HelpContextID   =   1602
         Index           =   1
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "S&olid Colors Only"
         HelpContextID   =   1603
         Index           =   2
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "I&NI File Color Palette"
         HelpContextID   =   1604
         Index           =   3
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "&Save Color Palette to INI file"
         HelpContextID   =   1605
         Index           =   5
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "&Left/Right Colors"
         Checked         =   -1  'True
         HelpContextID   =   1606
         Index           =   7
      End
      Begin VB.Menu Menu_ColorSelection 
         Caption         =   "Screen/&Inverse Colors"
         HelpContextID   =   1607
         Index           =   8
      End
   End
   Begin VB.Menu Menu_Brush 
      Caption         =   "&Brush"
      HelpContextID   =   1700
      Begin VB.Menu Menu_BrushSelection 
         Caption         =   "&Small (1 x 1)"
         Checked         =   -1  'True
         HelpContextID   =   1701
         Index           =   1
      End
      Begin VB.Menu Menu_BrushSelection 
         Caption         =   "&Medium (3 x 3)"
         HelpContextID   =   1702
         Index           =   3
      End
      Begin VB.Menu Menu_BrushSelection 
         Caption         =   "&Large (5 x 5)"
         HelpContextID   =   1703
         Index           =   5
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Index"
         Index           =   1
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Keyboard"
         Index           =   2
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Commands"
         Index           =   3
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Using Help"
         Index           =   4
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&About..."
         Index           =   6
      End
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-G, I-Z
#If Win32 Then
    DefLng H
#Else
    DefInt H
#End If

' When the View.Grid option is selected or deselected, this routine
' performs the neccessary checking, unchecking of the Grid menu item
' and Enables or disables the Grid type menu items appropriately,
' and creates the Grid if selected.
Private Sub Activate_Grid()
    Menu_ViewSelection(MID_GRID).Checked = Not Menu_ViewSelection(MID_GRID).Checked
    Menu_ViewSelection(MID_LINE_GRID).Enabled = Menu_ViewSelection(MID_GRID).Checked
    Menu_ViewSelection(MID_DOTTED_GRID).Enabled = Menu_ViewSelection(MID_GRID).Checked
    If Menu_ViewSelection(MID_GRID).Checked Then Create_Grid
    Magnify_Icon 0, 0, 31, 31
End Sub

Private Sub Adjust_Color_Palette()
    ' Determine if Editor is large enough for ColorPalette
    If (ScaleHeight < (Pic_ColorPalette.Height + 34)) And Menu_ViewSelection(MID_COLOR_PALETTE).Checked Then
        Menu_ViewSelection(MID_COLOR_PALETTE).Checked = False
        Menu_ViewSelection(MID_COLOR_PALETTE).Enabled = False
    Else
        Menu_ViewSelection(MID_COLOR_PALETTE).Enabled = True
        ' Check if the ColorPalette is selected to be displayed.
        If Menu_ViewSelection(MID_COLOR_PALETTE).Checked Then
            NewTop = ScaleHeight - Pic_ColorPalette.Height
            NewWidth = ScaleWidth
            ' Check if the StatusArea is currently displayed, since the StatusArea
            ' affects the width of the ColorPalette.
            If Menu_ViewSelection(MID_STATUS_BAR).Checked Then NewWidth = NewWidth - Pic_StatusArea.Width + 1
            ' Reposition and Resize the ColorPalette.
            Pic_ColorPalette.Move Pic_ColorPalette.Left, NewTop, NewWidth
        End If
    End If
End Sub

' When the Editor is Resized, the the Zoom Scrollbar must be resized,
' repositioned.  This routine performs the neccessary repositioning
' for the Zoom Scrollbar.
'
Private Sub Adjust_ScrollBar()
    If Menu_ViewSelection(MID_TOOL_PALETTE).Checked Then NewLeft = Pic_ToolPalette.Width Else NewLeft = -1
    NewHeight = ScaleHeight + 2 + Menu_ViewSelection(MID_COLOR_PALETTE).Checked * Pic_ColorPalette.Height
    NewMax = ScaleWidth - (NewLeft - Menu_ViewSelection(MID_ZOOM_SCROLLBAR).Checked * Scrl_Zoom.Width) + 3 + (Menu_ViewSelection(MID_STATUS_BAR).Checked * Pic_StatusArea.Width)
    ' Set the new Max value to the smaller of the NewHeight and the NewMax value.
    If NewMax > NewHeight Then NewMax = NewHeight
    ' We must adjust the New Max value obtained above so it is a multiple of 32.
    While (NewMax - 2) Mod 32
        NewMax = NewMax - 1
    Wend
    Scrl_Zoom.Move NewLeft, Scrl_Zoom.Top, Scrl_Zoom.Width, NewHeight
    Scrl_Zoom.Max = NewMax
    ' Determine if current size of the Editing area is larger than the new maximum size.
    If Pic_Edit.Height > Scrl_Zoom.Max Then
        Scrl_Zoom.Value = Scrl_Zoom.Max
    Else
        Scrl_Zoom_Change
    End If
End Sub

' When the Editor is Resized, the the StatusArea must be resized, and hidden if necessary
Private Sub Adjust_Status_Bar()
    ' Determine if Editor is large enough for StatusArea
    If ((ScaleHeight < (Pic_Icons(5).Top + Pic_Icons(5).Height + HIGHLIGHT)) Or (ScaleWidth < 2 * Pic_StatusArea.Width)) And Menu_ViewSelection(MID_STATUS_BAR).Checked Then
        Menu_ViewSelection(MID_STATUS_BAR).Checked = False
        Menu_ViewSelection(MID_STATUS_BAR).Enabled = False
    Else
        Menu_ViewSelection(MID_STATUS_BAR).Enabled = True
        If Menu_ViewSelection(MID_STATUS_BAR).Checked Then
            OldHeight = Pic_StatusArea.Height
            Pic_StatusArea.Move ScaleWidth - Pic_StatusArea.Width + 1, Pic_StatusArea.Top, Pic_StatusArea.Width, ScaleHeight + 2
            NewTop = Pic_StatusArea.ScaleHeight - 4.25 * Opt_Mouse(0).Height
            Opt_Mouse(0).Top = NewTop + 1.25 * Opt_Mouse(0).Height
            Opt_Mouse(1).Top = NewTop + 3 * Opt_Mouse(0).Height
            Visibility = NewTop >= (Pic_Icons(5).Top + Pic_Icons(5).Height + 2 * HIGHLIGHT + 4)
            Opt_Mouse(0).Visible = Visibility
            Opt_Mouse(1).Visible = Visibility
            ' Determine if there is room to display the Undo Icon
            If Opt_Mouse(0).Visible Then
                Pic_Undo.Visible = NewTop >= (Pic_Undo.Top + Pic_Undo.Height)
            Else
                Pic_Undo.Visible = (Pic_Undo.Top + Pic_Undo.Height) <= Pic_StatusArea.ScaleHeight
            End If
            ' We only want to force it to repaint if its size has changed
            If OldHeight <> Pic_StatusArea.Height Then Pic_StatusArea.Refresh
        End If
    End If

End Sub

' When the Editor is Resized, the the ToolPalette may need to be hidden.
Private Sub Adjust_Tool_Palette()
    If ((ScaleHeight < (Pic_ToolPalette.Height + Pic_ColorPalette.Height)) And Menu_ViewSelection(MID_COLOR_PALETTE).Checked) Or ((ScaleHeight < Pic_ToolPalette.Height) And (Not Menu_ViewSelection(MID_COLOR_PALETTE).Checked)) Then
        Menu_ViewSelection(MID_TOOL_PALETTE).Checked = False
        Menu_ViewSelection(MID_TOOL_PALETTE).Enabled = False
    Else
        Menu_ViewSelection(MID_TOOL_PALETTE).Enabled = True
    End If
End Sub

' To make magnified pixels appear more like single pixels when working
' with the various tools, this routine is used to adjust the actual pixel
' clicked to the Center or to the Edge of the selected magnified pixel.
Private Sub Adjust_X_and_Y(X As Single, Y As Single, NewX, NewY, Adjustment)
    Select Case Adjustment
        Case TO_CENTER
            NewX = (X \ PixelSize) * PixelSize + (PixelSize \ 2)
            NewY = (Y \ PixelSize) * PixelSize + (PixelSize \ 2)
        Case TO_EDGE
            Xpix = X \ PixelSize
            Ypix = Y \ PixelSize
            If (X Mod PixelSize) >= (PixelSize / 2) Then Xpix = Xpix + 1
            If (Y Mod PixelSize) >= (PixelSize / 2) Then Ypix = Ypix + 1
            NewX = Xpix * PixelSize
            NewY = Ypix * PixelSize
    End Select
End Sub

' This routine is called when ever the selection is to be comitted to the Icon.
Private Sub Commit_Selection(ResetFlags)
    ' Erase the Rectangle highlighting the selection
    If Selecting Then Draw_Selection_Rectangle
    Scale_Region True, X1, Y1, X2, Y2, True
    If Menu_ViewSelection(MID_GRID).Checked Then Display_Grid (Pic_Edit.hDC), X1, Y1, X2, Y2
    w = X2 - X1
    h = Y2 - Y1
    If Opaque Then
        ' An exact copy of Mask and Image of the Selection are copied
        R = BitBlt(Pic_Image.hDC, X1, Y1, w, h, Pic_TempImage.hDC, X1SelectFrom, Y1SelectFrom, SRCCOPY)
        R = BitBlt(Pic_Mask.hDC, X1, Y1, w, h, Pic_TempMask.hDC, X1SelectFrom, Y1SelectFrom, SRCCOPY)
    Else
        ' Only the Foreground colors are copied.
        R = BitBlt(Pic_Image.hDC, X1, Y1, w, h, Pic_TempMask.hDC, X1SelectFrom, Y1SelectFrom, SRCAND)
        R = BitBlt(Pic_Image.hDC, X1, Y1, w, h, Pic_TempImage.hDC, X1SelectFrom, Y1SelectFrom, SRCINVERT)
        R = BitBlt(Pic_Mask.hDC, X1, Y1, w, h, Pic_TempMask.hDC, X1SelectFrom, Y1SelectFrom, SRCAND)
    End If
    If ResetFlags Then
        Selecting = False
        MovingSelection = False
    End If
    ' Replace the Icons current Mask and Image bitmaps with the new ones.
    Update_Icon Pic_Icons(CurrentIcon)
End Sub

' This routine is called either in response to an Edit.Cut or Edit.Copy
Private Sub Copy_Selection_To_ClipBoard(X1, Y1, X2, Y2)
    Scale_Region True, X1, Y1, X2, Y2, True
    Pic_Work.Picture = LoadPicture()
    Pic_Work.Move Pic_Work.Left, Pic_Work.Top, X2 - X1, Y2 - Y1
    R = BitBlt(Pic_Work.hDC, 0, 0, Pic_Work.Width, Pic_Work.Height, Pic_Icons(CurrentIcon).hDC, X1, Y1, SRCCOPY)
    Clipboard.Clear
    Clipboard.SetData Pic_Work.Image
End Sub

' The Grid is drawn on a separate Picture control with a White BackGround and
' Black Foreground, and is only drawn once.  When a piece of the
' grid is to be redisplayed, the part of the Grid that corresponds
' to the area of the Icon that needs to be updated is BitBlt'd
' from the Picture of the Grid over the top of the Displayed Icon,
' Using the SCRAND raster operation.
Private Sub Create_Grid()
    Pic_Grid.Move 0, 0, Scrl_Zoom.Value, Scrl_Zoom.Value
    Pic_Grid.Cls
    If Menu_ViewSelection(MID_LINE_GRID).Checked Then
        For i = 1 To 31
            Pic_Grid.Line (0, i * PixelSize)-(Pic_Grid.ScaleWidth, i * PixelSize)
            Pic_Grid.Line (i * PixelSize, 0)-(i * PixelSize, Pic_Grid.ScaleHeight)
        Next i
    Else
        For Y = 0 To 31
            For X = 0 To 31
                Pic_Grid.PSet ((X * PixelSize) + (PixelSize \ 2), (Y * PixelSize) + (PixelSize \ 2))
            Next X
        Next Y
    End If
End Sub

' This routine is called in response to an Edit.Cut, Copy, or
' Delete menu selection.
Private Sub Cut_Copy_Or_Delete_Selection(Index)
    If Index = MID_DELETE Then
        Scale_Region True, X1, Y1, X2, Y2, True
    Else
        Copy_Selection_To_ClipBoard X1, Y1, X2, Y2
    End If
    If Index <> MID_COPY Then
        Delete_Selection_From_Icon
        Selecting = False
        Magnify_Icon X1, Y1, X2, Y2
    End If
End Sub

' This routine is called either in response to an Edit.Cut or Edit.Delete
Private Sub Delete_Selection_From_Icon()
    Scale_Region True, X1, Y1, X2, Y2, True
    Pic_Image.Line (X1, Y1)-(X2 - 1, Y2 - 1), BLACK, BF
    Pic_Mask.Line (X1, Y1)-(X2 - 1, Y2 - 1), WHITE, BF
    Update_Icon Pic_Icons(CurrentIcon)
End Sub

' This routine is called to display the mouse coordinates.  The
' coordinates are that of the magnified pixels, not the screen
' coordinates, thus, regardless of the magnification of the Icon,
' the coordinates will all range from 0-31.
Private Sub Display_Mouse_Coordinates(Xpix, Ypix)
    If (Xpix >= 0) And (Xpix <= 31) Then
        Pic_StatusArea.CurrentX = Pic_StatusArea.ScaleWidth * 0.25
        Pic_StatusArea.Print Space$(Abs(Xpix < 10)) + Format$(Xpix);
    End If
    If (Ypix >= 0) And (Ypix <= 31) Then
        Pic_StatusArea.CurrentX = Pic_StatusArea.ScaleWidth * 0.75
        Pic_StatusArea.Print Space$(Abs(Ypix < 10)) + Format$(Ypix);
    End If
End Sub

' The routine is called to display
' the Line or Box for the above tools when requested.
Private Sub Draw_Line_Or_Box()
    If CurrentTool = TID_LINE Then
        Pic_Edit.Line (X1Region, Y1Region)-(X2Region, Y2Region)
    Else
        Pic_Edit.Line (X1Region, Y1Region)-(X2Region, Y2Region), , B
    End If
End Sub

' Since Windows INItialization files only allow the writing string information,
' the Color Palette color values are stored in 3 strings, 1 for each row of the
' color palette.  This routine is called to extract these color values from
' the strings and assign them to the Long Integer Color array.
Private Sub Extract_Colors_From_String(ColorString As String, FirstColor, RefreshPalette)
    N = 1
    For C = FirstColor To FirstColor + 15
        i = InStr(N, ColorString, " ")
        If i = 0 Then i = Len(ColorString) + 1
        Colors(C) = Val(Mid$(ColorString, N, i - N))
        N = i + 1
    Next C
    If RefreshPalette Then
        Display_Color_Palette Pic_ColorPalette
        If ColorPaletteLoaded Then Display_Color_Palette ColorPalette.Pic_ColorPalette
    End If
End Sub

' This routine performs a Surface Flood fill of the selected area
' within the Icon when the Fill tool is in use.  The Mask color
' of all pixels that are different is then compare to the Mask
' color of the original starting Pixel.  If the Mask colors are the
' same the Pixel should be changed, otherwise it is unchanged.
Private Sub Fill_Tool_MouseUp(X As Single, Y As Single, X2, Y2)
Dim MaskColor As Long, NewPixelColor As Long
    If (Abs(X) >= 0) And (Abs(X) <= Pic_Edit.ScaleWidth) And (Abs(Y) >= 0) And (Abs(Y) <= Pic_Edit.ScaleHeight) Then
        Screen.MousePointer = HOURGLASS
        Xpixel = X \ PixelSize
        Ypixel = Y \ PixelSize
        MaskColor = Pic_Mask.Point(Xpixel, Ypixel)
        Pic_Work.FillColor = Pic_Edit.ForeColor
        Pic_Work.Picture = Pic_Icons(CurrentIcon).Image
        R = ExtFloodFill(Pic_Work.hDC, Xpixel, Ypixel, Pic_Icons(CurrentIcon).Point(Xpixel, Ypixel), FLOODFILLSURFACE)
        For Ypix = 0 To 31
            For Xpix = 0 To 31
                NewPixelColor = Pic_Work.Point(Xpix, Ypix)
                If (Pic_Icons(CurrentIcon).Point(Xpix, Ypix) <> NewPixelColor) And (Pic_Mask.Point(Xpix, Ypix) = MaskColor) Then
                    '
                    ' Pixel was changed by the FloodFill and its Mask matched that of the
                    ' starting Pixel, so we make the change on the Icons Image and Mask
                    '
                    If Opt_Mouse(SCREEN_COLORS).Value Then NewPixelColor = Pic_Image.ForeColor
                    Pic_Image.PSet (Xpix, Ypix), NewPixelColor
                    Pic_Mask.PSet (Xpix, Ypix)
                End If
            Next Xpix
        Next Ypix
        Screen.MousePointer = DEFAULT
        X2 = 31
        Y2 = 31
    End If
End Sub

' This routine Flips (Mirrors) the current Selection either
' Horizontally or Vertically.   The Windows API StretchBlt()
' routine is used to perform the Flip operation.
Private Sub Flip_Selection(Index)
    w = X2SelectFrom - X1SelectFrom
    h = Y2SelectFrom - Y1SelectFrom
    If Index = MID_FLIP_HORIZONTAL Then
        DestX = X2SelectFrom - 1
        DestY = Y1SelectFrom
        DestWidth = -w
        DestHeight = h
    Else
        DestX = X1SelectFrom
        DestY = Y2SelectFrom - 1
        DestWidth = w
        DestHeight = -h
    End If
    R = StretchBlt(Pic_TempImage.hDC, DestX, DestY, DestWidth, DestHeight, Pic_TempImage.hDC, X1SelectFrom, Y1SelectFrom, w, h, SRCCOPY)
    R = StretchBlt(Pic_TempMask.hDC, DestX, DestY, DestWidth, DestHeight, Pic_TempMask.hDC, X1SelectFrom, Y1SelectFrom, w, h, SRCCOPY)
    R = StretchBlt(Pic_Work.hDC, DestX, DestY, DestWidth, DestHeight, Pic_Work.hDC, X1SelectFrom, Y1SelectFrom, w, h, SRCCOPY)
    Scale_Region True, X1, Y1, X2, Y2, False
    If Not MovingSelection Then Commit_Selection False
    Magnify_Icon X1, Y1, X2, Y2
End Sub



Private Sub Form_Load()

    ' Determine color mode and set Icons Image bitmap size accordingly
    ' This value, ImageSize, is used within the routines Update_Icon and
    ' Extract_Image_And_Mask within ICONWRKS.BAS
    '
    If GetDeviceCaps(hDC, BITSPIXEL) = 8 Then
        ImageSize = 1024 'bytes
    Else
        ImageSize = 512  'bytes
    End If

    ' Depending on the Video mode and video driver being used for Windows,
    ' the controls on a form might not be the same dimensions when loaded
    ' under differenct resolutions.  Since the Editor is based on a
    ' 32x32 pixel Icon, the Picture controls that are used to manipulate
    ' the icons must be set at runtime to 32x32.
    '
    ' Position and size Icon within StatusArea
    '
    Edge = (Pic_StatusArea.ScaleWidth - 64) \ 3
    For Y = 0 To 2
        For X = 0 To 1
            Pic_Icons(X + Y * 2).Move Edge + X * (32 + Edge), Edge + Y * (32 + Edge), 32, 32
        Next X
    Next Y

    Pic_Mask.Move 0, 0, 32, 32
    Pic_Image.Move 0, 0, 32, 32
    Pic_TempMask.Move 0, 0, 32, 32
    Pic_TempImage.Move 0, 0, 32, 32
    Pic_Work.Move 0, 0, 32, 32
    Pic_Grid.Move 0, 0, 32, 32
    Pic_ToolPalette.Move 0, 0, 32, 256
    Pic_Undo.Move Pic_StatusArea.ScaleWidth - 34, (Editor.Pic_Icons(5).Top + Editor.Pic_Icons(5).Height + 2.5 * HIGHLIGHT), 32, 32
    
    Get_INI_File_Settings

    ' Adjust size of Editor appropriately for the Video mode.
    '
    
    If Screen.Height = EGA_HEIGHT Then
        '
        ' EGA mode
        '
        
        NewWidth = Width
        NewHeight = Screen.Height
        NewTop = 0
    Else
        ' VGA or Better
        '
        NewWidth = Screen.Width * 0.75
        NewHeight = Screen.Height * 0.75
        NewTop = (Screen.Height - NewHeight) \ 2
    End If
    Move (Screen.Width - NewWidth) \ 2, NewTop, NewWidth, NewHeight

    ' Icon is initially displayed at full maginification
    '
    Scrl_Zoom.Value = Scrl_Zoom.Max

    ' Create Monochrome Hdc and Bitmap to be used when replace the
    ' Icons Monochrome Mask.
    '
    hDCMono = CreateCompatibleDC(hDC)
    hBMMono = CreateCompatibleBitmap(hDCMono, 32, 32)
    hBMOldMono = SelectObject(hDCMono, hBMMono)

    ' The BackColor of the Editing area picture control, Pic_Edit,
    ' is set to Red at design time just so the picture controls contained
    ' within it are visible at design time.
    '
    Pic_Edit.BackColor = WHITE

    ' Select initial tool and brushsize
    '
    CurrentTool = TID_PAINT
    BrushSize = 1
  
    ' Initialize Default Mouse Colors
    '
    MouseColors(0) = BLACK  'Left
    MouseColors(1) = RED    'Right
    MouseColors(2) = WHITE  'Screen
    MouseColors(3) = BLACK  'Inverse

    ' Add Accelerators that are not available from the Menu Design window
    ' to the appropriate menu items.
    '
    Menu_FileSelection(MID_EXIT).Caption = "E&xit" + A_TAB + "Alt+F4"
    'Menu_EditSelection(MID_UNDO).Caption = "&Undo" + A_TAB + "Alt+Bksp"
    'Menu_EditSelection(MID_CUT).Caption = "Cu&t" + A_TAB + "Shift+Del"
    'Menu_EditSelection(MID_COPY).Caption = "&Copy" + A_TAB + "Ctrl+Ins"
    'Menu_EditSelection(MID_PASTE).Caption = "&Paste" + A_TAB + "Shift+Ins"
    'Menu_EditSelection(MID_PASTE_OPAQUE).Caption = "Paste &Opaque" + A_TAB + "Ctrl+Shift+Ins"
    'Menu_EditSelection(MID_DELETE).Caption = "&Delete" + A_TAB + "Del"

    ' Initialize all 6 Icons filename to UNTITLED, so a File.Save opertion
    ' will react just like a File.Save As, the first time.
    '
    For i = 0 To 5
      ICONINFO(i).FileName = UNTITLED
    Next i
    
    EditorLoaded = True

End Sub

' Just for an added convience when selecting tools, Clicking anywhere
' on the Editor that is not covered by another control, will display
' the Tools menu as a pupup menu at that location.
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Menu is displayed on if the left button is used.
    '
    If (Button = RIGHT_BUTTON) And Menu_ViewSelection(MID_POPUP_MENU).Checked Then
        '
        ' The X and Y coordintes passed to the TrackPopupMenu() API routine
        ' are Screen coordinates in pixels, not Form coordinates, so, the
        ' X and Y coordinates of the Mouse must be adjusted, taking into
        ' account the position of the form.  And for a little more added
        ' convience, another adjustment is made so when the menu pops up,
        ' the mouse poiner is directly over the currently selected tool.
        '
        InPixels = ScaleWidth
        ScaleMode = TWIPS
        TwipsToPixels = ScaleWidth \ InPixels
        Borders = (Width - ScaleWidth)
        TitleBarAndMenu = (Height - ScaleHeight - Borders)
        IX = X + (Left \ TwipsToPixels)
        IY = Y + (Top + TitleBarAndMenu - CurrentTool * (TitleBarAndMenu \ 2)) \ TwipsToPixels
        ScaleMode = PIXELS
        Menu_ToolsSelection(CurrentTool).Checked = True
        R = TrackPopupMenu(GetSubMenu(GetMenu(hWnd), 3), 0, IX, IY, 0, hWnd, 0)
    End If

End Sub

' Resizing of the Editor can, depending on the new size, force any
' of the palettes to be hidden, so all major controls must be
' resized, repositioned and possibly hidden in reponse to resizing
' the Editor.
'
Private Sub Form_Resize()
Static LastWindowState

    ' Nothing is done if the Editor is minimized.
    '
    If WindowState = MINIMIZED Then
        '
        ' Hide ColorPalette since it should not be visible while the
        ' Editor is Minimized.
        '
        If ColorPaletteLoaded Then ColorPalette.Hide

    ElseIf ScaleHeight >= 34 Then
        '
        ' Hide the Editing area while resizing other controls
        ' to prevent uneccessary repaints.
        '
        Pic_Edit.Visible = False

        ' Adjust all major controls
        '
        Adjust_Status_Bar
        Adjust_Tool_Palette
        Adjust_Color_Palette
        Adjust_ScrollBar
    
        ' Hide or Show the major controls, based on their corresponding View
        ' menu item Checked value, which was set within the above *Adjust*
        ' routines.
        '
        Pic_ColorPalette.Visible = Menu_ViewSelection(MID_COLOR_PALETTE).Checked
        Pic_ToolPalette.Visible = Menu_ViewSelection(MID_TOOL_PALETTE).Checked
        Pic_StatusArea.Visible = Menu_ViewSelection(MID_STATUS_BAR).Checked
        Scrl_Zoom.Visible = Menu_ViewSelection(MID_ZOOM_SCROLLBAR).Checked
        
        If (LastWindowState = MINIMIZED) And ColorPaletteLoaded Then ColorPalette.Show

    End If

    LastWindowState = WindowState

End Sub

' Editor is being terminated
'
Private Sub Form_Unload(Cancel As Integer)
    
    ' Check fo any icons that have been modified since they
    ' were last saved.
    '
    Text = ""
    For i = 0 To 5
        If ICONINFO(i).Changed Then Text = Text + "Icon #" + Right$(Menu_IconsSelection(i).Caption, Len(Menu_IconsSelection(i).Caption) - 1) + CRLF
    Next i
    
    If Text <> "" Then
        '
        ' One or more icons were changed, so prompt user.
        '
        Text = Text + CRLF + "The icons listed have not been saved since their last modification.  Quit anyway?"
        Cancel = MsgBox(Text, 36, "ICONS HAVE CHANGED") = MBNO
    End If

    If Not Cancel Then
        '
        ' User wants to terminate Editor, so clean up.
        '
        EditorLoaded = False
        Unload ColorPalette
        SelectObject hBMMono, hBMOldMono
        DeleteObject hBMMono
        DeleteDC hDCMono
        WinHelp hWnd, dummy$, HELP_QUIT, 0
        Save_Settings_To_INI_File
        '
        ' If Editor was started up first (Null command line or anything other than
        ' "v" or "V") then we treat it as the main Form.  So, if the Viewer is loaded,
        ' we should as the user if the Viewer should also be terminated.
        '
        If (MainForm = ICONWORKS_EDITOR) And ViewerLoaded Then
            '
            ' Editor was started first and the Viewer is loaded so ask the user
            ' if the Viewer should also be terminated.
            '
            Text = "Terminate Viewer also?"
            If MsgBox(Text, 36, "IconWorks") = MBYES Then Unload Viewer
            MainForm = ICONWORKS_VIEWER
        End If
    End If

End Sub

' In response to a MouseDown event in the Edit area, the first thing
' that must be done is deterine what color to use when drawing the object.
Private Sub Get_Current_Colors(Button As Integer)
Dim Index As Integer
    ' Determine which of the 4 currently selected colors to use, and
    ' assign to the Edit area's foreground color.
    Index = Opt_Mouse(SCREEN_COLORS).Value * (-2) + Button - 1
    Pic_Edit.ForeColor = MouseColors(Index)
    ' If using the Paint tool, we must use the Neaest Solid color to the selected color.
    If CurrentTool = TID_PAINT Then Pic_Edit.ForeColor = GetNearestColor(hDC, Pic_Edit.ForeColor)
    If Opt_Mouse(MOUSE_COLORS).Value Then
        Pic_Image.ForeColor = Pic_Edit.ForeColor
    Else
        Pic_Image.ForeColor = (Button = RIGHT_BUTTON) And WHITE
    End If
    Pic_Mask.ForeColor = Opt_Mouse(SCREEN_COLORS).Value And WHITE
    Pic_Image.FillColor = Pic_Image.ForeColor
    Pic_Mask.FillColor = Pic_Mask.ForeColor
    Pic_Edit.FillColor = Pic_Edit.ForeColor
End Sub

' This routine is called to read the color value strings from
' the IconWrks.INI file, and then have The values extracted
Private Sub Get_INI_File_Colors()
    ColorString = Space$(144)
    R = GetPrivateProfileString(APP_NAME, KEY_COLOR + "1", DEFAULT_COLORS_ROW1, ColorString, Len(ColorString), INI_FILENAME)
    Extract_Colors_From_String ColorString, 0, False
    R = GetPrivateProfileString(APP_NAME, KEY_COLOR + "2", DEFAULT_COLORS_ROW2, ColorString, Len(ColorString), INI_FILENAME)
    Extract_Colors_From_String ColorString, 16, False
    R = GetPrivateProfileString(APP_NAME, KEY_COLOR + "3", DEFAULT_COLORS_ROW3, ColorString, Len(ColorString), INI_FILENAME)
    Extract_Colors_From_String ColorString, 32, True
End Sub

' This routine is called only once, which is at load time to obtain
' all settings saved to the IconWrks.INI file.
Private Sub Get_INI_File_Settings()
Dim Index As Integer
    Menu_ViewSelection(MID_STATUS_BAR).Checked = -GetPrivateProfileInt(APP_NAME, KEY_STATUS_BAR, 1, INI_FILENAME)
    Menu_ViewSelection(MID_TOOL_PALETTE).Checked = -GetPrivateProfileInt(APP_NAME, KEY_TOOL_PALETTE, 1, INI_FILENAME)
    Menu_ViewSelection(MID_COLOR_PALETTE).Checked = -GetPrivateProfileInt(APP_NAME, KEY_COLOR_PALETTE, 1, INI_FILENAME)
    Menu_ViewSelection(MID_ZOOM_SCROLLBAR).Checked = -GetPrivateProfileInt(APP_NAME, KEY_ZOOM_SCROLLBAR, 1, INI_FILENAME)
    Menu_ViewSelection(MID_GRID).Checked = -GetPrivateProfileInt(APP_NAME, KEY_GRID, 0, INI_FILENAME)
    Menu_ViewSelection(MID_LINE_GRID).Checked = -GetPrivateProfileInt(APP_NAME, KEY_LINE_GRID, 1, INI_FILENAME)
    Menu_ViewSelection(MID_DOTTED_GRID).Checked = -GetPrivateProfileInt(APP_NAME, KEY_DOTTED_GRID, 0, INI_FILENAME)
    Menu_ViewSelection(MID_FOCUS).Checked = -GetPrivateProfileInt(APP_NAME, KEY_FOCUS, 1, INI_FILENAME)
    Menu_ViewSelection(MID_BORDER).Checked = -GetPrivateProfileInt(APP_NAME, KEY_BORDER, 1, INI_FILENAME)
    Menu_ViewSelection(MID_POPUP_MENU).Checked = -GetPrivateProfileInt(APP_NAME, KEY_POPUP, 1, INI_FILENAME)
    ' Get all Checked Color Menu item values
    Menu_ColorSelection(MID_DEFAULT_COLORS).Checked = -GetPrivateProfileInt(APP_NAME, KEY_DEFAULT, 1, INI_FILENAME)
    Menu_ColorSelection(MID_SOLID_COLORS).Checked = -GetPrivateProfileInt(APP_NAME, KEY_SOLID, 0, INI_FILENAME)
    Menu_ColorSelection(MID_INI_COLORS).Checked = -GetPrivateProfileInt(APP_NAME, KEY_INI, 0, INI_FILENAME)
    ' Enable or Disable the Grid type menu items based on the Check value
    Menu_ViewSelection(MID_LINE_GRID).Enabled = Menu_ViewSelection(MID_GRID).Checked
    Menu_ViewSelection(MID_DOTTED_GRID).Enabled = Menu_ViewSelection(MID_GRID).Checked
    ' Set initial color palette
    If Menu_ColorSelection(MID_DEFAULT_COLORS).Checked Then Index = MID_DEFAULT_COLORS
    If Menu_ColorSelection(MID_SOLID_COLORS).Checked Then Index = MID_SOLID_COLORS
    If Menu_ColorSelection(MID_INI_COLORS).Checked Then Index = MID_INI_COLORS
    Menu_ColorSelection_Click Index
End Sub

' In response to an Edit.Invert menu item selection, this routine
' Inverts the Colors of the selection region.  Using a DrawMode
' of INVERSE, a filled box is drawn over the selected region, which
' inverts the colors.
Private Sub Invert_Selection()
    Pic_TempImage.DrawMode = INVERSE
    Pic_Work.DrawMode = INVERSE
    ' Invert Selected Region.  Do not need to Invert Mask.
    Pic_TempImage.Line (X1SelectFrom, Y1SelectFrom)-(X2SelectFrom, Y2SelectFrom), , BF
    Pic_Work.Line (X1SelectFrom, Y1SelectFrom)-(X2SelectFrom, Y2SelectFrom), , BF
    Pic_TempImage.DrawMode = COPY_PEN
    Pic_Work.DrawMode = COPY_PEN
    If MovingSelection Then Draw_Selection_Rectangle Else Commit_Selection False
    Magnify_Icon X1SelectFrom, Y1SelectFrom, (X2SelectFrom), (Y2SelectFrom)
End Sub

' Once an object, Line, Square, Filled Square, Circle, or Filled Circle
' has been size to what the user wants, this routine is called from the
' MouseUp event for the Editing area (Pic_Edit) to display the actual
' object.
Private Sub Line_To_Filled_Circle_MouseUp(X1, Y1, X2, Y2)
    Pic_Image.DrawWidth = BrushSize
    Pic_Mask.DrawWidth = BrushSize
    If CurrentTool = TID_FILLED_SQUARE Or CurrentTool = TID_FILLED_CIRCLE Then
        Pic_Image.FillStyle = SOLID
        Pic_Mask.FillStyle = SOLID
    End If
    ' Scale the dimensions of the object from the Editing area down
    ' to that of the Icon, 32x32.
    Scale_Region True, X1, Y1, X2, Y2, False
    Select Case CurrentTool
        Case TID_LINE
            Pic_Image.Line (X1, Y1)-(X2, Y2)
            '
            ' Draw end-point only if BrushSize is one Pixel
            If BrushSize = 1 Then Pic_Image.PSet (X2, Y2)
            Pic_Mask.Line (X1, Y1)-(X2, Y2)
            If BrushSize = 1 Then Pic_Mask.PSet (X2, Y2)
        Case TID_SQUARE, TID_FILLED_SQUARE
            Pic_Image.Line (X1, Y1)-(X2, Y2), , B
            Pic_Mask.Line (X1, Y1)-(X2, Y2), , B
        Case TID_CIRCLE, TID_FILLED_CIRCLE
            If Valid_Circle(X1, Y1, X2, Y2) Then
                Pic_Image.Circle (XCenter, YCenter), Radius, , , , Aspect
                Pic_Mask.Circle (XCenter, YCenter), Radius, , , , Aspect
                Radius = 0
            End If
    End Select
    If X1 > X2 Then Swap_Values X1, X2
    If Y1 > Y2 Then Swap_Values Y1, Y2
    Pic_Image.DrawWidth = 1
    Pic_Mask.DrawWidth = 1
    Pic_Image.FillStyle = TRANSPARENT
    Pic_Mask.FillStyle = TRANSPARENT

    ' Only the area modified is redisplayed, which is the rectangle
    ' defined by the values of X1, Y1, X2, Y2.  If the Brushsize is
    ' greater than 1 pixel, then the actual area modified is slightly
    ' larger than the this rectangle, so the values X1, Y1, X2, Y2
    ' must be adjusted to include the thickness of the Brushsize so
    ' the entire area modified is redisplayed.
    '
    If (BrushSize > 1) And (CurrentTool >= TID_LINE) Then
        Adjustment = BrushSize \ 2
        X1 = X1 - Adjustment - 1
        Y1 = Y1 - Adjustment - 1
        X2 = X2 + Adjustment + 1
        Y2 = Y2 + Adjustment + 1
    End If

End Sub

' For the Line, Square, Filled Square, Circle, and Filled Circle Tools,
' an Inverted resizeable version of the object is displayed in response
' to a MouseMove.  The object is anchored at the point set by the MouseDown
' event within the Editing area.  This routine is called to erase the
' old object and display the new object during a MouseMove event.
'
Private Sub Line_To_Filled_Circle_Tool_Move(Shift As Integer, X As Single, Y As Single)
Static LastCircleGood

    ' Erase Old object
    '
    Draw_Line_Or_Box
    If (CurrentTool >= TID_CIRCLE) And LastCircleGood Then Pic_Edit.Circle (XCenter, YCenter), Radius, , , , Aspect
    Adjust_X_and_Y X, Y, X2Region, Y2Region, TO_CENTER

    ' Holding the Shift key down forces a perfect Square or Circle.
    ' For a line, it forces either a Horizontal, Vertical or 45 degree
    ' line to be displayed.
    '
    If Shift And SHIFT_MASK Then Make_Region_Square

    If CurrentTool >= TID_CIRCLE Then
        '
        ' When sizing a circle, a rectangle with the circle displayed
        ' within the circle is drawn.  This is done to give a better
        ' feeling that the mouse is attached to the object when sizing.
        ' The rectanle is not drawn one the mouse is released.
        '
        LastCircleGood = Valid_Circle(X1Region, Y1Region, X2Region, Y2Region)
        If LastCircleGood Then Pic_Edit.Circle (XCenter, YCenter), Radius, , , , Aspect
    End If

    ' Draw New Object
    '
    Draw_Line_Or_Box

End Sub

' To make creating a perfect Square, Circle, or a Diagonal, Vertical,
' or Horizontal line, the routine is called whenever the Shift key
' is depressed in conjuction to dragging the Mouse within the Editing
' area.  This routine adjusts the actual mouse coordinates so that
' the displayed object is Square, or in the case of a line, Diagonal,
' Vertical, or Horizontal.
'
Private Sub Make_Region_Square()
    
    Xside = Abs(X2Region - X1Region)
    Yside = Abs(Y2Region - Y1Region)
    If Xside > Yside Then RectSide = Xside Else RectSide = Yside
    X2Region = X1Region + RectSide * Sgn(X2Region - X1Region)
    Y2Region = Y1Region + RectSide * Sgn(Y2Region - Y1Region)

End Sub

' Changes size of brush to either 1x1, 3x3, or 5x5, and
' Checks the corresponding Bush menu item
'
Private Sub Menu_BrushSelection_Click(Index As Integer)
  
    Menu_BrushSelection(BrushSize).Checked = False
    Menu_BrushSelection(Index).Checked = True
    BrushSize = Index

End Sub

' Processes all Color Menu selections.
'
Private Sub Menu_ColorSelection_Click(Index As Integer)
    
    Select Case Index

        Case MID_COLORS
            '
            ' Invoke Custom ColorPalette form
            '
            ColorIndex = 0
            ColorPalette.Show

        Case MID_DEFAULT_COLORS
            '
            ' Default color palette is the same colors as the VB colorpalette
            '
            Extract_Colors_From_String DEFAULT_COLORS_ROW1, 0, False
            Extract_Colors_From_String DEFAULT_COLORS_ROW2, 16, False
            Extract_Colors_From_String DEFAULT_COLORS_ROW3, 32, True
        
        Case MID_SOLID_COLORS
            '
            ' Solid colors are the 16 colors returned from QBColors()
            '
            Menu_ColorSelection(MID_SOLID_COLORS).Checked = True
            Show_Solid_Colors_Only

        Case MID_INI_COLORS
            '
            ' If there are Colors saved to the INI file, they will
            ' be displayed, otherwise, the Default colors will be displayed
            '
            Menu_ColorSelection(MID_INI_COLORS).Checked = True
            Get_INI_File_Colors
 
        Case MID_SAVE_COLORS
            Save_Colors_To_INI_File

        Case MID_LEFT_RIGHT, MID_SCREEN_INVERSE
            '
            ' Same functionality as clicking either of the Mouse Option buttons
            '
            Opt_Mouse(Index - MID_LEFT_RIGHT).Value = True
    
    End Select

    ' If the selection was 1 of the 3 colorpalette selections, Check
    ' the current selection, and Uncheck the previous selection.
    '
    If (Index >= MID_DEFAULT_COLORS) And (Index <= MID_INI_COLORS) Then
        Menu_ColorSelection(MID_DEFAULT_COLORS).Checked = (Index = MID_DEFAULT_COLORS)
        Menu_ColorSelection(MID_SOLID_COLORS).Checked = (Index = MID_SOLID_COLORS)
        Menu_ColorSelection(MID_INI_COLORS).Checked = (Index = MID_INI_COLORS)
    End If

End Sub

' Prepares the Edit menus Sub menu items, by Enabling or Disabling
' the menu items depending on whether there is a current selection
' or if the current selection is being moved, and the current format
' of the system ClipBoard.
'
Private Sub Menu_Edit_Click()

    Menu_EditSelection(MID_CUT).Enabled = Selecting And Not MovingSelection
    Menu_EditSelection(MID_COPY).Enabled = Selecting And Not MovingSelection
    Menu_EditSelection(MID_PASTE).Enabled = Clipboard.GetFormat(CF_BITMAP)
    Menu_EditSelection(MID_PASTE_OPAQUE).Enabled = Clipboard.GetFormat(CF_BITMAP)
    Menu_EditSelection(MID_DELETE).Enabled = Selecting And Not MovingSelection
    Menu_EditSelection(MID_FLIP_HORIZONTAL).Enabled = Selecting
    Menu_EditSelection(MID_FLIP_VERTICAL).Enabled = Selecting
    Menu_EditSelection(MID_ROTATE_LEFT).Enabled = Selecting And (Not MovingSelection) And ((X2Region - X1Region) = (Y2Region - Y1Region))
    Menu_EditSelection(MID_ROTATE_RIGHT).Enabled = Selecting And (Not MovingSelection) And ((X2Region - X1Region) = (Y2Region - Y1Region))
    Menu_EditSelection(MID_INVERT).Enabled = Selecting

End Sub

' Processes all Edit Menu selections.
'
Private Sub Menu_EditSelection_Click(Index As Integer)
    Select Case Index

        Case MID_UNDO
            UnDo_Edits

        Case MID_CUT, MID_COPY, MID_DELETE
            Cut_Copy_Or_Delete_Selection Index
            
        Case MID_PASTE, MID_PASTE_OPAQUE
            Paste_ClipBoard_Contents Index

        Case MID_SELECT_ALL
            Select_Entire_Icon

        Case MID_FLIP_HORIZONTAL, MID_FLIP_VERTICAL
            Flip_Selection Index

        Case MID_ROTATE_RIGHT, MID_ROTATE_LEFT
            Rotate_Selection Index

        Case MID_INVERT
            Invert_Selection

    End Select
    
End Sub

' Processes all File Menu selections.
'
Private Sub Menu_FileSelection_Click(Index As Integer)

    Select Case Index

        Case MID_NEW
            Prepare_For_New_Icon
      
        Case MID_OPEN
            Viewer.Show MODELESS

        Case MID_SAVE, MID_SAVE_AS
            Save_Icon Index

        Case MID_EXIT
            Unload Editor

    End Select

End Sub

' Processes all Help Menu selections.
'
Private Sub Menu_HelpSelection_Click(Index As Integer)

    If Index = MID_ABOUT Then
        AboutBox.Show MODAL
    Else
        Get_Help Index
    End If

End Sub

' Processes all Icons Menu selections.  Same as clicking
' one of the Icons within the StatusArea.
'
Private Sub Menu_IconsSelection_Click(Index As Integer)

    Pic_Icons_MouseDown Index, LEFT_BUTTON, 0, 0, 0
  
End Sub

' Prepares for the Tools sub menu to be shown, by Checking
' the currently selected tool.
'
Private Sub Menu_Tools_Click()

    Menu_ToolsSelection(CurrentTool).Checked = True

End Sub

' Processes all Tools Menu selections.  Same as clicking any
' tool within the ToolPalette
'
Private Sub Menu_ToolsSelection_Click(Index As Integer)

    Pic_ToolPalette_MouseDown LEFT_BUTTON, 0, 0, Index * 32
    
End Sub

' Processes all View Menu selections.
'
Private Sub Menu_ViewSelection_Click(Index As Integer)

    Select Case Index

        Case MID_STATUS_BAR, MID_TOOL_PALETTE, MID_COLOR_PALETTE, MID_ZOOM_SCROLLBAR, MID_FOCUS, MID_BORDER, MID_POPUP_MENU
            Menu_ViewSelection(Index).Checked = Not Menu_ViewSelection(Index).Checked
            If Index = MID_BORDER Then
                HighLight_Current_Icon
            ElseIf Index < MID_FOCUS Then
                '
                ' Must process a From_Resize to make the Above selection
                ' take affect since all hiding and showing of palettes is
                ' done within the Form_Resize event.
                '
                Form_Resize
            End If

        Case MID_SHOW_ALL
            For i = 0 To 3
                Menu_ViewSelection(i).Checked = True
            Next i
            Form_Resize

        Case MID_GRID
            Activate_Grid

        Case MID_LINE_GRID, MID_DOTTED_GRID
            Select_Grid_Type Index

        Case MID_ZOOM_IN
            Scrl_Zoom.Value = Scrl_Zoom.Value + 32

        Case MID_ZOOM_OUT
            Scrl_Zoom.Value = Scrl_Zoom.Value - 32
      
    End Select
  
    If ((Index = MID_STATUS_BAR) Or (Index = MID_SHOW_ALL)) And Menu_ViewSelection(MID_STATUS_BAR).Checked Then Display_Color_Palette Pic_ColorPalette

End Sub

' One of the Mouse Button Options Buttons was selected, so update
' the corresponding Color menu items Checked value.
'
Private Sub Opt_Mouse_Click(Index As Integer)
    
    Menu_ColorSelection(MID_LEFT_RIGHT).Checked = Not (-Index)
    Menu_ColorSelection(MID_SCREEN_INVERSE).Checked = -Index

End Sub

' The paint tools is the only tool that paints directly to the
' Editing area during the MouseMove event.  This routine draws
' the pixels to the Icons Image and Mask, and to the Editing area,
' so, there is no need to redisplay the modified area of the Icon
' in response to a MouseUp event.  However, the StatusArea Icon
' will not display the modificatioin until a MouseUp event.
'
Private Sub Paint_Tool_Move(Xpix, Ypix)
                
    ' Adjust Mouse position so it is centered within the Pixel being
    ' drawn.
    '
    Adjustment = BrushSize \ 2
    X1 = Xpix - Adjustment
    Y1 = Ypix - Adjustment
    X2 = Xpix + Adjustment
    Y2 = Ypix + Adjustment

    ' Draw pixel on Icons Image and Mask
    '
    Pic_Image.Line (X1, Y1)-(X2, Y2), , BF
    Pic_Mask.Line (X1, Y1)-(X2, Y2), , BF

    ' Draw pixel on Editing area
    '
    X2 = X2 + 1
    Y2 = Y2 + 1
    Pic_Edit.Line (X1 * PixelSize, Y1 * PixelSize)-(X2 * PixelSize - 1, Y2 * PixelSize - 1), , BF

    ' Redisplay Grid in modified area if Grid is selected.
    '
    If Menu_ViewSelection(MID_GRID).Checked Then Display_Grid (Pic_Edit.hDC), X1, Y1, X2, Y2

End Sub

' In response to an Edit.Paste or Edit.Paste Opaque, this routine is called
' to perform the acutal paste operation.  There are two forms of the Paste:
'
'   Paste Shift+Ins:  Creates a Mask from the bitmap in the clipboard
'                     based on the current Screen Color selected.  All
'                     Pixels in the Bitmap that match the Screen Color
'                     become transparent.
'
'   Paste Opaque Ctrl+Shift+Ins:  No mask is created from the bitmap and
'                     It is pasted as is with no Screen Color attributes.
'
Private Sub Paste_ClipBoard_Contents(Index)
            
    ' Get Bitmap from System Clipboard
    '
    Pic_Work.Picture = Clipboard.GetData(CF_BITMAP)

    ' Determine if Bitmap is too large to paste.  Must be 32x32 pixels or
    ' smaller.
    '
    If Pic_Work.Width <= 32 And Pic_Work.Height <= 32 Then
        '
        ' Change tool selection to SelectTool, which will allow us to
        ' move the Bitmap being pasted to anywhere within the Icon.
        '
        Pic_ToolPalette_MouseDown LEFT_BUTTON, 0, 0, TID_SELECT * 32
        
        Pic_TempImage.Picture = Pic_Work.Image
        Opaque = MID_PASTE - Index

        If Opaque Then
            ' An Opaque paste, so the Bitmap is pasted as is.  So we simply create
            ' a solid Black mask.
            '
            Pic_TempMask.Line (0, 0)-(Pic_TempMask.ScaleWidth, Pic_TempMask.ScaleHeight), BLACK, BF
        Else
            ' A normal paste, so a Mask must be created from the Bitmap being pasted,
            ' based on the currently selected Screen color.  The Mask is created, by
            ' BitBlt'ing the the 16-color bitmap to a mono Hdc.  All pixels in the color
            ' bitmap which are the screen color will be white in the monochrome bitmap,
            ' all others will be black.  That is why the Backcolor of the Picture control
            ' containing the bitmap is first set to the currently selected Screen color.
            ' This Monochrome bitmap is then BitBlt'd back to a 16-Color Picture control
            ' so we can easily manipulate it.
            '
            Pic_Work.BackColor = MouseColors(2)
            R = BitBlt(hDCMono, 0, 0, Pic_Work.Width, Pic_Work.Height, Pic_Work.hDC, 0, 0, SRCCOPY)
            R = BitBlt(Pic_TempMask.hDC, 0, 0, Pic_Work.Width, Pic_Work.Height, hDCMono, 0, 0, SRCCOPY)
                    
            ' Once the Mask is created, all the pixels in the Image bitmap which are of
            ' currently selected screen color must be changed to Black.  To do this we
            ' Logically AND the Mask created above with the Image bitmap.  This has the
            ' affect of turning all pixels which are not the screen color to Black, leaving
            ' screen color pixels unaffected.  We then XOR the image with this result, which
            ' gives us the desired result of all screen color pixels of the origianl
            ' bitmap going to Black, all other pixels are unaffected.
            '
            R = BitBlt(Pic_TempImage.hDC, 0, 0, Pic_Work.Width, Pic_Work.Height, Pic_TempMask.hDC, 0, 0, SRCAND)
            R = BitBlt(Pic_TempImage.hDC, 0, 0, Pic_Work.Width, Pic_Work.Height, Pic_Work.hDC, 0, 0, SRCINVERT)
        End If

        If (Pic_Work.Width = 32) And (Pic_Work.Height = 32) Then
            ' Bitmap Pasted is too large to allow it to be moved, so we
            ' commit the Bitmap to the Icon with out allowing it to be
            ' moved, and then redisplay the Icon.
            '
            Scale_Region False, 0, 0, 32, 32, False
            Scale_Region True, X1SelectFrom, Y1SelectFrom, X2SelectFrom, Y2SelectFrom, False
            Commit_Selection True
            Magnify_Icon 0, 0, 31, 31
        Else
            ' The Bitmap pasted can be moved, so go into Move mode by setting
            ' the appropriate flags.
            '
            Selecting = True
            
            MovingSelection = True

            ' Set the Selection Region to the size of the Bitmap pasted, in
            ' both the scale of the Icon and the Magnified Icon.
            '
            Scale_Region False, 0, 0, (Pic_Work.Width), (Pic_Work.Height), False
            Scale_Region True, X1SelectFrom, Y1SelectFrom, X2SelectFrom, Y2SelectFrom, True
            Scale_Region True, LastX1, LastY1, LastX2, LastY2, True

            ' Redisplay Icon.  Passing a values of 0,0,0,0 simply force the Magnify_Icon
            ' routine to magnify the pasted Bitmap over the top of the magnified Icon.
            '
            Magnify_Icon 0, 0, 0, 0
        End If
    Else
        ' Bitmap was too large to paste, so prompt the user.
        '
        Pic_Work.Picture = LoadPicture()
        MsgBox "ClipBoard image is too large:  " + Format$(Pic_Work.Width) + " x " + Format$(Pic_Work.Height) + CRLF + "Image not pasted", 16, "Error"
    End If

End Sub

' ColorPalette operations are performed on the MouseUp event,
' and DoubleClicking the ColorPalette invokes the Custom
' color ColorPalette Form so a flag must be set to inform the
' MouseUp event to invoke the Custom ColorPalette form.
'
Private Sub Pic_ColorPalette_DblClick()

    DoubleClicked = True

End Sub

' Checks if Mouse Coordinates are within the ColorPalette, since the mouse
' could have been dragged outside of the ColorPalette between the MouseDown
' and MouseUp events, in which case the ColorPalette would still get the
' MouseUp event with coordinates outside the actual ColorPalette which would
' generate and error when accessing the Color Array.
'
Private Sub Pic_ColorPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (X >= 0) And (X <= 16) And (Y >= 0) And (Y <= 3) Then Update_Mouse_Colors Button, X, Y

End Sub

Private Sub Pic_ColorPalette_Paint()

    Display_Color_Palette Pic_ColorPalette

End Sub

' If the Icon dropped is from the Viewer, it is loaded into the
' currently selected Icon, otherwise nothing is done.
'
Private Sub Pic_Edit_DragDrop(Ctl As Control, X As Single, Y As Single)

    If Ctl.Parent.Tag <> Editor.Tag Then Load_An_Icon

End Sub

' If the Icon being dragged is from the Viewer and it has just entered
' or left the Edit area, invert the Edit area.
'
Private Sub Pic_Edit_DragOver(Ctl As Control, X As Single, Y As Single, State As Integer)

    If (Ctl.Parent.Tag <> Editor.Tag) And (State <> OVER) Then Invert_Control Pic_Edit

End Sub

' A MouseDown event in the Editing area signals the beginning of a new
' Tool operation if a tool operation is not already in progress.  All
' preparation for the Tool operation is performed here, within the
' MouseDown event.
'
Private Sub Pic_Edit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Propare for Tool operation only if there is not an operation already in
    ' progress, and only one mouse button is down, Left or Right.
    '
    If (ToolInProcess = 0) And ((Button = LEFT_BUTTON) Or (Button = RIGHT_BUTTON)) Then
        '
        ' Save Button value that started Tool Operation.
        '
        ToolInProcess = Button

        Get_Current_Colors Button
        Pic_Edit.DrawStyle = SOLID

        If CurrentTool = TID_PAINT Then
            '
            ' Begin a Paint tool operation.
            '
            Pic_Edit.DrawMode = COPY_PEN
            
            ' So a pixel is drawn in repsonse to just a Click within the editing
            ' area, the MouseMove event must be called since all the actual
            ' drawing to the Editing area for the Paint Tool is done within the
            ' MouseMove event.
            '
            Pic_Edit_MouseMove Button, 0, X, Y

        ElseIf CurrentTool <> TID_FILL Then
            '
            ' Prepare for all other tools other than Fill and Paint tools.
            '
            Pic_Edit.DrawMode = INVERSE
            If CurrentTool = TID_SELECT Then Adjustment = TO_EDGE Else Adjustment = TO_CENTER

            ' Determine if a Selection is in process, the Selection can be moved,
            ' and whether the MousePointer is within the current selection.
            '
            If Selecting And Moveable And ((X >= X1Region) And (X <= X2Region) And (Y >= Y1Region) And (Y <= Y2Region)) Then
                '
                ' The selection is moveable and the MousePointer is within the
                ' selection.  So prepare for a move.
                '
                Adjust_X_and_Y X, Y, XMoveSelection, YMoveSelection, Adjustment
                If Button = RIGHT_BUTTON Then Opaque = True Else Opaque = False

                If Not MovingSelection Then
                    '
                    ' The selection has not yet been moved from its original
                    ' position, so prepare to move selection.
                    '
                    ' If the Shift key is down, then the a copy of the Selection
                    ' will be moved, otherwise the selection will be move from
                    ' its current location so its previous location must be deleted.
                    '
                    If (Shift And SHIFT_MASK) = 0 Then Delete_Selection_From_Icon
                    Draw_Selection_Rectangle
                    MovingSelection = True
                    Draw_Selection_Rectangle

                ElseIf (Shift And SHIFT_MASK) Then
                    '
                    ' The selection has already been moved from its original
                    ' location and is being moved again.  If the Shift key is
                    ' down, make a copy of the selection at its current location
                    ' and continue the move.
                    '
                    Commit_Selection False
                    Draw_Selection_Rectangle
                End If
            Else
                '
                ' No selection is in progress, or a selection is in progress
                ' but the MousePointer is not within the current selection,
                ' or a tool other than the Select tool is in use.
                '
                ' If a Selection is in progress, commit the Selection at its
                ' current location.
                '
                If Selecting Then Commit_Selection True

                ' Prepare for new Tool operation
                '
                Opaque = True
                Adjust_X_and_Y X, Y, X1Region, Y1Region, Adjustment
                X2Region = X1Region
                Y2Region = Y1Region
            End If
        End If

    End If

End Sub

' For all tools except for the Paint tool and Fill tool, the MouseMove
' event is used to size the object to be drawn.  The actual object is
' drawn within the MouseUp event based on the coodinates obtained here
' in the MouseMove event.  The Paint tool, however, does all of its
' actual drawing to the Editing area within the MouseMove event.  The
' Fill tool does nothing until the MouseUp event.
'
Private Sub Pic_Edit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    ' Calculate pixel mouse is currently over.
    '
    Xpix = X \ PixelSize
    Ypix = Y \ PixelSize

    ' The Mouse coordinates are always displayed within the StatusArea
    '
    Display_Mouse_Coordinates Xpix, Ypix

    ' Depending on the Current tool, and in the case of the Select tool,
    ' the state of the Selection, the MousePointer may change.
    '
    Set_MousePointer X, Y

    ' We only want to process the MouseMove event if the Button causing
    ' the MouseMove event is the same that started the operation.
    ' The value of the Button that started the operation is stored in
    ' the global variable ToolInProcess which is set in the MouseDown event.
    '
    If (Button = ToolInProcess) And (ToolInProcess <> 0) Then
    
        Select Case CurrentTool

            Case TID_SELECT
                Select_Tool_Move Shift, X, Y

            Case TID_PAINT
                Paint_Tool_Move Xpix, Ypix

            Case TID_LINE To TID_FILLED_CIRCLE
                Line_To_Filled_Circle_Tool_Move Shift, X, Y

        End Select

    End If

End Sub

' For all tools except for the Paint tool, the actual object drawn
' is not drawn until the Mouse Button is released, which is within
' the the edit areas MouseUp event.
'
Private Sub Pic_Edit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' There is no operation that occurs with both mouse buttons down,
    ' so check to make sure that the MouseUp event is due to the release
    ' of the same button that started the drawing operation.  The global
    ' variable ToolInProcess holds the value of the Button.
    '
    If Button = ToolInProcess Then

        Pic_Edit.DrawMode = COPY_PEN
      
        Select Case CurrentTool
    
            Case TID_SELECT
                Select_Tool_MouseUp X1, Y1, X2, Y2
    
            Case TID_FILL
                Fill_Tool_MouseUp X, Y, X2, Y2
    
            Case TID_LINE To TID_FILLED_CIRCLE
                Line_To_Filled_Circle_MouseUp X1, Y1, X2, Y2
    
        End Select

        ' The Icon needs to be updated to show the affect of the
        ' Tool operation for all tools except for the Select tool if
        ' just making a selection, since all that occurs is a Rectangle
        ' showing the selection is drawn.
        '
        If Not ((CurrentTool = TID_SELECT) And (Not MovingSelection)) Then
            Update_Icon Pic_Icons(CurrentIcon)
            '
            ' Redisplay the Icon for all tools except for the Paint tool.
            ' The paint tool draws directly on the Edit area within the
            ' MouseMove event.
            '
            If SeletedTool <> TID_PAINT Then Magnify_Icon X1, Y1, X2, Y2
        End If
    
        ToolInProcess = False

        ' Set enabled state of Edit Menu items.
        '
        Menu_Edit_Click

    End If

End Sub

Private Sub Pic_Edit_Paint()

    ' Repaint Edit area
    Magnify_Icon 0, 0, 31, 31

End Sub

' If the Icon dropped is from the Viewer, it is loaded into the
' StatusArea Icon on which is was dropped and that icon is made
' the current Icon, otherwise nothing is done.
'
Private Sub Pic_Icons_DragDrop(Index As Integer, Ctl As Control, X As Single, Y As Single)

    If Ctl.Parent.Tag <> Editor.Tag Then
        CurrentIcon = Index
        Load_An_Icon
    End If

End Sub

' If the Icon being dragged is from the Viewer and it has just entered
' or left one of the StatusArea Icons, invert the StatusArea icon.
'
Private Sub Pic_Icons_DragOver(Index As Integer, Ctl As Control, X As Single, Y As Single, State As Integer)

    If (Ctl.Parent.Tag <> Editor.Tag) And (State <> OVER) Then
        Invert_Control Pic_Icons(Index)
        Pic_Icons(Index).Refresh
    End If

End Sub

' Clicking on any of the icons within the StatusArea, makes that
' Icon the current Icon.  If the the Icon is already the current
' Icon, and selection currently in progress is canceled, and
' the Icon can be dragged to test what it looks like begin dragged.
' It cannot be dropped on anything.
'
Private Sub Pic_Icons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Determine if new Icon is being selected
    '
    If (Button = LEFT_BUTTON) And (Index <> CurrentIcon) Then
        '
        ' New Icon has been selected, so prepare for new Icon
        '
        Menu_IconsSelection(CurrentIcon).Checked = False
        Menu_IconsSelection(Index).Checked = True
        If Selecting Then Commit_Selection True
        CurrentIcon = Index
        Select_New_Icon
    ElseIf (Button = LEFT_BUTTON) And (Index = CurrentIcon) Then
        '
        ' The current Icon was selected.
        '
        If Selecting Then
            '
            ' A selection was in progress so cancel it.
            '
            Draw_Selection_Rectangle
            Selecting = False
            MovingSelection = False
            Scale_Region True, X1, Y1, X2, Y2, True
            Magnify_Icon X1, Y1, X2, Y2
        End If
        
        ' Set the DragIcon of the selected Icon, and begin dragging
        '
        Pic_Icons(Index).DragIcon = Pic_Icons(Index).Picture
        Pic_Icons(Index).Drag
    End If

End Sub

' Although the StatusArea appears as though it contains many controls
' it actually contains only 7 picture controls and 2 Option buttons.
' The rest of the information displayed within the StatusArea is painted
' directly onto the StatusArea.  This is done to limit the actual number
' of controls on the Editor.  The more controls, the more System resources
' it uses and the longer it takes to load.  It may be easier to display
' certain data using controls, but in some cases, its benificial to not
' use controls but display the information yourself.  The easiest controls
' to eliminate are Labels since they usually just display text and do
' not change throughout the life of the program, and the .Print method
' can be used to display the text instead of a label.
'
Private Sub Pic_StatusArea_Paint()

    ' Calculate center of StatusArea
    '
    Middle = Pic_StatusArea.ScaleWidth \ 2

    ' Set font attributes for X & Y Mouse coodinate display
    '
    Pic_StatusArea.FontName = "MS Sans Serif"
    Pic_StatusArea.FontSize = 8.25
    Pic_StatusArea.ForeColor = BLACK

    ' Display X & Y mouse coordinate labels, "X=" and Y="
    '
    Pic_StatusArea.CurrentY = Pic_Icons(5).Top + Pic_Icons(5).Height + HIGHLIGHT + 1
    Pic_StatusArea.CurrentX = (Middle - Pic_StatusArea.TextWidth("XXXX")) \ 2
    Pic_StatusArea.Print "X=";
    Pic_StatusArea.CurrentX = Middle + (Middle - Pic_StatusArea.TextWidth("XXXX")) \ 2
    Pic_StatusArea.Print "Y=";
    
    ' If the Undo Icon is visible, display its label, "Undo to:"
    '
    If Pic_Undo.Visible Then
        Text = "Undo to:"
        Pic_StatusArea.CurrentX = (Pic_Undo.Left - Pic_StatusArea.TextWidth(Text)) \ 2
        Pic_StatusArea.CurrentY = Pic_Undo.Top + ((Pic_Undo.Height - Pic_StatusArea.TextHeight(Text)) \ 2)
        Pic_StatusArea.Print "Undo to:"
    End If
    
    ' If the Mouse Option buttons are visible, then display the the
    ' option button Color labels and the colors themselves.  The labels are
    ' all based on the positions of the two Option buttons, so it they
    ' are moved at design time, the labels will still be displayed
    ' correctly.
    '
    If Opt_Mouse(0).Visible Then
        '
        ' Display Mouse Butons header label within a White box with a Black
        ' outline.
        '
        Text = "Mouse Buttons"
        Y1 = Opt_Mouse(0).Top - 1.25 * Opt_Mouse(0).Height
        Pic_StatusArea.Line (0, Y1)-(Pic_StatusArea.ScaleWidth, Y1 + Pic_StatusArea.TextHeight(Text)), WHITE, BF
        Pic_StatusArea.Line (0, Y1)-(Pic_StatusArea.ScaleWidth - 1, Y1 + Pic_StatusArea.TextHeight(Text)), BLACK, B
        Pic_StatusArea.CurrentX = (Pic_StatusArea.ScaleWidth - Pic_StatusArea.TextWidth(Text)) \ 2
        Pic_StatusArea.CurrentY = Y1
        Pic_StatusArea.FontTransparent = True
        Pic_StatusArea.Print Text
        Pic_StatusArea.FontBold = False

        ' Display the Mouse Button color labels, centered on either
        ' side of the Option buttons.
        '
        For i = 0 To 3
            Select Case i
                Case 0
                    Text = "Left"
                Case 1
                    Text = "Right"
                    Shift = Middle
                Case 2
                    Text = "Screen"
                    Shift = 0
                Case 3
                    Text = "Inverse"
                    Shift = Middle
            End Select
            Pic_StatusArea.CurrentX = Shift + (Middle - Pic_StatusArea.TextWidth(Text)) \ 2
            Pic_StatusArea.CurrentY = Opt_Mouse(i \ 2).Top - Pic_StatusArea.TextHeight(Text) - 1
            Pic_StatusArea.Print Text
        Next i

        ' Display the 4 color squares showing currently selected colors
        ' These could be picture controls.
        '
        Display_Mouse_Colors
    End If

    ' Set StatusArea font attributes back to that of the Mouse coordinates.
    '
    Pic_StatusArea.ForeColor = RED
    Pic_StatusArea.FontName = "Courier"
    Pic_StatusArea.FontBold = True
    Pic_StatusArea.FontTransparent = False

    HighLight_Current_Icon

End Sub

' Performs a "Select All" opertation, just as though the Edit.Select all
' menu item was selected.
'
Private Sub Pic_ToolPalette_DblClick()

    If CurrentTool = TID_SELECT Then Menu_EditSelection_Click MID_SELECT_ALL

End Sub

' Prepare for a new tool.  The actual Tool selection is done within
' the MouseMove event.
'
Private Sub Pic_ToolPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Commit and Selection currently being moved.
    '
    If Selecting Then Commit_Selection True

    ' Commit all edits to Undo Icon
    '
    Update_Icon Pic_Undo

    ' Remove Check from current tool in Tools menu
    '
    Menu_ToolsSelection(CurrentTool).Checked = False

    ' Since the actual selection is done within the MouseMove event, and
    ' a MouseMove event does not occur until the mouse is moved within
    ' the ToolPalette, we need to call the MouseMove event, to ensure
    ' the selected tool is in fact selected, since just clicking on
    ' the ToolPalette does not generate a MouseMove event.
    '
    Pic_ToolPalette_MouseMove Button, Shift, X, Y

End Sub

' Moving the mouse within the ToolPalette after a MouseDown event, will select
' the tool the mouse is currently over, so you can slide up and down the palette,
' and the tool selection will follow the mouse.
'
Private Sub Pic_ToolPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Calculate selected tool.  Each tool Icon is 32 pixels high, so this
    ' is done easily by Integer dividing the Y coordinate by 32.
    '
    NewTool = Y \ 32

    ' Check to make sure the Mouse is within the Y limits of the ToolPalette
    ' and not over the currently selected tool.
    '
    If (Button <> 0) And (NewTool >= TID_SELECT) And (NewTool <= TID_FILLED_CIRCLE) And (NewTool <> CurrentTool) Then
        '
        ' UnHighlight current tool
        '
        Pic_ToolPalette.Line (0, CurrentTool * 32 + 1)-(31, CurrentTool * 32 + 31), , BF

        ' Hightlight new tool
        '
        Pic_ToolPalette.Line (0, NewTool * 32 + 1)-(31, NewTool * 32 + 31), , BF

        CurrentTool = NewTool
    End If
  
End Sub

Private Sub Pic_ToolPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Set Enabled state of all Edit menu options.
    '
    Menu_Edit_Click

End Sub

' The ToolPalette takes care of painting itself since it contains
' a single bitmap displaying the Tools.  However, the currently
' selected Tool needs to be Hightlighted during the Paint event.
'
Private Sub Pic_ToolPalette_Paint()
    
    ' Clear ToolPalette to background bitmap of Tools
    '
    Pic_ToolPalette.Cls

    ' HightLight current tool.  The Drawmode of the ToolPalette is
    ' set to INVERSE, so we simply draw a filled box over the current
    ' tool.  This inverts the tool.
    '
    Pic_ToolPalette.Line (0, CurrentTool * 32 + 1)-(31, CurrentTool * 32 + 31), , BF

End Sub

' Clicking the Undo Icon in the StatusArea is the same thing as
' selecting the Edit.Undo menu item.
'
Private Sub Pic_Undo_Click()

    Menu_EditSelection_Click MID_UNDO

End Sub

' In response to a File.New menu opertation, this routine is called,
' to prepare creating a new Icon from scratch.
'
Private Sub Prepare_For_New_Icon()
    
    ' Determine if current Icon has changed since it was saved last and
    ' and prompt user if it has been changed and not saved.
    '
    If Ok_To_Discard_Changes() Then
        '
        ' It is ok to prepare for new Icon.
        
        ' Terminate any Selection in progress and simulate a Select all,
        ' then perform an Edit.Delete operation to clear the Icon.
        '
        Selecting = False
        MovingSelection = False
        Scale_Region False, 0, 0, 32, 32, False
        Menu_EditSelection_Click MID_DELETE

        ' Intialize the IconInfo record to default values
        '
        ICONINFO(CurrentIcon).FileName = UNTITLED
        ICONINFO(CurrentIcon).Changed = False

        ' Reset both the Icon menu and the Editors Titlebar
        '
        Menu_IconsSelection(CurrentIcon).Caption = "&" + Format$(CurrentIcon + 1) + " - " + UNTITLED

        Caption = "IconWorks Editor: " + Format$(CurrentIcon + 1) + " - " + UNTITLED
    End If

End Sub

' Once a Selection has been made, or an Edit.Select All or Edit.Paste
' operation has been made, this routine is called to prepare for a
' potential movement of the Selection.  A copy of the actual Icons
' Mask and Image are made, and this is the Source from which the
' Selection is display whenever the Selection is in fact moved.
'
Private Sub Prepare_For_Possible_Move()
              
    Pic_Work.Picture = Pic_Icons(CurrentIcon).Image
    Pic_TempImage.Picture = Pic_Image.Image
    Pic_TempMask.Picture = Pic_Mask.Image

End Sub

' When a selection has been made and the selection is Square, it may
' be rotated Left or Right by 90 degrees.  Cannot Rotate a selection
' that is being moved.  This routine performs the rotation in repsonse
' to an Edit.Rotate Left/Right menu operation.
'
Private Sub Rotate_Selection(Index)
    
    ' Rotation is not instant, so set Cursor to HourGlass
    '
    Screen.MousePointer = HOURGLASS

    ' Caculate size of Region
    '
    NumPixels = X2SelectFrom - X1SelectFrom - 1

    ' Perform rotation.  The rotation is performed by reading one pixel at
    ' a time from the source and copying it to is new location in rotated
    ' either Left or Right 90 degrees.  The Point Method reads the Pixel
    ' and the Pset method is used to write the pixel.  Both the Icons
    ' Mask and Image must be rotated.
    '
    For Y = 0 To NumPixels
        For X = 0 To NumPixels
            If Index = MID_ROTATE_RIGHT Then
                '
                ' Rotate Right
                '
                Pic_Image.PSet (X2SelectFrom - Y - 1, Y1SelectFrom + X), Pic_TempImage.Point(X1SelectFrom + X, Y1SelectFrom + Y)
                Pic_Mask.PSet (X2SelectFrom - Y - 1, Y1SelectFrom + X), Pic_TempMask.Point(X1SelectFrom + X, Y1SelectFrom + Y)
            Else
                ' Rotate Left
                '
                Pic_Image.PSet (X1SelectFrom + Y, Y2SelectFrom - X - 1), Pic_TempImage.Point(X1SelectFrom + X, Y1SelectFrom + Y)
                Pic_Mask.PSet (X1SelectFrom + Y, Y2SelectFrom - X - 1), Pic_TempMask.Point(X1SelectFrom + X, Y1SelectFrom + Y)
            End If
        Next X
    Next Y

    ' Replace Icons Mask and Image with new Mask and Images.
    '
    Update_Icon Pic_Icons(CurrentIcon)
    
    Prepare_For_Possible_Move

    ' Redisplay selected region to show rotation.
    '
    Magnify_Icon X1SelectFrom, Y1SelectFrom, (X2SelectFrom), (Y2SelectFrom)

    ' Restore Cursor
    '
    Screen.MousePointer = DEFAULT

End Sub

' In response to a Color.Save colors to INI file menu operatioin, this
' routine is called to write the color values to IconWrks.INI.
' WritePrivateProfileString() will create the INI file if it does not
' exist, and will create the KeyName if it does not exist in the INI FILE.
' The Color values are packed into a single string, one row of colors at
' a time, and then written to the INI file.  The colors are formated
' into 8-byte fields separated by one space.  After this routine,
' IconWrks.INI will contain 3 key entries of:
'
'   Colors Row1=16777215 14737632 ...
'   Colors Row2=12632256 04210752 ...
'   Colors Row3=08421504 00000000 ...
'
Private Sub Save_Colors_To_INI_File()
    
    For i = 0 To 2
        ColorString = ""
        For C = 0 To 15
            ColorString = ColorString + Format$(Colors(i * 16 + C), "00000000 ")
        Next C
        X = WritePrivateProfileString(APP_NAME, KEY_COLOR + Format$(i + 1), ColorString, INI_FILENAME)
    Next i

End Sub

' In Response to a File.Save or File.Save As menu operation, this routine
' is called to either immediately save the Icon or invoke the SaveFileDlg
' to get a Filename to save the Icon to.  A File.Save As operation always
' invokes the Dialog, but a File.Save will only invoke the Dialog if the
' current icon does not yet have a name, denoted by a name of UNTITLED.
'
Private Sub Save_Icon(Index)
    On Error Resume Next
    Dim LongName As String
    Dim Title As String
    Dim CDPath As String

    If (Index = MID_SAVE_AS) Or (ICONINFO(CurrentIcon).FileName = UNTITLED) Then
        ' Set default filter.
        Editor.CMDialog1.Filter = "Icon Files (*.ico)|*.ico"
        ' Display Save As dialog.
        Editor.CMDialog1.FilterIndex = 1
        Editor.CMDialog1.DefaultExt = "ICO"
        If Right(ICONINFO(CurrentIcon).FullPath, 1) <> "\" Then
          Editor.CMDialog1.FileName = ICONINFO(CurrentIcon).FullPath & "\" & ICONINFO(CurrentIcon).FileName
        End If
        ' If file has not been saved yet, set common dialog's
        ' default file name to *.ico
        If ICONINFO(CurrentIcon).FileName = "[Untitled]" Then
          Editor.CMDialog1.FileName = "*.ico"
        End If
        
        Editor.CMDialog1.Action = 2
        ' Exit if user cancels the dialog.
        If Err = 32755 Then Exit Sub
        Title = Editor.CMDialog1.FileTitle
        LongName = Editor.CMDialog1.FileName
        ICONINFO(CurrentIcon).FullPath = Left(LongName, Len(LongName) - Len(Title))
        ICONINFO(CurrentIcon).FileName = Editor.CMDialog1.FileTitle
        ' Adjust for root directory.
        If Len(ICONINFO(CurrentIcon).FullPath) = 3 Then
            CDPath = ICONINFO(CurrentIcon).FullPath
        Else
            CDPath = Left(LongName, Len(LongName) - Len(Title) - 1)
        End If
        ' Change directories for the editor and the viewer.
        ChDir CDPath
        Viewer.Dir_DirectoryList.Path = CDPath
    End If
        Write_Icon_To_File (ICONINFO(CurrentIcon).FullPath), ICONINFO(CurrentIcon).FileName
End Sub

' When the Editor is terminated, this routine is called to save all
' configurable items of the Editor to the IconWrks.INI file.  Size
' and Location of the Editor are not saved.
'
' WritePrivateProfileString() will create the INI file and/or any
' KeyName that does not exist.
'
Private Sub Save_Settings_To_INI_File()
  
    R = WritePrivateProfileString(APP_NAME, KEY_GRID, Format$(Abs(Menu_ViewSelection(MID_GRID).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_LINE_GRID, Format$(Abs(Menu_ViewSelection(MID_LINE_GRID).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_DOTTED_GRID, Format$(Abs(Menu_ViewSelection(MID_DOTTED_GRID).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_STATUS_BAR, Format$(Abs(Menu_ViewSelection(MID_STATUS_BAR).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_TOOL_PALETTE, Format$(Abs(Menu_ViewSelection(MID_TOOL_PALETTE).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_COLOR_PALETTE, Format$(Abs(Menu_ViewSelection(MID_COLOR_PALETTE).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_ZOOM_SCROLLBAR, Format$(Abs(Menu_ViewSelection(MID_ZOOM_SCROLLBAR).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_FOCUS, Format$(Abs(Menu_ViewSelection(MID_FOCUS).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_BORDER, Format$(Abs(Menu_ViewSelection(MID_BORDER).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_POPUP, Format$(Abs(Menu_ViewSelection(MID_POPUP_MENU).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_DEFAULT, Format$(Abs(Menu_ColorSelection(MID_DEFAULT_COLORS).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_SOLID, Format$(Abs(Menu_ColorSelection(MID_SOLID_COLORS).Checked)), INI_FILENAME)
    R = WritePrivateProfileString(APP_NAME, KEY_INI, Format$(Abs(Menu_ColorSelection(MID_INI_COLORS).Checked)), INI_FILENAME)

End Sub

' Scrolling the Zoom scrollbar Zooms in or Zooms out on the
' selected Icon, and maintains it centered within the available
' editing space.
'
Private Sub Scrl_Zoom_Change()
    
    ' Check if new value is multiple of 32 pixels.  Magnified Icon must
    ' always be increased or decreased my muliples of 32 pixels so the
    ' Win API routine StretchBlt() can be used to magnify the Icon, and
    ' the result will be an exact enlargement of the original.
    '
    If (Scrl_Zoom.Value - 2) Mod Pic_Icons(0).Width Then
        '
        ' Scrollbar value is not a multiple of 32, so adjust it so it is.
        '
        Scrl_Zoom.Value = Scrl_Zoom.Value - ((Scrl_Zoom.Value - 2) Mod Pic_Icons(0).Width)
    Else
        ' Scrollbar value is a multiple of 32 pixels.
        '
        ' Hide Edit area while resizing and repainting.  It prevents a
        ' rather annoying visual affect.
        '
        Pic_Edit.Visible = False

        ' Calculate the new size and position values for the Editing area
        ' and then reposition and resize both the Editing area and the
        ' temporary Editing area which is used when magnifying the Icon.
        '
        MaxLeft = Scrl_Zoom.Left - (Menu_ViewSelection(MID_ZOOM_SCROLLBAR).Checked * Scrl_Zoom.Width)
        NewLeft = MaxLeft + (ScaleWidth + Menu_ViewSelection(MID_STATUS_BAR).Checked * Pic_StatusArea.Width - MaxLeft - Scrl_Zoom.Value) \ 2
        NewTop = (ScaleHeight + Menu_ViewSelection(MID_COLOR_PALETTE).Checked * Pic_ColorPalette.Height - Scrl_Zoom.Value) \ 2
        Pic_Edit.Move NewLeft, NewTop, Scrl_Zoom.Value, Scrl_Zoom.Value
        Pic_EditTemp.Move 0, 0, Scrl_Zoom.Value, Scrl_Zoom.Value

        ' To allow a selection to be maintained if the Zoom scrollbar is scrolled
        ' while a selection is in progress, the current selection must be scaled
        ' down to the scale of the actual Icon, the new PixelSize for the Editing
        ' area must be calculated, and then the selection is Scaled back to the
        ' new scale of the Editing area.
        '
        If Selecting Then Scale_Region True, X1Region, Y1Region, X2Region, Y2Region, False
        PixelSize = Pic_Edit.ScaleWidth \ 32
        If Selecting Then Scale_Region False, X1Region, Y1Region, X2Region, Y2Region, False

        ' Recreate grid if selected and redisplay the Editing area.
        '
        If Menu_ViewSelection(MID_GRID).Checked Then Create_Grid
        Pic_Edit.Visible = True

        ' Set enable state of Zoom in/out View menu items based on new value
        ' of the Zoom Scrollbar.
        '
        Menu_ViewSelection(MID_ZOOM_IN).Enabled = Scrl_Zoom.Value < Scrl_Zoom.Max
        Menu_ViewSelection(MID_ZOOM_OUT).Enabled = Scrl_Zoom.Value > Scrl_Zoom.Min
    End If

End Sub

' In response to Double Clicking the Select tool, or an Edit.Select All
' menu operation, this routine is called to perform a selection on the
' entire Icon, which is the same as if the Entire icon was selected using
' the mouse.
'
Private Sub Select_Entire_Icon()
    
    ' Make the select to the selected tool.  This is neccessary since
    ' this routine could have been called due to a menu selection, and
    ' the current tool may not be the Select Tool.
    '
    Pic_ToolPalette_MouseDown LEFT_BUTTON, 0, 0, TID_SELECT * 32

    ' Make all preparations for a selection just as if it was done
    ' using the mouse.
    '
    Prepare_For_Possible_Move
    Selecting = True
    
    Opaque = True
    Scale_Region False, 0, 0, 32, 32, False
    Scale_Region True, X1SelectFrom, Y1SelectFrom, X2SelectFrom, Y2SelectFrom, False
    Draw_Selection_Rectangle

End Sub

' There are two type of grids to choose from when the Grid selection
' is active. Line and Dotted grid.  This routine is called in response
' to a View.Line Grid or View.Dotted Grid menu operation, to set and
' display the selected grid type.  The actual Grid is created in the
' Create_Grid procedure.
'
Private Sub Select_Grid_Type(Index)
    
    ' Check if selected Grid is already the currently selected grid type,
    ' since there is no need to redisplay it.
    '
    If Not Menu_ViewSelection(Index).Checked Then
        '
        ' Toggle Checks on Grid type menu items.
        '
        Menu_ViewSelection(MID_LINE_GRID).Checked = Not Menu_ViewSelection(MID_LINE_GRID).Checked
        Menu_ViewSelection(MID_DOTTED_GRID).Checked = Not Menu_ViewSelection(MID_DOTTED_GRID).Checked

        ' Create the new Grid and redisplay Icon with new Grid.
        '
        Create_Grid
        Magnify_Icon 0, 0, 31, 31
    End If

End Sub

Private Sub Select_Tool_MouseUp(X1, Y1, X2, Y2)
            
    ' Erase old Selection rectangle
    '
    Draw_Selection_Rectangle

    ' Adjust selection rectangle so it is completely within the
    ' editing area.
    '
    If X1Region > X2Region Then Swap_Values X1Region, X2Region
    If Y1Region > Y2Region Then Swap_Values Y1Region, Y2Region
    If X1Region < 0 Then X1Region = 0
    If Y1Region < 0 Then Y1Region = 0
    MaxSize = PixelSize * 32
    If X2Region > MaxSize Then X2Region = MaxSize
    If Y2Region > MaxSize Then Y2Region = MaxSize

    If MovingSelection Then
        '
        ' The selection is being moved, so all that needs to be done is
        ' to update the values containing the last location of the
        ' selection, which are used when the selection is moved.
        '
        X1 = LastX1
        Y1 = LastY1
        X2 = LastX2
        Y2 = LastY2

    ElseIf (X1Region <> X2Region) And (Y1Region <> Y2Region) Then
        '
        ' A new selection has been made
        '
        Prepare_For_Possible_Move
        Selecting = True
        Opaque = True

        ' Determine if selection can be moved.  The selection cannot be moved
        ' if it consists of the entire icon.
        '
       

        Draw_Selection_Rectangle
        Scale_Region True, X1SelectFrom, Y1SelectFrom, X2SelectFrom, Y2SelectFrom, True
    End If
    Scale_Region True, LastX1, LastY1, LastX2, LastY2, True

End Sub

' A move event using the Select tool, can do one of three things,
' size the selection rectangle to make the selection, move the
' current selection, or copy the current selection to another
' location.  This routine either draws the selection rectangle
' as the selection is being made or moves/copies the current
' selection.
'
Private Sub Select_Tool_Move(Shift As Integer, X As Single, Y As Single)

    ' Erase old selection rectangle
    '
    Draw_Selection_Rectangle

    If MovingSelection Then
        '
        ' Calculate new location of Selection rectangle.
        '
        XLastMove = XMoveSelection
        YLastMove = YMoveSelection
        Adjust_X_and_Y X, Y, XMoveSelection, YMoveSelection, TO_EDGE
        X1Region = X1Region + (XMoveSelection - XLastMove)
        X2Region = X2Region + (XMoveSelection - XLastMove)
        Y1Region = Y1Region + (YMoveSelection - YLastMove)
        Y2Region = Y2Region + (YMoveSelection - YLastMove)

        ' Prevent Selection from being moved of the left or right
        ' of the editing area
        '
        If (X1Region < 0) Or (X2Region > (32 * PixelSize)) Then
            X1Region = X1Region - (XMoveSelection - XLastMove)
            X2Region = X2Region - (XMoveSelection - XLastMove)
        End If
        
        ' Prevent Selection from being moved of the top or bottom
        ' of the editing area
        '
        If (Y1Region < 0) Or Y2Region > (32 * PixelSize) Then
            Y1Region = Y1Region - (YMoveSelection - YLastMove)
            Y2Region = Y2Region - (YMoveSelection - YLastMove)
        End If
    Else
        ' A selection is being made, so adjust the mouse coordinates
        ' to the Edge of the pixels so the rectangle is completely
        ' around the selction.
        '
        Adjust_X_and_Y X, Y, X2Region, Y2Region, TO_EDGE

        ' Make the Selection perfectly square if the Shift key is down
        '
        If Shift And SHIFT_MASK Then Make_Region_Square
    End If

    ' Draw the new selection recangle
    '
    Draw_Selection_Rectangle

End Sub

' When the MousePointer is within the Editing area, the MousePointer
' is changed depending on the Tool being used.  This routine is called
' from the MouseMove event for the Editing area (Pic_Edit) to change
' the MousePointer.
'
Private Sub Set_MousePointer(X As Single, Y As Single)
    
    If Selecting And Moveable And (X >= X1Region) And (X <= X2Region) And (Y >= Y1Region) And (Y <= Y2Region) Then
        '
        ' A selection is in progress and the selection is moveable,
        ' and the MousePointer is within the the selection, so the
        ' MousePointer is changed to a SIZE pointer to indacte
        ' that the Selection can be moved.
        '
        Pic_Edit.MousePointer = SIZE
    ElseIf (CurrentTool = TID_PAINT) Or (CurrentTool = TID_FILL) Then
        Pic_Edit.MousePointer = UP_ARROW
    Else
        Pic_Edit.MousePointer = CROSSHAIR
    End If

End Sub

' Since some of the tools can only use Solid colors and not Dithered
' colors, it is useful to have the option of displayin only the solid
' colors within the ColorPalette.  This routine changes the Color
' array to contain only Solid color values then redisplays the
' ColorPalette.  There are still 3 rows of colors, but each color
' cell in the same column are set to the same solid color.
'
Private Sub Show_Solid_Colors_Only()
    
    ' Set Color array to all solid colors returned from QBColors()
    '
    For i = 0 To 15
        Colors(i) = QBColor(i)
        Colors(i + 16) = QBColor(i)
        Colors(i + 32) = QBColor(i)
    Next

    ' Redisplay ColorPalettes.
    '
    Display_Color_Palette Pic_ColorPalette
    If ColorPaletteLoaded Then Display_Color_Palette ColorPalette.Pic_ColorPalette

End Sub

' This routine is called in response to an Edit.Undo menu operation.
' The Icon displayed in the Undo Icon within the StatusArea is
' Copied to the Selected Icon.
'
Private Sub UnDo_Edits()
    
    ' Terminate any Select or Move opteration
    '
    Selecting = False
    MovingSelection = False

    ' Get Image and Mask of Icon displayed in Undo Icon in StatusArea
    '
    Extract_Image_And_Mask Pic_Undo

    ' Copy Undo Icons Mask and Image to selected icon
    '
    Update_Icon Pic_Icons(CurrentIcon)

    ' Redisplay entire icon
    '
    Magnify_Icon 0, 0, 31, 31

End Sub

' When using either the Circle or Filled Circle tool, this routine is called
' to verify that the resulting circle from the current Mouse coordinates is
' a valid circle, since in some cases the result would be a vertical or
' horizontal line which would generate and "Division by Zero" if not checked
' for.  If the Circle is valid, this routine calculates and sets the Global
' values of XCenter, YCenter, and Raduis for new Circle.
'
Private Function Valid_Circle(X1, Y1, X2, Y2)
          
    Xleg = Abs(X2 - X1)
    If Xleg <> 0 Then
        Yleg = Abs(Y2 - Y1)
        Aspect = Abs(Yleg / Xleg)
        If Xleg > Yleg Then Radius = Xleg / 2 Else Radius = Yleg / 2
        XCenter = X1 + (X2 - X1) / 2
        YCenter = Y1 + (Y2 - Y1) / 2
    End If

    ' Return value of XLeg at Function value, since if XLeg is non-zero,
    ' (True) it is a valid circle, and if XLeg is zero (False) it is
    ' an Invalid circle.
    '
    Valid_Circle = Xleg

End Function

