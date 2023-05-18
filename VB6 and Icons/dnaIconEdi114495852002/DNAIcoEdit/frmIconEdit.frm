VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIcoEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNA Icon Editor"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmIconEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00AB8F8D&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   7455
      TabIndex        =   12
      Top             =   5985
      Width           =   7455
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " DNA Icon Editor"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   50
         Width           =   7455
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00AB8F8D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   1155
      Width           =   788
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   8
         ToolTipText     =   "Colore corrente"
         Top             =   2670
         Width           =   465
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   15
         Left            =   30
         TabIndex        =   36
         Top             =   2670
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   14
         Left            =   510
         TabIndex        =   35
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   13
         Left            =   270
         TabIndex        =   34
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   30
         TabIndex        =   33
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   510
         TabIndex        =   32
         Top             =   2130
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   270
         TabIndex        =   31
         Top             =   2130
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   9
         Left            =   30
         TabIndex        =   30
         Top             =   2130
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   8
         Left            =   510
         TabIndex        =   29
         Top             =   1860
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   270
         TabIndex        =   28
         Top             =   1860
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   30
         TabIndex        =   27
         Top             =   1860
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   510
         TabIndex        =   26
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   270
         TabIndex        =   25
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   24
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   510
         TabIndex        =   23
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   22
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Top             =   1320
         Width           =   225
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   460
         Left            =   120
         Top             =   120
         Width           =   460
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   50
         X2              =   720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "32 * 32"
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
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Format"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   45
         X2              =   715
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H00AB8F8D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   5640
      ScaleHeight     =   4815
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   1155
      Width           =   1822
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   43
         Top             =   3180
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AB8F8D&
         Caption         =   "Manual size"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   44
         Top             =   3440
         Width           =   375
      End
      Begin VB.PictureBox ctlXPButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   15
         ToolTipText     =   "Set BackGround Default Color"
         Top             =   360
         Width           =   230
      End
      Begin VB.PictureBox ctlXPButton3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1220
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   16
         ToolTipText     =   "Change BackGround Color"
         Top             =   360
         Width           =   225
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   980
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   42
         Top             =   360
         Width           =   225
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   39
         ToolTipText     =   "Colore trasparente"
         Top             =   2040
         Width           =   1080
      End
      Begin VB.CheckBox chkMask 
         Appearance      =   0  'Flat
         BackColor       =   &H00AB8F8D&
         Caption         =   "Apply Mask"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2320
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00AB8F8D&
         Caption         =   "32*32 - 16 col"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AB8F8D&
         Caption         =   "16*16 - 16 col"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         MaskColor       =   &H80000000&
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtBack 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Text            =   "16777215"
         Top             =   360
         Width           =   1245
      End
      Begin VB.PictureBox btnApplyMask 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   41
         ToolTipText     =   "Set Color Mask"
         Top             =   2040
         Width           =   225
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   1320
         Picture         =   "frmIconEdit.frx":000C
         Top             =   4440
         Width           =   360
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   3440
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3180
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         Top             =   3000
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00AB8F8D&
         Caption         =   "Trasparent Color"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         Top             =   1920
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00AB8F8D&
         Caption         =   "New Format"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "BackGround Color"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   795
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageIndex      =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paint"
            ImageIndex      =   5
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Change"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Down"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fOriz"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fVert"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rotate"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   810
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   1155
      Width           =   4815
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture12 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7455
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4680
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":062E
               Key             =   "Open"
               Object.Tag             =   "Open"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":0D02
               Key             =   "Save"
               Object.Tag             =   "Save"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1154
               Key             =   "New"
               Object.Tag             =   "New"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":12AE
               Key             =   "Fill"
               Object.Tag             =   "Fill"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1408
               Key             =   "Paint"
               Object.Tag             =   "Paint"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1562
               Key             =   "Down"
               Object.Tag             =   "Down"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":16BC
               Key             =   "Up"
               Object.Tag             =   "Up"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1816
               Key             =   "Right"
               Object.Tag             =   "Right"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1970
               Key             =   "Left"
               Object.Tag             =   "Left"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1ACA
               Key             =   "Refresh"
               Object.Tag             =   "Refresh"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1C24
               Key             =   "Exit"
               Object.Tag             =   "Exit"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIconEdit.frx":1ED8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3480
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Developer's CodeBook 2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   10440
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DNA Icon Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   525
         TabIndex        =   3
         Top             =   400
         Width           =   5295
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   6480
         Picture         =   "frmIconEdit.frx":2032
         Stretch         =   -1  'True
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   576
      Y1              =   75
      Y2              =   75
   End
End
Attribute VB_Name = "frmIcoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Long, b As Long, c As Long, d As Long, e As Long
Private W As Long, Lin As Long, UNL As Long
Private lwFormat As Boolean
Private valStep As Integer, valAdd As Integer, upPoint As Integer
Private colorGrid As Long

Private Sub btnApplyMask_Click()
' Applicazione maschera
  Picture4.BackColor = Val(txtBack)
End Sub

Private Sub Check1_Click()
' Lock seleziona manuale delle dimensioni dell'icona
  If Check1.Value = 0 Then
     Text2.Locked = True
     Text1.Locked = True
  Else
     Text2.Locked = False
     Text1.Locked = False
  End If
End Sub

Private Sub Form_Load()
' Inizializzazione parametri e caricamento
  'Me.Icon = frmCodeLib.Icon
  valStep = 32
  valAdd = 10
  lwFormat = False
  'Image2.Picture = LoadResImage("LOGO", "LOGO")
  Picture1.BackColor = &HC0C0C0    'vbWhite
  Picture3.BackColor = &HC0C0C0    'vbWhite
  Picture3.Height = 321
  Picture3.Width = 321
  ImageList1.MaskColor = &HC0C0C0  'vbwhite
  For e = 0 To 15
      lblColor(e).BackColor = QBColor(e)
  Next e
  Picture10 = Picture1.Image
  txtBack = "16777215"
  Option1.Value = False
  upPoint = 9
  LoadGrid valStep
End Sub

Private Sub Picture2_Change()
' posizionamento della picture icona
  If Picture2.Width < 300 Then
     Picture1.Top = Shape2.Top + 120
     Picture1.Left = Shape2.Left + 120
     Label3.Caption = "16 * 16"
     Text1 = "16": Text2 = "16"
     lwFormat = True
     Option1.Value = True
  Else
     Picture1.Top = Shape2.Top
     Picture1.Left = Shape2.Left
     Label3.Caption = "32 * 32"
     Text1 = "32": Text2 = "32"
     lwFormat = False
     Option2.Value = True
  End If
End Sub

Private Sub Text1_Change()
' max valore 64
  If Val(Text1) > 64 Then Text1 = 64
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
' accetta solo numeri
  Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub Text2_Change()
' max valore 64
  If Val(Text2) > 64 Then Text2 = 64
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
' accetta solo numeri
  Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' Selezione utente
  Select Case Button.Key
    Case "Open":   open_Ico
    Case "New":    new_Ico
    Case "Save":   save_Ico
    Case "Fill":   Fill
    Case "Paint":  Toolbar1.Buttons("Change").Value = tbrUnpressed
      Toolbar1.Buttons("Paint").Value = tbrPressed
      Picture3.MousePointer = 99
    Case "Change": Toolbar1.Buttons("Paint").Value = tbrUnpressed
      Toolbar1.Buttons("Change").Value = tbrPressed
      Picture3.MousePointer = 10
    Case "Up":     movePict "up"
    Case "Down":   movePict "down"
    Case "Left":   movePict "left"
    Case "Right":  movePict "right"
    Case "fOriz":  Picture1.PaintPicture Picture10, Picture10.ScaleWidth - 1, 0, -Picture10.ScaleWidth, Picture10.ScaleHeight
        PaintDown
        LoadGrid valStep
        Picture10 = Picture1.Image
    Case "fVert":  Picture1.PaintPicture Picture10, 0, Picture10.ScaleHeight - 1, Picture10.ScaleWidth, -Picture10.ScaleHeight
        PaintDown
        LoadGrid valStep
        Picture10 = Picture1.Image
    Case "Rotate": Rotat_Click
    Case "Exit":   Unload Me
  End Select
End Sub

Private Sub new_Ico()
' Crea nuova icona
  a = &HC0C0C0    'vbWhite
  btnFillPic
  W = 0
  Picture1.BackColor = &HC0C0C0    'vbWhite '&HFFC0C0
  If Option1.Value Then
     Picture1.Width = 240
     Picture1.Height = 240
     Picture1.Top = Shape2.Top + 120
     Picture1.Left = Shape2.Left + 120
     Label3.Caption = "16 * 16"
     Text1 = "16": Text2 = "16"
     valStep = 16
  Else
     Picture1.Width = 480
     Picture1.Height = 480
     Picture1.Top = Shape2.Top
     Picture1.Left = Shape2.Left
     Label3.Caption = "32 * 32"
     Text1 = "32": Text2 = "32"
     valStep = 32
  End If
  PaintDown
  UNL = 1
  Picture10 = Picture1.Image
  btnFillPic
End Sub

Private Sub open_Ico()
' Caricamento Icona
'  Dim sOpen As SelectedFile
'  Dim RetString As String
'  FileDialog.Filter = "Icons Files (*.ico)" & Chr$(0) & "*.ico"
'  FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
'  FileDialog.DialogTitle = "Caricamento File Icona"
'  FileDialog.InitDir = LastDir
'  sOpen = ShowOpen(Me.hWnd)
'  If err.Number <> 32755 And sOpen.bCanceled = False Then
'     RetString = stripSpace(FileDialog.file)
'     If Trim(RetString) = "" Then Exit Sub
'     Label6.Caption = " " & RetString & "... " & Len(RetString) & " byte"
'  Else
'     Exit Sub
'  End If
  Dim fileN As String
  CommonDialog1.CancelError = True
  On Error GoTo EX
  CommonDialog1.FileName = ""
  CommonDialog1.Flags = cdlOFNFileMustExist
  CommonDialog1.Filter = "Icons (*.ico)|*.ico"
  CommonDialog1.ShowOpen
  fileN = CommonDialog1.FileName
  Label6.Caption = " " & fileN & "... " & Len(fileN) & " byte"
  Fill
  W = 0
  Toolbar1.Buttons("Save").Enabled = True
  Picture1.BackColor = &HC0C0C0    'vbWhite '&HFFC0C0
  Picture1 = LoadPicture(fileN)
  Picture2 = LoadPicture(fileN)
  resizePicture
  PaintDown
  UNL = 1
  Picture1.Cls
  Picture10 = Picture1.Image
  If lwFormat Then
     LoadGrid 16:  Option1.Value = True
  Else
     LoadGrid 32:  Option2.Value = True
  End If
  Exit Sub
EX:
End Sub

Private Sub PaintDown()
' scelta del colore della grid (chiaro/scuro)
  Static Punt As Long
  Punt = Picture3.Point(0, 0)
  Picture3.PaintPicture Picture1.Image, 0, 0, 321, 321
  If Punt = &HFFC0FF Then colorGrid = &HFFC0FF Else colorGrid = &H80000001
End Sub

Private Sub Fill()
' Modifica del colore di sfondo
  UNL = 1:  W = 0
  Picture1.Line (0, 0)-(31, 31), a, BF
  Picture3.BackColor = a
  LoadGrid valStep, colorGrid
  Picture10 = Picture1.Image
End Sub

Private Sub movePict(moveType As String)
' Modifica delle posizioni e dell'orientamento della picture
  Select Case moveType
    Case "left":  RtoL -1, 0, 31, 0
    Case "right": RtoL 1, 0, -31, 0
    Case "up":    RtoL 0, -1, 0, 31
    Case "down":  RtoL 0, 1, 0, -31
  End Select
End Sub

Private Sub ctlXPButton1_Click()
' Assegnazione colore de sfondo di default
  txtBack.Text = "16777215"
  Picture5.BackColor = Val(txtBack)
  btnFillPic
End Sub

Private Sub btnFillPic()
' Pulizia della picture di editing
  a = Picture5.BackColor
  Fill
End Sub

Private Sub ctlXPButton3_Click()
' Pulizia della picture di editing
  Picture3.BackColor = Val(txtBack)
  picBack.BackColor = Val(txtBack)
  btnFillPic
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Unload e salvataggio icone
  If UNL <> 0 Then
     Dim msG As String, style, Resp
     msG = "Vuoi salvare l'icona ?"
     style = vbYesNoCancel + vbExclamation
     Resp = MsgBox(msG, style)
     If Resp = vbYes Then save_Ico: If UNL <> 0 Then Cancel = True
     If Resp = vbNo Then Cancel = False
     If Resp = vbCancel Then Cancel = True
  End If
End Sub

Private Sub lblcolor_Click(Index As Integer)
' Aggiornamento colore di sfondo
  a = QBColor(Index)
  Picture5.BackColor = a
End Sub

Private Sub lblColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Aggiornamento label visualizzazione codice colori
  txtBack = Hex2Long(Hex(lblColor(Index).BackColor))
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Impostazione del colore e disegno sulla picture
  b = Picture3.Point(X, Y)
  If b = &H80000001 Or b = &HFFC0FF Then Exit Sub
  Toolbar1.Buttons("Save").Enabled = True
  If Button = vbRightButton Then
     a = Picture3.Point(X, Y)
     Select Case a
         Case QBColor(0): lblcolor_Click (0): Exit Sub
         Case QBColor(1): lblcolor_Click (1): Exit Sub
         Case QBColor(2): lblcolor_Click (2): Exit Sub
         Case QBColor(3): lblcolor_Click (3): Exit Sub
         Case QBColor(4): lblcolor_Click (4): Exit Sub
         Case QBColor(5): lblcolor_Click (5): Exit Sub
         Case QBColor(6): lblcolor_Click (6): Exit Sub
         Case QBColor(7): lblcolor_Click (7): Exit Sub
         Case QBColor(8): lblcolor_Click (8): Exit Sub
         Case QBColor(9): lblcolor_Click (9): Exit Sub
         Case QBColor(10): lblcolor_Click (10): Exit Sub
         Case QBColor(11): lblcolor_Click (11): Exit Sub
         Case QBColor(12): lblcolor_Click (12): Exit Sub
         Case QBColor(13): lblcolor_Click (13): Exit Sub
         Case QBColor(14): lblcolor_Click (14): Exit Sub
         Case QBColor(15): lblcolor_Click (15): Exit Sub
         'Case vbWhite: btnFillPic: Exit Sub
         Case &HC0C0C0: btnFillPic: Exit Sub
: btnFillPic: Exit Sub
     End Select
  End If
  If Button <> vbLeftButton Then Exit Sub
  UNL = 1:  W = 0
  If Picture3.MousePointer = 10 Then
     b = Picture3.Point(X, Y)
     If b = a Then Exit Sub
     For j = 0 To Picture1.ScaleWidth - 1
         For p = 0 To Picture1.ScaleHeight - 1
             c = Picture1.Point(j, p)
             If c = b Then Picture1.PSet (j, p), a
         Next p
     Next j
     PaintDown
     'Fill
     LoadGrid valStep
     Exit Sub
  End If
  Lin = 1
  X1 = 0:  Y1 = 0
  For j = 0 To 31
      For p = 0 To 31
          If X < X1 + valAdd And X > X1 And Y < Y1 + valAdd And Y > Y1 Then
             Picture3.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), a, BF
             Picture1.PSet (X1 \ valAdd, Y1 \ valAdd), a
          End If
          X1 = X1 + valAdd
          If X1 = 320 Then
             X1 = 0
             Y1 = Y1 + valAdd
          End If
      Next p
  Next j
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Disegno continuo
  If Lin = 0 Then Exit Sub
  X1 = 0:  Y1 = 0
  For j = 0 To 31
      For p = 0 To 31
          If X < X1 + valAdd And X > X1 And Y < Y1 + valAdd And Y > Y1 Then
             Picture3.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), a, BF
             Picture1.PSet (X1 \ valAdd, Y1 \ valAdd), a
          End If
          X1 = X1 + valAdd
          If X1 = 320 Then
             X1 = 0
             Y1 = Y1 + valAdd
          End If
      Next p
  Next j
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Assegnazione Picture da modificare
  Lin = 0
  If Button = vbLeftButton Then
     Picture10 = Picture1.Image
  End If
End Sub

Private Sub Rotat_Click()
' Effetto rotazione Picture
  For j = 0 To Picture1.ScaleWidth - 1
      For p = 0 To Picture1.ScaleHeight - 1
          Picture1.PSet (j, p), Picture10.Point(p, j)
      Next p
  Next j
  Picture10 = Picture1.Image
  Picture1.PaintPicture Picture10, Picture10.ScaleWidth - 1, 0, -Picture10.ScaleWidth, Picture10.ScaleHeight
  PaintDown
  LoadGrid valStep
  Picture10 = Picture1.Image
End Sub

Private Sub RtoL(Pr1 As Long, Pr2 As Long, Pr3 As Long, Pr4 As Long)
' Routines di rotazione dell'immagine
  On Error GoTo err
  W = 0
  Picture1.PaintPicture Picture10, Pr1, Pr2
  Picture1.PaintPicture Picture10, Pr3, Pr4
  PaintDown
  LoadGrid valStep
  Picture10 = Picture1.Image
err:
End Sub

Private Sub LoadGrid(gridFormat As Integer, Optional backCk As Long)
' Caricamento Griglia
  Dim varStep As Integer
  If backCk = 0 Then backCk = &H80000001
  If gridFormat = 16 Then varStep = 20 Else varStep = 10
  For F = 0 To Picture3.ScaleHeight Step varStep
      Picture3.Line (0, F)-(Picture3.ScaleWidth, F), backCk
  Next F
  For F = 0 To Picture3.ScaleWidth Step varStep
      Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), backCk
  Next F
  If gridFormat = 16 Then
     upPoint = 19
     valAdd = 20
     valStep = 16
  Else
     upPoint = 9
     valAdd = 10
     valStep = 32
  End If
End Sub

Public Sub Hex2RGB(strHexColor As String, R As Byte, G As Byte, b As Byte)
' Funzioni di codifica Hex - Long
  Dim hexColor As String
  Dim i As Byte
  On Error Resume Next
  strHexColor = Right((strHexColor), 6)
  For i = 1 To (6 - Len(strHexColor))
      hexColor = hexColor & "0"
  Next
  hexColor = hexColor & strHexColor
  R = CByte("&H" & Right$(hexColor, 2))
  G = CByte("&H" & Mid$(hexColor, 3, 2))
  b = CByte("&H" & Left$(hexColor, 2))
End Sub

Public Function Hex2Long(strHexColor As String) As Long
' Funzioni di codifica Hex - Long
  Dim R As Byte
  Dim G As Byte
  Dim b As Byte
  On Error Resume Next
  Hex2RGB strHexColor, R, G, b
  Hex2Long = RGB(R, G, b)
End Function

Private Sub save_Ico()
' Salvataggio Icona
  'Dim sSave As SelectedFile
  Dim RecPercorso As String
  Dim fPic As Integer
  fPic = IIf(lwFormat, 16, 32)
  If Check1.Value Then fPic = 0
  On Error GoTo err
  'sSave = ShowSave(Me.hWnd, , "icoFile" & ".ico")
  'If err.Number <> 32755 And sSave.bCanceled = False Then
  '   RecPercorso = stripSave(FileDialog.file, "ico")
  CommonDialog1.CancelError = True
  CommonDialog1.FileName = ""
  CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
  CommonDialog1.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp"
  CommonDialog1.ShowSave
  If CommonDialog1.FilterIndex = 1 Then
     RecPercorso = CommonDialog1.FileName
     If SaveICO(RecPercorso, Picture1.Image, ImageList1, picBack.BackColor, fPic, chkMask.Value, Picture4.BackColor) = False Then
        Label6.Caption = "Copia annullata"
     Else
        Label6.Caption = "Nuova Icona... " & RecPercorso
     End If
  End If
  UNL = 0
err:
End Sub

Public Function SaveICO(ByVal FilePath As String, ByVal tmpPicture As StdPicture, ByRef ImageList1 As ImageList, ByVal BackColor As Long, ByVal formatPic As Integer, Optional ByVal UseMask As Boolean = False, Optional ByVal MaskColor As Long = 0, Optional ByVal PromptToOverwrite As Boolean = True) As Boolean
  On Error GoTo ErrorTrap
  Dim li As ListImage
  Dim ThePic As StdPicture
  Dim MyAnswer As VbMsgBoxResult
' Verifica la Path ed il formato del file
  If FilePath = "" Then
    MsgBox "Path non valida.", vbOKOnly + vbExclamation, "  Errore"
    Exit Function
  ElseIf Dir(FilePath) <> "" Then
    'If PromptToOverwrite = True Then
    '   MyAnswer = MsgBox(FilePath & Chr(13) & "Questo file e' gia presente." & Chr(13) & Chr(13) & "Vuoi sostituirlo ?", vbYesNo + vbExclamation, "  File e' gia presente")
    'Else
       MyAnswer = vbYes
    'End If
    'If MyAnswer <> vbYes Then
    '   Exit Function
    'End If
  ElseIf tmpPicture = 0 Then
    MsgBox "Impossibile salvare l'icona in questo formato,", vbOKOnly + vbExclamation, "  Errore"
    Exit Function
  End If
' Se la picture e' gia un'Icona, la salva ed esce
  If tmpPicture.Type = vbPicTypeIcon Then
     SavePicture tmpPicture, FilePath
     GoTo Cleanup
  End If
' Setup ImageList control
  ImageList1.ListImages.Clear
  ImageList1.ImageHeight = IIf(formatPic <> 0, formatPic, Val(Text2))
  ImageList1.ImageWidth = IIf(formatPic <> 0, formatPic, Val(Text1))
  ImageList1.BackColor = BackColor
  ImageList1.MaskColor = MaskColor
  ImageList1.UseMaskColor = UseMask
' Set picture
  Set li = ImageList1.ListImages.Add(, , tmpPicture)
' Estrae la picture come file icona
  Set ThePic = li.ExtractIcon
  If ThePic = 0 Then
     Exit Function
  End If
' Salva la picture
  SavePicture ThePic, FilePath
  SaveICO = True
Cleanup:
' Clean Up
  Set ThePic = Nothing
  Set li = Nothing
  Exit Function
ErrorTrap:
  If err.Number = 0 Then
    Resume Next
  ElseIf err.Number = 20 Then
    Resume Next
  Else
  MsgBox err.Source & " Errore :" & Chr(13) & Chr(13) & "Error Number = " & CStr(err.Number) & Chr(13) & "Error Description = " & err.Description, vbOKOnly + vbExclamation, "  Error  -  " & err.Description
    err.Clear
    SaveICO = False
    Resume Cleanup
  End If
End Function

Sub resizePicture()
  If Picture1.Width > 480 Then Picture1.Width = 480
  If Picture1.Height > 480 Then Picture1.Height = 480
  If Picture2.Width > 480 Then Picture1.Width = 480
  If Picture2.Height > 480 Then Picture1.Height = 480
End Sub
