VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusy 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBusy.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1500
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox ViewPic2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3600
      Picture         =   "frmBusy.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar prgFile 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   4515
   End
   Begin VB.Image ViewPic1 
      Height          =   480
      Left            =   360
      Picture         =   "frmBusy.frx":0BD4
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    frmMain.ZOrder
End Sub
