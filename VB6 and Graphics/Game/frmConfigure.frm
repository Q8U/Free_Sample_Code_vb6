VERSION 5.00
Begin VB.Form frmConfigure 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configure Default Settings"
   ClientHeight    =   4305
   ClientLeft      =   1620
   ClientTop       =   1530
   ClientWidth     =   5955
   Icon            =   "frmConfigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdConfigure 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2160
      TabIndex        =   6
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1230
      Left            =   1215
      TabIndex        =   1
      Top             =   2115
      Width           =   3435
      Begin VB.OptionButton optViewInWindow 
         BackColor       =   &H00000000&
         Caption         =   "View Game in Window"
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   270
         TabIndex        =   5
         Top             =   765
         Width           =   2085
      End
      Begin VB.OptionButton optFullScreen 
         BackColor       =   &H00000000&
         Caption         =   "Full Screen"
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Sound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1230
      Left            =   1215
      TabIndex        =   0
      Top             =   855
      Width           =   3435
      Begin VB.OptionButton optSoundOFF 
         BackColor       =   &H00000000&
         Caption         =   "OFF"
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   765
         Width           =   2715
      End
      Begin VB.OptionButton optSoundON 
         BackColor       =   &H00000000&
         Caption         =   "ON"
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2445
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Configure Default Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1350
      TabIndex        =   7
      Top             =   315
      Width           =   3255
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   225
      Picture         =   "frmConfigure.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5520
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   225
      Picture         =   "frmConfigure.frx":4714
      Stretch         =   -1  'True
      Top             =   4095
      Width           =   5520
   End
   Begin VB.Image Image2 
      Height          =   4290
      Left            =   5715
      Picture         =   "frmConfigure.frx":89E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   4290
      Left            =   0
      Picture         =   "frmConfigure.frx":C308
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   240
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub cmdConfigure_Click()

If optSoundON Then
    SoundOn = True
Else
    SoundOn = False
End If

If optFullScreen Then
    DefaultGameSize = 2
    frmGameScreen.WindowState = 2   'maximized
Else
    DefaultGameSize = 0
    frmGameScreen.WindowState = 0   'normal
End If

'write a new config file
Dim Filename As String
Dim ShortName As String

ShortName = "\4000cfg.txt"
Filename = App.Path & ShortName

'get a free file number
gFileNum = FreeFile

'create the file
Open Filename For Output As gFileNum

'write new default options
Write #gFileNum, optSoundON, DefaultGameSize

'close the file
Close gFileNum

Unload Me


End Sub

Private Sub Form_Load()

'disable the X button on the control box
Call DisableX(Me)

'set option buttons as per existing configuration
If SoundOn Then
    optSoundON = True
Else
    optSoundON = False
End If

If frmGameScreen.WindowState = vbMaximized Then
    optFullScreen = True
    optViewInWindow = False
Else
    optFullScreen = False
    optViewInWindow = True
End If

End Sub


Private Sub Option1_Click()

End Sub


