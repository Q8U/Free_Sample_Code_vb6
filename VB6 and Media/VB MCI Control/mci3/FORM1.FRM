VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1560
   ScaleWidth      =   4950
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   1080
   End
   Begin VB.ComboBox cboFilename 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   960
      List            =   "Form1.frx":0016
      TabIndex        =   2
      Text            =   "C:\WinNT\Media\Jungle Asterisk.wav"
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WAV File"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Open the device and play the sound.
Private Sub cmdPlay_Click()
    If Timer1.Enabled Then
        Timer1.Enabled = False
        cmdPlay.Caption = "Play"

        ' Close the device.
        MMControl1.Command = "Close"
    Else
        Timer1.Enabled = True
        cmdPlay.Caption = "Stop"
    
        ' Set the file name.
        MMControl1.FileName = cboFilename.Text

        ' Open the MCI device.
        MMControl1.Wait = True
        MMControl1.Command = "Open"
    End If
End Sub
Private Sub Form_Load()
    ' Prepare the MCI control for WaveAudio.
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
End Sub


Private Sub Timer1_Timer()
    ' See what mode the MCI control is in.
    If MMControl1.Mode = mciModeStop Then
            MMControl1.Wait = True
            MMControl1.Command = "Prev"
            MMControl1.Command = "Play"
    End If
End Sub


