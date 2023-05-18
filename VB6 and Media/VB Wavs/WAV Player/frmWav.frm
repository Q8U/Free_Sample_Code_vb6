VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWav 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play the Wav"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmWav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

MMControl1.DeviceType = "WaveAudio"
MMControl1.FileName = "Directory path of wav file" 'example C:\WINDOWS\MEDIA\tada.wav
MMControl1.Command = "Open"
MMControl1.Command = "Play"
MMControl1.Command = "Prev"

End Sub

