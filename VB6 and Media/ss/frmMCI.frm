VERSION 5.00
Begin VB.Form frmMCI 
   Caption         =   "Using the MCI"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Sound"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Sound"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Copy the 'sound.wav' file to the root of your HD drive before trying out this sample project"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "frmMCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
                                            (ByVal lpstrCommand As String, _
                                             ByVal lpstrReturnString As String, _
                                             ByVal uReturnLength As Long, _
                                             ByVal hwndCallback As Long) As Long

Dim FileName As String

Private Sub cmdPlay_Click()
Dim CommandString As String
  'Build the command string
  CommandString = "open " & FileName & " type waveaudio alias WAVE10"
  'Open the sound
  mciSendString CommandString, 0&, 0, 0
  'Play the sound
  mciSendString "PLAY WAVE1 FROM 0", 0&, 0, 0

End Sub

Private Sub cmdStop_Click()

'Stop the sound
mciSendString "STOP WAVE1", 0&, 0, 0

End Sub

Private Sub Form_Load()

'Remember to copy the sound.wav file to the root of your HD drive nad change the letter
'so it fits appropriately
FileName = "c:\test.wav"

End Sub
