VERSION 5.00
Begin VB.Form frmWavMix 
   Caption         =   "Using the WavMix"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlaySecond 
      Caption         =   "Play Second Sond"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayThird 
      Caption         =   "Play Third Sound"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayFourth 
      Caption         =   "Play Fourth Sound"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayOne 
      Caption         =   "Play First Sound"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmWavMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'The session handle
Dim Session As Long

'The channels
Dim ChannelOne As Long
Dim ChannelTwo As Long
Dim ChannelThree As Long
Dim ChannelFour As Long

'The wave handles
Dim SoundOne As Long
Dim SoundTwo As Long
Dim SoundThree As Long
Dim SoundFour As Long



Private Sub cmdPlayFourth_Click()
Dim SoundTypeFour As MIXPLAYPARAMS

'Set the structure
With SoundTypeFour
    .Size = Len(SoundTypeFour)
    .ChannelHi = HighWord(3)
    .ChannelLo = LowerWord(3)
    .hWndNotifyHi = HighWord(0)
    .hWndNotifyLo = LowerWord(0)
    .MixSessionHi = HighWord(Session)
    .MixSessionLo = LowerWord(Session)
    .FlagsHi = HighWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .FlagsHi = LowerWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .MixWaveHi = HighWord(SoundFour)
    .MixWaveLo = LowerWord(SoundFour)
    .wLoops = 0

End With

'Play it
WaveMixPlay SoundTypeFour

End Sub

Private Sub cmdPlayOne_Click()
Dim SoundTypeOne As MIXPLAYPARAMS

'Set the structure
With SoundTypeOne
    .Size = Len(SoundTypeOne)
    .ChannelHi = HighWord(0)
    .ChannelLo = LowerWord(0)
    .hWndNotifyHi = HighWord(0)
    .hWndNotifyLo = LowerWord(0)
    .MixSessionHi = HighWord(Session)
    .MixSessionLo = LowerWord(Session)
    .FlagsHi = HighWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .FlagsHi = LowerWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .MixWaveHi = HighWord(SoundOne)
    .MixWaveLo = LowerWord(SoundOne)
    .wLoops = 0

End With

'Play it
WaveMixPlay SoundTypeOne

End Sub

Private Sub cmdPlaySecond_Click()
Dim SoundTypeTwo As MIXPLAYPARAMS

'Set the structure
With SoundTypeTwo
    .Size = Len(SoundTypeTwo)
    .ChannelHi = HighWord(1)
    .ChannelLo = LowerWord(1)
    .hWndNotifyHi = HighWord(0)
    .hWndNotifyLo = LowerWord(0)
    .MixSessionHi = HighWord(Session)
    .MixSessionLo = LowerWord(Session)
    .FlagsHi = HighWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .FlagsHi = LowerWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .MixWaveHi = HighWord(SoundTwo)
    .MixWaveLo = LowerWord(SoundTwo)
    .wLoops = 0

End With

'Play it
WaveMixPlay SoundTypeTwo
End Sub

Private Sub cmdPlayThird_Click()
Dim SoundTypeThree As MIXPLAYPARAMS

'Set the structure
With SoundTypeThree
    .Size = Len(SoundTypeThree)
    .ChannelHi = HighWord(2)
    .ChannelLo = LowerWord(2)
    .hWndNotifyHi = HighWord(0)
    .hWndNotifyLo = LowerWord(0)
    .MixSessionHi = HighWord(Session)
    .MixSessionLo = LowerWord(Session)
    .FlagsHi = HighWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .FlagsHi = LowerWord(WMIX_CLEARQUEUE Or WMIX_HIGHPRIORITY)
    .MixWaveHi = HighWord(SoundThree)
    .MixWaveLo = LowerWord(SoundThree)
    .wLoops = 0

End With

'Play it
WaveMixPlay SoundTypeThree

End Sub

Private Sub Form_Load()

Dim AppPath As String

AppPath = App.Path

'Start the session
Session = WaveMixInit()

If Session = 0 Then 'The session could be created
    MsgBox "Unable to create session"
    Exit Sub
    Unload Me
End If

'Load the wave files
SoundOne = WaveMixOpenWave(Session, ByVal (App.Path & "\1.wav"), 0, WMIX_FILE) 'App.Path & "\1.wav", 0, WMIX_FILE)
SoundTwo = WaveMixOpenWave(Session, ByVal (App.Path & "\2.wav"), 0, WMIX_FILE)
SoundThree = WaveMixOpenWave(Session, ByVal (App.Path & "\3.wav"), 0, WMIX_FILE)
SoundFour = WaveMixOpenWave(Session, ByVal (App.Path & "\4.wav"), 0, WMIX_FILE)

'Open the channels 0 to 3
Debug.Print WaveMixOpenChannel(Session, 4, WMIX_OPENCOUNT)

'Activate the channels
Debug.Print WaveMixActivate(Session, WMIX_ACTIVATE)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Close all the channels
WaveMixCloseChannel Session, 0, WMIX_ALL

'Free the wave resources
WaveMixFreeWave Session, SoundOne
WaveMixFreeWave Session, SoundTwo
WaveMixFreeWave Session, SoundThree
WaveMixFreeWave Session, SoundFour

'and finally close the session
WaveMixCloseSession Session

End Sub
