VERSION 5.00
Begin VB.Form frmsndSound 
   Caption         =   "Playing with sndPlaySound"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStopFile 
      Caption         =   "Stop Sound"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdMemory 
      Caption         =   "Play from memory"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Play from File"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmsndSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'''sndPlaySound Constants
Const SND_ALIAS = &H10000
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_MEMORY = &H4
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_SYNC = &H0


Dim SoundFile As String

'Sound in Memory variable
Dim SoundInMemory As String

Private Sub cmdFile_Click()
    sndPlaySound SoundFile, SND_ASYNC Or SND_FILENAME
End Sub

Private Sub cmdMemory_Click()

    sndPlaySound SoundInMemory, SND_MEMORY Or SND_ASYNC

End Sub

Private Sub cmdStopFile_Click()

sndPlaySound 0, 0

End Sub

Private Sub Form_Load()
Dim TempCaption As String
'Update and refresh the form
TempCaption = Me.Caption

SoundFile = App.Path & "\sound.wav"

Me.Caption = "Loading " & SoundFile & " Please wait..."

'Show and refresh the form
Me.Show
Me.Refresh

'Load the sound into memory
SoundInMemory = LoadSound(SoundFile)

'Set the old caption back
Me.Caption = TempCaption

End Sub


'Loads a sound from file into memory
Private Function LoadSound(FileName As String) As String
On Error GoTo Error_Handler

Dim FreeFileNumber As Integer
Dim SoundBuffer As String

FreeFileNumber = FreeFile
SoundBuffer = Space$(FileLen(FileName)) 'Make room for sound file

Open FileName For Binary As #FreeFileNumber
         
         Get #FreeFileNumber, , SoundBuffer
         
Close FreeFileNumber

'remove wasted spaces
LoadSound = Trim$(SoundBuffer)
    
Error_Handler:

    Select Case Err
    
        Case 0 'No error
        
        Case Else
            
            MsgBox Err.Description & vbCrLf & "Error Number: " & Err.Number
            Err.Clear
            LoadSound = ""
    End Select
   
End Function
