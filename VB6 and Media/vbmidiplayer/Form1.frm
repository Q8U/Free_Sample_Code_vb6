VERSION 5.00
Object = "{6E2C9294-8D14-11D1-8213-0060975EACAF}#3.0#0"; "VBUMidiPlayer.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   1860
   ClientTop       =   2415
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   Begin VB.CommandButton Command4 
      Caption         =   "Stop2"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2100
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VBUMidiFilePlayer.VBUMidiPlayer VBUMidiPlayer2 
      Left            =   480
      Top             =   1680
      _ExtentX        =   794
      _ExtentY        =   688
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   975
   End
   Begin VBUMidiFilePlayer.VBUMidiPlayer VBUMidiPlayer1 
      Left            =   480
      Top             =   480
      _ExtentX        =   794
      _ExtentY        =   688
      Repeat          =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    VBUMidiPlayer1.MidiPlay
End Sub

Private Sub Command2_Click()
    VBUMidiPlayer1.MidiStop
    
End Sub

Private Sub Command3_Click()
    VBUMidiPlayer2.MidiPlay

End Sub

Private Sub Command4_Click()
    VBUMidiPlayer2.MidiStop
    
End Sub

Private Sub Form_Load()
    If VBUMidiPlayer1.filename = "" Then
        VBUMidiPlayer1.filename = App.Path & "\MCITest.mid"
    End If
    If VBUMidiPlayer2.filename = "" Then
        VBUMidiPlayer2.filename = App.Path & "\MyMusic.mid"
    End If
End Sub

Private Sub VBUMidiPlayer1_Completed()
    Debug.Print "Midi1 Finished"
End Sub

Private Sub VBUMidiPlayer1_MidiError(ErrorNumber As Long, ErrorMsg As String)
    Debug.Print "Midi1 " & ErrorMsg

End Sub

Private Sub VBUMidiPlayer2_Completed()
    Debug.Print "Midi2 Finished"

End Sub

Private Sub VBUMidiPlayer2_MidiError(ErrorNumber As Long, ErrorMsg As String)
    Debug.Print "Midi2 " & ErrorMsg
End Sub
