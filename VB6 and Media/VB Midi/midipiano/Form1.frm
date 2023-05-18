VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "VB Midi Piano"
   ClientHeight    =   2190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vol 
      Height          =   1815
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   ";"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   15
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   13
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   10
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   8
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   6
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000007&
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   2175
      Index           =   16
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "."
      Height          =   2175
      Index           =   14
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   ","
      Height          =   2175
      Index           =   12
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "M"
      Height          =   2175
      Index           =   11
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "N"
      Height          =   2175
      Index           =   9
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "B"
      Height          =   2175
      Index           =   7
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "V"
      Height          =   2175
      Index           =   5
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "C"
      Height          =   2175
      Index           =   4
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "X"
      Height          =   2175
      Index           =   2
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox key 
      BackColor       =   &H80000009&
      Caption         =   "Z"
      Height          =   2175
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "vol"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu midi_devices 
      Caption         =   "Midi Device"
      Begin VB.Menu device 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu ChannelOption 
      Caption         =   "Channel"
      Begin VB.Menu chan 
         Caption         =   "1"
         Index           =   0
      End
      Begin VB.Menu chan 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu chan 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu chan 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu chan 
         Caption         =   "5"
         Index           =   4
      End
      Begin VB.Menu chan 
         Caption         =   "6"
         Index           =   5
      End
      Begin VB.Menu chan 
         Caption         =   "7"
         Index           =   6
      End
      Begin VB.Menu chan 
         Caption         =   "8"
         Index           =   7
      End
      Begin VB.Menu chan 
         Caption         =   "9"
         Index           =   8
      End
      Begin VB.Menu chan 
         Caption         =   "10"
         Index           =   9
      End
      Begin VB.Menu chan 
         Caption         =   "11"
         Index           =   10
      End
      Begin VB.Menu chan 
         Caption         =   "12"
         Index           =   11
      End
      Begin VB.Menu chan 
         Caption         =   "13"
         Index           =   12
      End
      Begin VB.Menu chan 
         Caption         =   "14"
         Index           =   13
      End
      Begin VB.Menu chan 
         Caption         =   "15"
         Index           =   14
      End
      Begin VB.Menu chan 
         Caption         =   "16"
         Index           =   15
      End
   End
   Begin VB.Menu base 
      Caption         =   "Base note"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INVALID_NOTE = -1     ' Code for keyboard keys that we don't handle

Dim numDevices As Long      ' number of midi output devices
Dim curDevice As Long       ' current midi device
Dim hmidi As Long           ' midi output handle
Dim rc As Long              ' return code
Dim midimsg As Long         ' midi output message buffer
Dim channel As Integer      ' midi output channel
Dim volume As Integer       ' midi volume
Dim baseNote As Integer     ' the first note on our "piano"

' Set the value for the starting note of the piano
Private Sub base_Click()
   Dim s As String
   Dim i As Integer
   s = InputBox("Enter the new base note for the keyboard (0 - 111)", "Base note", CStr(baseNote))
   If IsNumeric(s) Then
      i = CInt(s)
      If (i >= 0 And i < 112) Then
         baseNote = i
      End If
   End If
End Sub

' Select the midi output channel
Private Sub chan_Click(Index As Integer)
   chan(channel).Checked = False
   channel = Index
   chan(channel).Checked = True
End Sub

' Open the midi device selected in the menu. The menu index equals the
' midi device number + 1.
Private Sub device_Click(Index As Integer)
   device(curDevice + 1).Checked = False
   device(Index).Checked = True
   curDevice = Index - 1
   rc = midiOutClose(hmidi)
   rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
   If (rc <> 0) Then
      MsgBox "Couldn't open midi out, rc = " & rc
   End If
End Sub

' If user presses a keyboard key, start the corresponding midi note
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   StartNote NoteFromKey(KeyCode)
End Sub

' If user lifts a keyboard key, stop the corresponding midi note
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   StopNote NoteFromKey(KeyCode)
End Sub

Private Sub Form_Load()
   Dim i As Long
   Dim caps As MIDIOUTCAPS
   
   ' Set the first device as midi mapper
   device(0).Caption = "MIDI Mapper"
   device(0).Visible = True
   device(0).Enabled = True
   
   ' Get the rest of the midi devices
   numDevices = midiOutGetNumDevs()
   For i = 0 To (numDevices - 1)
      midiOutGetDevCaps i, caps, Len(caps)
      device(i + 1).Caption = caps.szPname
      device(i + 1).Visible = True
      device(i + 1).Enabled = True
   Next
   
   ' Select the MIDI Mapper as the default device
   device_Click (0)
   
   ' Set the default channel
   channel = 0
   chan(channel).Checked = True
   
   ' Set the base note
   baseNote = 60
   
   ' Set volume range
   volume = 127
   vol.Min = 127
   vol.Max = 0
   vol.Value = volume

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Close current midi device
   rc = midiOutClose(hmidi)
End Sub

' Start a note when user click on it
Private Sub key_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   StartNote (Index)
End Sub

' Stop the note when user lifts the mouse button
Private Sub key_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   StopNote (Index)
End Sub

' Press the button and send midi start event
Private Sub StartNote(Index As Integer)
   If (Index = INVALID_NOTE) Then
      Exit Sub
   End If
   If (key(Index).Value = 1) Then
      Exit Sub
   End If
   key(Index).Value = 1
   midimsg = &H90 + ((baseNote + Index) * &H100) + (volume * &H10000) + channel
   midiOutShortMsg hmidi, midimsg
End Sub

' Raise the button and send midi stop event
Private Sub StopNote(Index As Integer)
   If (Index = INVALID_NOTE) Then
      Exit Sub
   End If
   key(Index).Value = 0
   midimsg = &H80 + ((baseNote + Index) * &H100) + channel
   midiOutShortMsg hmidi, midimsg
End Sub

' Get the note corresponding to a keyboard key
Private Function NoteFromKey(key As Integer)
   NoteFromKey = INVALID_NOTE
   Select Case key
   Case vbKeyZ
      NoteFromKey = 0
   Case vbKeyS
      NoteFromKey = 1
   Case vbKeyX
      NoteFromKey = 2
   Case vbKeyD
      NoteFromKey = 3
   Case vbKeyC
      NoteFromKey = 4
   Case vbKeyV
      NoteFromKey = 5
   Case vbKeyG
      NoteFromKey = 6
   Case vbKeyB
      NoteFromKey = 7
   Case vbKeyH
      NoteFromKey = 8
   Case vbKeyN
      NoteFromKey = 9
   Case vbKeyJ
      NoteFromKey = 10
   Case vbKeyM
      NoteFromKey = 11
   Case 188 ' comma
      NoteFromKey = 12
   Case vbKeyL
      NoteFromKey = 13
   Case 190 ' period
      NoteFromKey = 14
   Case 186 ' semicolon
      NoteFromKey = 15
   Case 191 ' forward slash
      NoteFromKey = 16
   End Select

End Function

' Set the volume
Private Sub vol_Change()
   volume = vol.Value
End Sub
