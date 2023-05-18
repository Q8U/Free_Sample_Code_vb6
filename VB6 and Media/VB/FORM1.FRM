VERSION 5.00
Object = "{EF85CC23-AFDF-101D-85F5-6EBA1EE93AF4}#1.1#0"; "Wave32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin WaveLib.Wave Wave1 
      Left            =   120
      Top             =   120
      _Version        =   65537
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   64
      Exclusive       =   0   'False
      Filename        =   ""
      FileLength      =   -1
      Loop            =   0   'False
      PlayEnd         =   -1
      PlayStart       =   -1
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play File #2"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play File #1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Background"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "File #2:"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "File #1:"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Background File:"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lPos As Long

Private Sub Command1_Click()
    'This variable is used to
    'keep track of the position of
    'the background WAV file.
    lPos = 1
    '
    'all of the Wave starting code is in
    'a sub since we need to repeat this after
    'playing the separate files, so no sense
    'in having duplicate code
    StartLoop
End Sub

Private Sub Command2_Click()
'make sure the file exists
If Text2.Text = "" Or Dir$(Text2.Text) = "" Then
    Exit Sub
End If
'
'the code is the same whether playing File 1 or File 2 so I created a
'sub for this
PlayFile (Text2.Text)
End Sub

Private Sub Command3_Click()
    'make sure the file exists
    If Text3.Text = "" Or Dir$(Text3.Text) = "" Then
        Exit Sub
    End If
    '
    'the code is the same whether playing File 1 or File 2 so I created a
    'sub for this
    PlayFile (Text3.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Form1 = Nothing
End Sub
Private Sub StartLoop()
    'make sure the file exists
    If Text1.Text = "" Or Dir$(Text1.Text) = "" Then
        Exit Sub
    End If
    Command1.Enabled = False
    
    Wave1.Tag = ""
    Wave1.filename = Text1.Text
    '
    'set the start position
    Wave1.PlayStart = lPos
    If lPos = 1 Then
        '
        'if pos = 1, we can start in
        'loop mode at the beginning
        Wave1.Exclusive = False
        Wave1.Loop = True
    Else
        'otherwise, we must start it
        'from where we left off in non-loop
        'mode. The PlayDone event restarts
        'it in loop mode.
        '
        'After that finishes, we can restart
        'it in loop mode from the beginning
        '
        'If we didn't do this, it would re-start
        'in loop mode from lPos and the beginning
        'of the WAV would be skipped in each loop
        Wave1.Exclusive = False
        Wave1.Loop = False
    End If
    Wave1.Action = wAPlay
End Sub

Private Sub Wave1_PlayDone()
    If Wave1.Tag = "halted by code" Then
        '
        'if the event fired because we set the
        'property to 4 then we don't do
        'anything
        Exit Sub
    ElseIf Wave1.Tag = "file" Then
        '
        'we just finished playing the
        'sound, so start the loop from
        'where we left off
        StartLoop
    Else
        '
        'otherwise, we set the position to the
        'beginning and call  StartLoop which
        'will start it in loop mode again
        lPos = 1
        StartLoop
    End If
End Sub

Private Sub PlayFile(sFile As String)
'
'pause the background sound
Wave1.Action = wAPause
'record the position
'of the background WAV
lPos = Wave1.Position
'
'now we stop the file because we
'need to free up the device for the
'other file
'
'Note -- when the Action is set to
'4, the PlayDone event fires. We need
'to know whether PlayDone has fired
'because the file has finished or whether
'it was because the Action was set to 4
'
'If the file is finished, we need to know
'whether it was our background file or not
'
'We will use the Tag property to track this
Wave1.Tag = "halted by code"
Wave1.Action = wAStop
DoEvents
'
'make note that we are playing a file
Wave1.Tag = "file"
'
'When it finishes, PlayDone will examine
'the Tag and restart the loop for us
Wave1.filename = sFile
Wave1.Loop = False
Wave1.Exclusive = True
Wave1.Action = wAPlay
End Sub
