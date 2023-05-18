VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   600
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   4620
      ScaleWidth      =   6060
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   6060
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   -3720
      MouseIcon       =   "Form1.frx":1ED5C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1F066
      ScaleHeight     =   4620
      ScaleWidth      =   6060
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Image Image9 
         Height          =   2295
         Left            =   0
         MouseIcon       =   "Form1.frx":3DAB8
         MousePointer    =   99  'Custom
         Top             =   2280
         Width           =   6015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7080
      Top             =   360
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   5640
      Picture         =   "Form1.frx":3DDC2
      ScaleHeight     =   308
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   403
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   6045
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   4350
         Width           =   6015
      End
      Begin VB.Image Image10 
         Height          =   1215
         Left            =   2040
         MouseIcon       =   "Form1.frx":5C814
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   735
         Left            =   0
         MouseIcon       =   "Form1.frx":5CB1E
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   6015
      End
      Begin VB.Image Image7 
         Height          =   3615
         Left            =   3960
         MouseIcon       =   "Form1.frx":5CE28
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image Image6 
         Height          =   3735
         Left            =   0
         MouseIcon       =   "Form1.frx":5D132
         MousePointer    =   99  'Custom
         Top             =   -120
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   -1320
      MouseIcon       =   "Form1.frx":5D43C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5D746
      ScaleHeight     =   308
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   6060
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   2295
         Picture         =   "Form1.frx":B89BA
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   51
         TabIndex        =   4
         Top             =   1365
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Image Image5 
         Height          =   1215
         Left            =   2160
         MouseIcon       =   "Form1.frx":B9788
         MousePointer    =   99  'Custom
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   4575
         Left            =   0
         MouseIcon       =   "Form1.frx":B9A92
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   1320
      MouseIcon       =   "Form1.frx":B9D9C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":BA0A6
      ScaleHeight     =   308
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Image Image11 
         Height          =   1215
         Left            =   1560
         MouseIcon       =   "Form1.frx":D8AF8
         MousePointer    =   99  'Custom
         Top             =   3480
         Width           =   3495
      End
      Begin VB.Image Image3 
         Height          =   4575
         Left            =   4320
         MouseIcon       =   "Form1.frx":D8E02
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   4575
         Left            =   0
         MouseIcon       =   "Form1.frx":D910C
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   1200
      MouseIcon       =   "Form1.frx":D9416
      Picture         =   "Form1.frx":D9720
      ScaleHeight     =   308
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   6060
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   1920
         MouseIcon       =   "Form1.frx":F8172
         MousePointer    =   99  'Custom
         Top             =   960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Dim ProjPath As String

Private Sub Command2_Click()
'Everything here closes all of our
'MCI devices
Close All

i = mciSendString("close voice1", 0&, 0, 0)
i = mciSendString("close voice2", 0&, 0, 0)
i = mciSendString("close video", 0&, 0, 0)
i = mciSendString("close video2", 0&, 0, 0)
i = mciSendString("close mid1", 0&, 0, 0)
       
End
End Sub


Private Sub Form_Load()
'If we're already running, then quit. We do not
'want two or more programs running at the same time
    If (App.PrevInstance = True) Then
        End
    End If
    
'To be safe, close any open MCI devices or files from prior use
Close All

'All of our multimedia is in the same
'directory as the EXE program. Thus we must
'always refer to that directory path
ProjPath = App.Path

'Since VB wasn't designed the way we like it. I have to
'fix a bug in it. The root directory of a hard drive
'(or any drive doesn't return a "\", so we must put one in.
If Right$(ProjPath, 1) = "\" Then
Else
  ProjPath = ProjPath & "\"
End If

'Close all devices that this program may be left opened
'if the user ran this program recently.
i = mciSendString("close voice1", 0&, 0, 0)
i = mciSendString("close voice2", 0&, 0, 0)
i = mciSendString("close video", 0&, 0, 0)
i = mciSendString("close video2", 0&, 0, 0)
i = mciSendString("close mid1", 0&, 0, 0)



'Open most of our MCI devices ahead of time. They will load into
'memory and are ready for instant useage.
i = mciSendString("open " & ProjPath & "no2.wav type waveaudio alias voice1", 0&, 0, 0)
i = mciSendString("open " & ProjPath & "click.wav type waveaudio alias voice2", 0&, 0, 0)
i = mciSendString("open " & ProjPath & "haunted.mid type sequencer alias mid1", 0&, 0, 0)


'Center the opening screen to the center of the monitor screen
Picture1.Left = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - (Picture1.Width / 2)
Picture1.Top = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - (Picture1.Height / 2)



'move all scenes (picture boxes) into the center of the
'screen, based on the opening scene position.
Picture2.Left = Picture1.Left
Picture2.Top = Picture1.Top

Picture3.Left = Picture1.Left
Picture3.Top = Picture1.Top

Picture5.Left = Picture1.Left
Picture5.Top = Picture1.Top

Picture6.Left = Picture1.Left
Picture6.Top = Picture1.Top

Picture7.Left = Picture1.Left
Picture7.Top = Picture1.Top


'Open and load "Earth Globe" AVI into a picture box
'ahead of time for quick usage.
       Last$ = Form1.Picture5.hWnd & " Style " & &H40000000
       ToDo$ = "open " & ProjPath & "yeah2.avi Type avivideo Alias video2 parent " & Last$
       i = mciSendString(ToDo$, 0&, 0, 0)
       i = mciSendString("put video2 window at 83 79 243 151", 0&, 0, 0)
 
'Set scroll bar values
HScroll1.Min = 1
HScroll1.Max = 45 '45 because the AVI has 45 frames
HScroll1.LargeChange = 1
HScroll1.SmallChange = 1


'Start the music!!
i = mciSendString("play mid1", 0&, 0, 0)


'Show the opening scene
Picture7.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Everything here closes all of our
'open MCI devices. No matter how the user
'closes this program.

Close All

  i = mciSendString("close voice1", 0&, 0, 0)
  i = mciSendString("close voice2", 0&, 0, 0)
  i = mciSendString("close video", 0&, 0, 0)
  i = mciSendString("close video2", 0&, 0, 0)
  i = mciSendString("close mid1", 0&, 0, 0)
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Everything here closes all of our
'open MCI devices. No matter how the user
'closes this program.
  i = mciSendString("close voice1", 0&, 0, 0)
  i = mciSendString("close voice2", 0&, 0, 0)
  i = mciSendString("close video", 0&, 0, 0)
  i = mciSendString("close video2", 0&, 0, 0)
  i = mciSendString("close mid1", 0&, 0, 0)
End Sub




Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()

'Here, we set the frame position of the "earth globe" according
'to the scroll bar value.
    i = mciSendString("seek video2 to " & HScroll1.Value, 0&, 0, 0)
End Sub


Private Sub Image1_Click()


'This string will allow thw AVI (Painting) to be displayed
'onto a picturebox of form1
       Last$ = Form1.Picture1.hWnd & " Style " & &H40000000

'Last$ = Form1.Picture1.hWnd & " Style " & &H40000000
'This string will use the information above, and also
'use the AVI file that you choose
       ToDo$ = "open " & ProjPath & "pic.avi Type avivideo Alias video parent " & Last$
'The strings above are now executed below to "open" the
'AVI file (The AVI file is not in view yet, just ready to be played)
       i = mciSendString(ToDo$, 0&, 0, 0)
 'This command will set the location of the AVI file
'onto your Form.
'
'Notice the AVI Coordinates:
'132 71 133 134 = Left Top Width Height
'
'The AVI file can be stretched horizontal and vertical
'based on your coordinate settings.
       i = mciSendString("put video window at 132 71 133 134", 0&, 0, 0)


'This command plays the AVI with the "wait" flag.
'You can eliminate the "wait" flag if you want, but
'you must remove the "close" statement below too.
       i = mciSendString("play video wait", 0&, 0, 0)


'This command closes the AVI file, otherwise the file will
'still be "open", even if you exit your program! it is best to al
'so place this in the Form Unload Sub.
'   We also close this AVI because the user CANNOT click a
'control that has an open VAI on top of it.
       i = mciSendString("close video", 0&, 0, 0)

End Sub




Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture6.Visible = True
   Picture5.Visible = False
End Sub


Private Sub Image11_Click()
Picture7.Visible = True
Picture2.Visible = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
Picture1.Visible = True
End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
Picture3.Visible = True

End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
Picture2.Visible = True

End Sub


Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'While mouse button is down
'Show the red button. it hides the blue button
  Picture4.Visible = True
  
  'The click WAV:
  'Play only from 0 to 250 because this is the
  'first half of a click noise
  i = mciSendString("play voice2 from 0 to 250", 0&, 0, 0)
  
  
  
  
Dim mssg As String * 255
Dim L As Integer


'Check to see if the MUSIC (mid1) is "paused"
RequestStat$ = "status mid1 mode"

i = mciSendString(RequestStat$, mssg, 255, 0)
  'If the music is paused, then resume it
  If Left$(mssg, 6) = "paused" Then
    i = mciSendString("resume mid1", 0&, 0, 0)
  Else
    'Since the music is NOT paused, we will pause it.
    'This will not affect our timer for music because our timer
    'only detects "stopped", not "pause"
    i = mciSendString("pause mid1", 0&, 0, 0)
  End If
End Sub


Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'When the user releases the mouse button,
'We show the blue button again by making the picture
'of the red button invisible.

Picture4.Visible = False

  'Play the rest of the click noise. Voice2
  'starts playing where we last finished from.
  i = mciSendString("play voice2", 0&, 0, 0)

End Sub


Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This shows a new scene and hides another.
'Gives the user the effect of moving to
'different areas of the room
Picture1.Visible = True
Picture5.Visible = False
End Sub


Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This shows a new scene and hides another.
'Gives the user the effect of moving to
'different areas of the room
Picture3.Visible = True
Picture5.Visible = False
End Sub


Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This shows a new scene and hides another.
'Gives the user the effect of moving to
'different areas of the room
Picture2.Visible = True
Picture5.Visible = False
End Sub





Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This shows a new scene and hides another.
'Gives the user the effect of moving to
'different areas of the room
Picture5.Visible = True
Picture6.Visible = False
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is the scene where you see the picture on the wall.
'This code  will determine whew the mouse
'cursor is. If cursor is on the right of the
'painting, then we hide this scene and make
'the next scene visible
If X <= 296 Then
  i = mciSendString("play voice1 from 0", 0&, 0, 0)
Else
Picture1.Visible = False
Picture2.Visible = True
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is the scene of the painting up close
'Here, we detect the location of the mouse cursor.
'If the cursor is on the right of the painting, then
'we make the cursor look like a hand pointing to the right.
If X <= 296 Then
Picture1.MousePointer = 0
Else
Picture1.MousePointer = 99
End If


End Sub




Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This shows a new scene and hides another.
'Gives the user the effect of moving to
'different areas of the room
Picture5.Visible = True
Picture2.Visible = False
End Sub
Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Show the second scene when the game starts
Picture2.Visible = True
'Close the opening scene
Picture7.Visible = False
End Sub
Private Sub Timer1_Timer()
Dim mssg As String * 255

'MUSIC LOOP:
'If the MIDI music is finished playing and reached the end,
'we want to loop back to the beginning and start playing again.
 
 RequestStat$ = "status mid1 mode" 'Check to see if the music stopped
 i = mciSendString(RequestStat$, mssg, 255, 0)
 
 'If if is "stopped" then start it from the beginning
 If Left$(mssg, 7) = "stopped" Then i = mciSendString("play mid1 from 0", 0&, 0, 0)


End Sub


