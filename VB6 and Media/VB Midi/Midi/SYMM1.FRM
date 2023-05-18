VERSION 5.00
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Send SYSEX"
   ClientHeight    =   3240
   ClientLeft      =   1125
   ClientTop       =   1605
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3240
   ScaleWidth      =   7425
   Begin VB.PictureBox Bender 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   6360
      ScaleHeight     =   16383
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   9
      Top             =   300
      Width           =   975
   End
   Begin VB.ListBox lst_sound_list 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   2940
      TabIndex        =   5
      Top             =   1020
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "All Sound Off"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   60
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   5760
      ScaleHeight     =   127
      ScaleMode       =   0  'User
      ScaleWidth      =   405
      TabIndex        =   4
      Top             =   300
      Width           =   435
   End
   Begin VB.VScrollBar sb_play 
      Height          =   2895
      LargeChange     =   12
      Left            =   5220
      Max             =   127
      TabIndex        =   3
      Top             =   300
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Play Note C4"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   2835
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Bender"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6360
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Play here"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5220
      TabIndex        =   6
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Sounds"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2940
      TabIndex        =   8
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Devices"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   705
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bender_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static last_y_value As Single

    ' Remove previous line by overdrawing with inverted draw mode
    ' scale width set from 0 to 1
    Bender.Line (0, last_y_value)-(1, last_y_value)
    Call Bendersub(0, y) ' scaleheight set to 0 to 16383
    ' Draw new line
    Bender.Line (0, y)-(1, y)
    last_y_value = y
End Sub

Private Sub Command1_Click()
    Call note_on(0, 60, 127)
End Sub

Private Sub Command2_Click()
    all_sounds_off
End Sub

Private Sub fill_sound_list()
Dim s As String

    Open App.Path & "\genmidi.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, s
        lst_sound_list.AddItem s
    Loop
    Close #1
End Sub

Private Sub Form_Load()
    Call midi_listoutdevs(List1)
    Call fill_sound_list
End Sub

Private Sub Form_Unload(Cancel As Integer)
    midi_out_close
End Sub

Private Sub List1_Click()
Dim x  As Integer

    midi_out_close
    x = midi_out_open(List1.ItemData(List1.ListIndex))
End Sub

Private Sub lst_sound_list_Click()
    Call program_change(0, 0, lst_sound_list.ListIndex)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call note_on(0, y, 127)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call note_off(0, y)
End Sub

Private Sub sb_play_Change()
Static prev_note As Integer ' remember variable (static)

    ' Turn off previous note
    Call note_off(0, prev_note)
    ' turn on this note
    Call note_on(0, sb_play.value, 127) ' Max velocity
    ' save note as previous
    prev_note = sb_play.value
End Sub

