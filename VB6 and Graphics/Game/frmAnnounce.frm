VERSION 5.00
Begin VB.Form frmAnnounce 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmAnnounce"
   ClientHeight    =   6795
   ClientLeft      =   180
   ClientTop       =   945
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4140
      TabIndex        =   0
      Top             =   5175
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picAnnounce 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3960
      Left            =   1980
      ScaleHeight     =   3960
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   945
      Width           =   5475
      Begin VB.Image Image4 
         Height          =   180
         Left            =   225
         Picture         =   "frmAnnounce.frx":0000
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   5025
      End
      Begin VB.Image Image2 
         Height          =   180
         Left            =   225
         Picture         =   "frmAnnounce.frx":42D2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5010
      End
      Begin VB.Image Image1 
         Height          =   4155
         Left            =   5265
         Picture         =   "frmAnnounce.frx":85A4
         Stretch         =   -1  'True
         Top             =   -180
         Width           =   225
      End
      Begin VB.Image Image5 
         Height          =   4170
         Left            =   0
         Picture         =   "frmAnnounce.frx":BEC6
         Stretch         =   -1  'True
         Top             =   -225
         Width           =   225
      End
   End
   Begin VB.Image imgSabotageFailed 
      Height          =   600
      Left            =   8280
      Picture         =   "frmAnnounce.frx":F7E8
      Stretch         =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image imgFailedInvasion 
      Height          =   600
      Left            =   6750
      Picture         =   "frmAnnounce.frx":27DAA
      Stretch         =   -1  'True
      Top             =   5940
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image imgCaptured 
      Height          =   600
      Left            =   8280
      Picture         =   "frmAnnounce.frx":4036C
      Stretch         =   -1  'True
      Top             =   1485
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Image imgBioFailed 
      Height          =   510
      Left            =   8460
      Picture         =   "frmAnnounce.frx":5892E
      Stretch         =   -1  'True
      Top             =   6060
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgSabotage 
      Height          =   495
      Left            =   8490
      Picture         =   "frmAnnounce.frx":6F00C
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgExpanding 
      Height          =   585
      Left            =   8550
      Picture         =   "frmAnnounce.frx":8586E
      Stretch         =   -1  'True
      Top             =   4515
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image imgVictorious 
      Height          =   750
      Left            =   8280
      Picture         =   "frmAnnounce.frx":9B730
      Stretch         =   -1  'True
      Top             =   3555
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgOverrun 
      Height          =   705
      Left            =   8370
      Picture         =   "frmAnnounce.frx":B51DA
      Stretch         =   -1  'True
      Top             =   2835
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgExplosion 
      Height          =   525
      Left            =   8370
      Picture         =   "frmAnnounce.frx":CF29C
      Stretch         =   -1  'True
      Top             =   2175
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
'close this form and return to the game
PlaySoundEffect "Quiet"

Unload Me

End Sub


Private Sub Form_Activate()
Randomize
'draw lots of stars on the screen
    Dim a, X, Y
    For a = 1 To 500
        X = Int(Rnd * frmAnnounce.ScaleWidth)
        Y = Int(Rnd * frmAnnounce.ScaleHeight)
        frmAnnounce.PSet (X, Y), vbWhite
    Next a
    
    'draw dark grey stars
    Dim grey
    grey = &H808080
    For a = 1 To 400
        X = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        frmAnnounce.PSet (X, Y), grey
    Next a
    
    'draw some blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 400
       X = Int(Rnd * frmAnnounce.ScaleWidth)
       Y = Int(Rnd * frmAnnounce.ScaleHeight)
       frmAnnounce.PSet (X, Y), blue
    Next a



Select Case MessageType
    'Announceline variables set in frm_Load of frmGamescreen
    'set picture, sound and text for each message

Case "BioFailed"
     picAnnounce.Picture = imgBioFailed.Picture
     PlaySoundEffect "BioFail"
     BeepingText Announceline1, Announceline2, Announceline3

Case "Captured"
     picAnnounce.Picture = imgCaptured.Picture
     PlaySoundEffect "Overrun"
     BeepingText Announceline1, Announceline2, Announceline3
     
Case "Expanding"
     picAnnounce.Picture = imgExpanding.Picture
     PlaySoundEffect "Abort"
     BeepingText Announceline1, Announceline2, Announceline3
    
Case "Explosion"
    'set picture, sound and text
    picAnnounce.Picture = imgExplosion.Picture
    PlaySoundEffect "Explosion"
    BeepingText Announceline1, Announceline2, Announceline3
        
Case "Failed Invasion"
    'set picture, sound and text
    picAnnounce.Picture = imgBioFailed.Picture
    PlaySoundEffect "Overrun"
    BeepingText Announceline1, Announceline2, Announceline3

Case "Overrun"
    picAnnounce.Picture = imgOverrun.Picture
    PlaySoundEffect "Overrun"
    BeepingText Announceline1, Announceline2, Announceline3

Case "Sabotage"
    picAnnounce.Picture = imgSabotage.Picture
    PlaySoundEffect "Sabotage"
    BeepingText Announceline1, Announceline2, Announceline3
    
Case "Sabotage Failed"
    picAnnounce.Picture = imgBioFailed.Picture
    PlaySoundEffect "Sabotage"
    BeepingText Announceline1, Announceline2, Announceline3

Case "Victorious"
    picAnnounce.Picture = imgVictorious.Picture
    PlaySoundEffect "Overrun"
    BeepingText Announceline1, Announceline2, Announceline3
    
End Select

End Sub

Public Sub BeepingText(Announceline1 As String, Announceline2 As String, Announceline3 As String)
'*********************
'Taking the Announceline strings from frmGamescreen's form_load procedure and prints
'them out one character at a time.
'I removed the beep, as it interferes with the
'sound effects playing
'*********************

Dim Length As Integer
Dim Counter As Integer
Dim X As Long

'line 1
picAnnounce.CurrentX = 650
picAnnounce.CurrentY = 140

Length = Len(Announceline1)
For Counter = 1 To Length
    picAnnounce.Print Mid$(Announceline1, Counter, 1);
    'If SoundOn Then
    '    Beep
    'End If
    For X = 1 To 100000: Next X
Next Counter

'carriage return
picAnnounce.Print
picAnnounce.CurrentX = 650

Length = Len(Announceline2)
If Length = 0 Then
    'clear variables, enable ok button, exit sub
    Announceline1 = ""
    Announceline2 = ""
    Announceline3 = ""
    frmAnnounce.cmdOK.Visible = True
    frmAnnounce.cmdOK.Enabled = True
    
    Exit Sub
Else
    For Counter = 1 To Length
        picAnnounce.Print Mid$(Announceline2, Counter, 1);
        'If SoundOn Then
        '    Beep
        'End If
        For X = 1 To 100000: Next X
    Next Counter
End If

'carriage return
picAnnounce.Print
picAnnounce.CurrentX = 650

'line 3
Length = Len(Announceline3)

If Length = 0 Then
    'clear variables, enable ok button, exit sub
    Announceline1 = ""
    Announceline2 = ""
    Announceline3 = ""
    frmAnnounce.cmdOK.Visible = True
    frmAnnounce.cmdOK.Enabled = True
    
    Exit Sub
Else
    For Counter = 1 To Length
        picAnnounce.Print Mid$(Announceline3, Counter, 1);
        'If SoundOn Then
        '    Beep
        'End If
        For X = 1 To 100000: Next X
    Next Counter
End If

'clear variables
Announceline1 = ""
Announceline2 = ""
Announceline3 = ""

frmAnnounce.cmdOK.Visible = True
frmAnnounce.cmdOK.Enabled = True

End Sub










Private Sub Form_Load()

'reposition form if running in window mode
If DefaultGameSize = 0 Then
    Me.Top = 0
    Me.Left = 0
End If



End Sub


