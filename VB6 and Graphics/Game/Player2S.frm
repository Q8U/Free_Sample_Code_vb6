VERSION 5.00
Begin VB.Form frmPlayer2Setup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "  Player 2 - Setup"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PLAYER2S.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   3660
      TabIndex        =   4
      Top             =   2910
      Width           =   1200
   End
   Begin VB.TextBox txtHomePlanetName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   4215
      TabIndex        =   3
      Top             =   1995
      Width           =   1740
   End
   Begin VB.TextBox txtPlayer2Name 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   4215
      TabIndex        =   1
      Top             =   1305
      Width           =   1740
   End
   Begin VB.Label lblHomePlanetName 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Planet:"
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   2730
      TabIndex        =   2
      Top             =   2025
      Width           =   1455
   End
   Begin VB.Label lblPlayer2Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2 Name:"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2565
      TabIndex        =   0
      Top             =   1350
      Width           =   1620
   End
End
Attribute VB_Name = "frmPlayer2Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Counter As Integer  'for number of letters in name
Private Sub cmdOK_Click()

'Don't let player 2 avoid entering a name
If txtPlayer2Name.Text = "" Then
    PlaySoundEffect "Quiet"
    MsgBox "Please choose a name", , ""
    txtPlayer2Name.SetFocus
    Exit Sub
End If

Player(1).Name = txtPlayer2Name.Text

'assign name for player 2 home planet if none entered
If txtHomePlanetName = "" Then
    'select one of 5 random names
    Randomize
    Dim X As Integer
    X = Int(Rnd * 5) + 1
    Select Case X
    Case 1
        Planet(Player(1).HomePlanet).Name = "Azegaar"
    Case 2
        Planet(Player(1).HomePlanet).Name = "Ilsonka"
    Case 3
        Planet(Player(1).HomePlanet).Name = "Rombi"
    Case 4
        Planet(Player(1).HomePlanet).Name = "Talagor"
    Case 5
        Planet(Player(1).HomePlanet).Name = "Myrdaltos"
    End Select
Else
    Planet(Player(1).HomePlanet).Name = txtHomePlanetName.Text
End If

Unload frmPlayer2Setup

End Sub


Private Sub Form_Activate()
Randomize

    'draw white stars on the screen
    Dim a, X, Y
    For a = 1 To 300
        X = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        Me.PSet (X, Y), vbWhite
    Next a
    
    'draw dark grey stars
    Dim grey
    grey = &H808080
    For a = 1 To 400
        X = Int(Rnd * frmCover.ScaleWidth)
        Y = Int(Rnd * frmCover.ScaleHeight)
        Me.PSet (X, Y), grey
    Next a
    
    'draw some blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 300
       X = Int(Rnd * frmCover.ScaleWidth)
       Y = Int(Rnd * frmCover.ScaleHeight)
       Me.PSet (X, Y), blue
    Next a

End Sub



Private Sub txtHomePlanetName_GotFocus()
txtHomePlanetName.Text = ""

End Sub


Private Sub txtPlayer2Name_GotFocus()
txtPlayer2Name.Text = ""

'enable command button to be used with enter key
cmdOK.Default = True

End Sub


Private Sub txtPlayer2Name_KeyDown(KeyCode As Integer, Shift As Integer)
If Counter > 12 Then
    KeyCode = 0
End If

End Sub


Private Sub txtPlayer2Name_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
        'backspace
        txtPlayer2Name.Text = ""
        Counter = 0
    Case 32
        'space bar
        Counter = Counter + 1
        If Counter > 14 Then
            KeyAscii = 0
            Beep
        End If
    Case Else
        Counter = Counter + 1
        If Counter > 14 Then
           KeyAscii = 0
           Beep
        End If
End Select

End Sub


