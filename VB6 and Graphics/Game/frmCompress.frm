VERSION 5.00
Object = "{D2D9B7C1-7650-11D1-9481-00A0247B7657}#1.0#0"; "ZLIBOCX2.DLL"
Begin VB.Form frmCompress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6900
   ClientLeft      =   930
   ClientTop       =   525
   ClientWidth     =   9495
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6900
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin ZLIBOCX2LibCtl.zlibIF zlibZipper 
      Height          =   285
      Left            =   1215
      OleObjectBlob   =   "frmCompress.frx":0000
      TabIndex        =   12
      Top             =   5085
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1890
      Left            =   4035
      TabIndex        =   8
      Top             =   2130
      Width           =   2790
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   4050
      TabIndex        =   7
      Top             =   4140
      Width           =   2790
   End
   Begin VB.Timer tmrExit 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   360
      Top             =   255
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3915
      TabIndex        =   5
      Top             =   4950
      Width           =   1200
   End
   Begin VB.OptionButton optGame 
      BackColor       =   &H00000000&
      Caption         =   "Game 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   4
      Left            =   2325
      TabIndex        =   4
      Top             =   3810
      Width           =   1125
   End
   Begin VB.OptionButton optGame 
      BackColor       =   &H00000000&
      Caption         =   "Game 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   3
      Left            =   2325
      TabIndex        =   3
      Top             =   3450
      Width           =   1125
   End
   Begin VB.OptionButton optGame 
      BackColor       =   &H00000000&
      Caption         =   "Game 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   2
      Left            =   2325
      TabIndex        =   2
      Top             =   3090
      Width           =   1125
   End
   Begin VB.OptionButton optGame 
      BackColor       =   &H00000000&
      Caption         =   "Game 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   1
      Left            =   2325
      TabIndex        =   1
      Top             =   2730
      Width           =   1125
   End
   Begin VB.OptionButton optGame 
      BackColor       =   &H00000000&
      Caption         =   "Game 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   0
      Left            =   2325
      TabIndex        =   0
      Top             =   2340
      Width           =   1125
   End
   Begin VB.Label lblQuitGame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   4185
      TabIndex        =   11
      Top             =   5895
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblContinue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Continue this game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   3375
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   2790
      TabIndex        =   9
      Top             =   4635
      Width           =   4020
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   3555
      TabIndex        =   6
      Top             =   1350
      Width           =   1680
   End
End
Attribute VB_Name = "frmCompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmCompressStarsDrawn As Boolean  'prevent starfield being drawn over and over
                                        'reset with lblContinue, so next turn the stars are drawn
                                        





Private Sub cmdCompress_Click()
'use the zlib compression OCX to compress the text file for emailing

Dim X As Integer
Dim TempTurnNumber As String
Dim Letter As String
Dim Path As String

'play button sound
PlaySoundEffect "Button5"

Path = Dir1.Path
'make sure path ends with backslash
If Right(Path, 1) <> "\" Then
    Path = Path + "\"
End If

'set tags for player 1 or 2
If Current = 0 Then         'player 1 is saving a game
    Letter = "a"
Else
    Letter = "b"            'player 2 is saving a game
End If

If Current = 1 Then
    TempTurnNumber = Trim(Str(TurnNumber - 1))
Else
    TempTurnNumber = Trim(Str(TurnNumber))
End If

For X = 0 To 4
    If optGame(X).Value = True Then
        GameNumber = X + 1
        'this is the game being saved
        'set the gamenumber variable again to let user continue playing this game

        'ZlibTool1.InputFile = App.Path + "\gameinfo.txt"
        zlibZipper.InputFileName = App.Path + "\gameinfo.txt"
        
        'ZlibTool1.OutputFile = Path + "g" + Trim(Str(x + 1)) + "-" + TempTurnNumber + Letter + ".zlb"
        zlibZipper.OutputFileName = Path + "g" + Trim(Str(X + 1)) + "-" + TempTurnNumber + Letter + ".zlb"
        
        GameName = "g" + Trim(Str(X + 1)) + "-" + TempTurnNumber + Letter + ".zlb"
        
        'compress the file, with save info built into the file name
        'ZlibTool1.Compress
        zlibZipper.Compress
        
    End If
Next X

lblMessage.Caption = "Game Saved As:" + GameName

tmrExit.Enabled = True

cmdCompress.Visible = False
cmdCompress.Enabled = False


End Sub

Private Sub cmdExit_Click()
End

End Sub




Private Sub cmdContinue_Click()
'continue with same game

frmCover.ReadBigFile

'change from one player to the other
If Current = 0 Then
    Current = 1
    Other = 0
ElseIf Current = 1 Then
    Current = 0
    Other = 1
End If
    
If Player(Current).Name = "" Then
    frmPlayer2Setup.Show Modal
End If
    
PlaySoundEffect "Ambient1"

Load frmGameScreen
Unload Me
frmGameScreen.Show


End Sub

Private Sub Drive1_Change()
'standard file navigation - update drive and directories as needed


On Error GoTo DriveError
Dir1.Path = Drive1.Drive
Exit Sub

DriveError:
PlaySoundEffect "Quiet"
MsgBox "Please Choose Another Drive", vbExclamation, "Drive Selection Error"
Drive1.Drive = Dir1.Path
Exit Sub

End Sub

Private Sub Form_Activate()

If frmCompressStarsDrawn = False Then

'draw stars on the screen
    Dim a, X, Y
    For a = 1 To 600
        X = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        Me.PSet (X, Y), vbWhite
    Next a

    'draw darker stars
    Dim grey
    grey = &H808080
    For a = 1 To 800
       X = Int(Rnd * Me.ScaleWidth)
       Y = Int(Rnd * Me.ScaleHeight)
       Me.PSet (X, Y), grey
    Next a
       
    'draw some blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 600
       X = Int(Rnd * Me.ScaleWidth)
       Y = Int(Rnd * Me.ScaleHeight)
       Me.PSet (X, Y), blue
    Next a
    
    frmCompressStarsDrawn = True   'prevent stars from being drawn again and again and again...
    
End If

End Sub

Private Sub Form_Load()

'set to same windowstate as frmgamescreen, which may be windowed
Me.WindowState = frmGameScreen.WindowState   'vbMaximized

Me.Top = frmGameScreen.Top
Me.Left = frmGameScreen.Left

Unload frmGameScreen


'set up explorer-type window
Drive1.Drive = App.Path
Dir1.Path = App.Path

'*** use GameNumber to activate the proper option box
'ie., if player loaded up game 4, this will default to save game 4
optGame(GameNumber - 1).Value = True

End Sub


Private Sub Label2_Click()

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblContinue.ForeColor = vbRed
lblQuitGame.ForeColor = vbRed

End Sub

Private Sub lblContinue_Click()

'continue with same game

frmCover.ReadBigFile

'change from one player to the other
If Current = 0 Then
    Current = 1
    Other = 0
ElseIf Current = 1 Then
    Current = 0
    Other = 1
End If
    
If Player(Current).Name = "" Then
    frmPlayer2Setup.Show Modal
End If
    
PlaySoundEffect "Ambient1"

Load frmGameScreen

'set game screen to same windowstate - ie. windowed
frmGameScreen.WindowState = Me.WindowState

'reset variable to allow stars to be drawn next time
frmCompressStarsDrawn = False
Unload Me


frmGameScreen.Show

End Sub

Private Sub lblContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change colour as mouse moves over label

lblContinue.ForeColor = vbBlue
lblQuitGame.ForeColor = vbRed

End Sub


Private Sub lblQuitGame_Click()


If TurnNumber = 1 Then
        'for some as-yet unknown reason (at least to me),
        'the program will not shut all the way down on turn 1
        'unless I use End - I know it makes no sense, but...
        
        End
    Else
       '***Alternative to using End:
   
        Dim F As Long
   
        'fade form into taskbar
        Me.WindowState = 1
   
        'count forms opened
        For F = Forms.Count - 1 To 0 Step -1
           Unload Forms(F)
        Next F

        'close any open files
        If (Forms.Count = 0) Then Close

        'set all open forms to Nothing
        Set frmGameScreen = Nothing
    End If

'delete the gaminfo.txt file
On Error Resume Next
Kill (App.Path + "\gameinfo.txt")
On Error GoTo 0


End Sub


Private Sub lblQuitGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change colour as mouse moves over label

lblQuitGame.ForeColor = vbBlue
lblContinue.ForeColor = vbRed

End Sub


Private Sub tmrExit_Timer()
'show labels to let user continue game or quit
lblContinue.Enabled = True
lblContinue.Visible = True

lblQuitGame.Enabled = True
lblQuitGame.Visible = True



End Sub


