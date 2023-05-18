VERSION 5.00
Object = "{D2D9B7C1-7650-11D1-9481-00A0247B7657}#1.0#0"; "ZLIBOCX2.DLL"
Begin VB.Form frmCover 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "4000 A.D."
   ClientHeight    =   7290
   ClientLeft      =   510
   ClientTop       =   810
   ClientWidth     =   9585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   30
   Icon            =   "frmcover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7290
   ScaleWidth      =   9585
   WindowState     =   2  'Maximized
   Begin ZLIBOCX2LibCtl.zlibIF zlibTester 
      Height          =   375
      Left            =   1935
      OleObjectBlob   =   "frmcover.frx":030A
      TabIndex        =   7
      Top             =   5535
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picStarfield2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   9015
      ScaleHeight     =   2265
      ScaleWidth      =   2760
      TabIndex        =   5
      Top             =   5505
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.PictureBox picStarfield1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   9120
      Picture         =   "frmcover.frx":0342
      ScaleHeight     =   1620
      ScaleWidth      =   2730
      TabIndex        =   6
      Top             =   3390
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Timer tmrOptions 
      Interval        =   500
      Left            =   375
      Top             =   6090
   End
   Begin VB.Label lblVersionNumber 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   7515
      TabIndex        =   8
      Top             =   3330
      Width           =   1545
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   585
      Picture         =   "frmcover.frx":18EC4
      Stretch         =   -1  'True
      Top             =   1335
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      Height          =   270
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   6570
      Width           =   300
   End
   Begin VB.Label lblChoice 
      BackStyle       =   0  'Transparent
      Caption         =   "F3  Quit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   2
      Left            =   5205
      TabIndex        =   2
      Top             =   5460
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblChoice 
      BackStyle       =   0  'Transparent
      Caption         =   "F2  Load A Game In Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   1
      Left            =   5205
      TabIndex        =   1
      Top             =   5010
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Label lblChoice 
      BackStyle       =   0  'Transparent
      Caption         =   "F1  Start A New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   0
      Left            =   5220
      TabIndex        =   0
      Top             =   4590
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Label lblCopyright2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1998-99, Gordon Stewart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   4095
      TabIndex        =   4
      Top             =   6600
      Width           =   2280
   End
   Begin VB.Label lblCopyright1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3810
      TabIndex        =   3
      Top             =   6585
      Width           =   225
   End
End
Attribute VB_Name = "frmCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blue, red
Public StarsDrawn As Boolean



Private Sub Form_Activate()
Randomize

If StarsDrawn = False Then
    'draw white stars on the screen
    Dim a, X, Y
    For a = 1 To 600
        X = Int(Rnd * frmCover.ScaleWidth)
        Y = Int(Rnd * frmCover.ScaleHeight)
        frmCover.PSet (X, Y), vbWhite
    Next a

    'draw darker stars
    Dim grey
    grey = &H808080
    For a = 1 To 800
       X = Int(Rnd * frmCover.ScaleWidth)
       Y = Int(Rnd * frmCover.ScaleHeight)
       frmCover.PSet (X, Y), grey
    Next a
       
    'draw some blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 600
       X = Int(Rnd * frmCover.ScaleWidth)
       Y = Int(Rnd * frmCover.ScaleHeight)
       frmCover.PSet (X, Y), blue
    Next a
        
    StarsDrawn = True   'prevent screen being redrawn later, if user chooses to continue game
End If




End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyF1
    'start a new game
    
    'hide the copyright labels
    lblCopyright1.Visible = False
    lblCopyright2.Visible = False
    Shape1.Visible = False   'the circle around the "c"

    'start newgame
    frmNewGame.Show

Case vbKeyF2
    'load a saved game
    'hide the copyright labels
    lblCopyright1.Visible = False
    lblCopyright2.Visible = False
    Shape1.Visible = False   'the circle around the "c"
    
    'show gameselect form
    frmSelectGame.Show Modal

    ReadBigFile
    'change from one player to another, depending on
    'value of "current" as read in from saved game file
      If Current = 0 Then
         Current = 1
         Other = 0
      ElseIf Current = 1 Then
         Current = 0
         Other = 1
      End If
      
    'show player 2 setup form if this is player 2's first turn, or if player
    '2 hasn't entered a name yet
    If Player(Current).Name = "" Then
       frmPlayer2Setup.Show Modal
    End If
    
    PlaySoundEffect "Ambient1"
    'show main game screen
    frmGameScreen.Show

Case vbKeyF3
    'quit
    PlaySoundEffect "Abort"
    End

Case Else
    'ignore other keystrokes
    KeyCode = 0
End Select

End Sub

Private Sub Form_Load()

'see if game is already running
If App.PrevInstance = True Then
    Call MsgBox("This program is already running!", vbExclamation)
    End
End If

'read config file
On Error Resume Next
    
    ReadConfigFile

    SoundOn = DefaultGameSound
    Me.WindowState = DefaultGameSize

    If Me.WindowState = 0 Then
        'readjust location
        Me.Top = 0
        Me.Left = 0
    End If
    
On Error GoTo 0

PlaySoundEffect "Intro"


blue = &HFF0000
red = &HFF&

'DO NOT try to pre-load the game screen to speed it up
'It doesn't work, and screws up everything!

'position the version number
'lblVersionNumber.Top = Me.Height - 50
'lblVersionNumber.Left = Me.Width - 50

    

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'reset label colours to red

Dim i
For i = 0 To 2
    lblChoice(i).ForeColor = red
Next i

End Sub







Private Sub lblChoice_Click(Index As Integer)
Select Case Index
Case 0
    'start a new game
    'load the new game screen
    Randomize
    
    'cover up title and labels so they don't show through later
    With picStarfield1
        .Left = 550
        .Top = 1300
        .Height = 1900
        .Width = 8500
    End With
    
    With picStarfield2
        .Left = 3435
        .Top = 4125
        .Height = 3105
        .Width = 6435
        .Picture = picStarfield1.Picture
    End With
    
    'hide the labels
    lblCopyright1.Visible = False
    lblCopyright2.Visible = False
    Shape1.Visible = False  'the circle around the "c"
    
    lblVersionNumber.Visible = False
    
    'load form to start new game
    frmNewGame.Show
    
Case 1
    'Load a saved game
    'hide the labels
    lblCopyright1.Visible = False
    lblCopyright2.Visible = False
    Shape1.Visible = False   'the circle around the "c"
    
    lblVersionNumber.Visible = False
    
    frmSelectGame.Show Modal
    
    'player cancels load
    If LoadCancelled = True Then
        'program returns to frmCover
        LoadCancelled = False
        Exit Sub
    End If
    
    'read in saved game info from file
    ReadBigFile
    
    'delete the gameinfo.txt file
    On Error Resume Next
    Kill (App.Path + "\gameinfo.txt")
    On Error GoTo 0
    
    
    
    'change from one player to another, depending on
    'value of "current" as read in from players.txt
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
    frmGameScreen.Show
    '**********************Unload frmCover
    
Case 2   'Quit Game
    'exit to system
    PlaySoundEffect "Abort"
    
    'deregister help file
    QuitHelp
    
    'End
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

End Select
  

End Sub

Private Sub lblChoice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Counter
  For Counter = 0 To 2
    lblChoice(Counter).ForeColor = red
  Next Counter
  
  lblChoice(Index).ForeColor = blue
End Sub















Private Sub tmrOptions_Timer()
'wait a bit before displaying labels, forcing the player to admire the logo before playing...

Dim X
For X = 0 To 2
    lblChoice(X).Visible = True
Next X

tmrOptions.Enabled = False

End Sub







Public Sub ReadBigFile()
'read in the file with saved game info

Dim Filename As String
Dim ShortName As String
Dim i As Integer

'set up error trapping if file not in directory
On Error GoTo FileError

ShortName = "\gameinfo.txt"
Filename = App.Path & ShortName

'get a free file number
gFileNum = FreeFile

'open the file
Open Filename For Input As gFileNum

'read galaxysize
Input #gFileNum, GalaxySize

'read the planet data
For i = 0 To 49
  Input #gFileNum, Planet(i).Name, Planet(i).Owner, Planet(i).Troops, _
  Planet(i).AssaultTroops, Planet(i).CombatStrength, Planet(i).Coordinate, _
  Planet(i).Resources, Planet(i).HaveMissiles, Planet(i).HaveShields, _
  Planet(i).ImprovedResources, Planet(i).HaveScanner, Planet(i).BackGround, _
  Planet(i).HaveJammer, Planet(i).BioRocketETA, Planet(i).Contaminated, Planet(i).NukedResources, _
  Planet(i).Sabotaged, Planet(i).SabotageReduction, Planet(i).SabotagedFactory, Planet(i).Damaged, _
  Planet(i).BioFailed
Next

'read the player data
For i = 0 To 1
  Input #gFileNum, Current, TurnNumber, Player(i).Name, Player(i).NumTroops, Player(i).NumAssaultTroops, Player(i).NumPlanets, Player(i).NumResources, _
  Player(i).HomePlanet, Player(i).Message1Given, Player(i).Message2Given, Player(i).WasBig, Player(i).Missile1ResearchDone, Player(i).Missile1Researched, _
  Player(i).Missile2ResearchDone, Player(i).Missile2Researched, Player(i).ShieldResearchDone, Player(i).ShieldResearched, _
  Player(i).LaserResearchDone, Player(i).LaserResearched, Player(i).PlasmaResearchDone, Player(i).PlasmaResearched, Player(i).MechResearchDone, Player(i).MechResearched, _
  Player(i).BioRocketResearchDone, Player(i).BioRocketResearched, Player(i).LongBioResearchDone, Player(i).LongBioResearched, Player(i).ShipShield1ResearchDone, Player(i).ShipShield1Researched, _
  Player(i).ShipShield2ResearchDone, Player(i).ShipShield2Researched, Player(i).BigShipResearchDone, Player(i).BigShipResearched, Player(i).UltraWarpResearchDone, Player(i).UltraWarpResearched, _
  Player(i).CloakingResearchDone, Player(i).CloakingResearched, Player(i).ResourceResearchDone, Player(i).ResourcesResearched, Player(i).BioCleanupResearchDone, Player(i).BioCleanupResearched, _
  Player(i).RegenerationResearchDone, Player(i).RegenerationResearched, Player(i).ScannerResearchDone, Player(i).ScannerResearched, Player(i).DeepScannerResearchDone, Player(i).DeepScannerResearched, _
  Player(i).JammerResearchDone, Player(i).JammerResearched, Player(i).WarpScannerResearchDone, Player(i).WarpScannerResearched
Next

'read the ship data
For i = 0 To 1
  Input #gFileNum, Player(0).Ship(i).Launched, Player(0).Ship(i).HaveCloakingDevice, _
  Player(0).Ship(i).Troops, Player(0).Ship(i).AssaultTroops, _
  Player(0).Ship(i).CombatStrength, Player(0).Ship(i).WarpPosition, _
  Player(0).Ship(i).Coordinate, Player(0).Ship(i).CenterX, _
  Player(0).Ship(i).CenterY, Player(0).Ship(i).ShipNumber, Player(0).Ship(i).Sabotage, _
  Player(1).Ship(i).Launched, Player(1).Ship(i).HaveCloakingDevice, _
  Player(1).Ship(i).Troops, Player(1).Ship(i).AssaultTroops, _
  Player(1).Ship(i).CombatStrength, Player(1).Ship(i).WarpPosition, _
  Player(1).Ship(i).Coordinate, Player(1).Ship(i).CenterX, _
  Player(1).Ship(i).CenterY, Player(1).Ship(i).ShipNumber, Player(1).Ship(i).Sabotage
Next

'read in the general data
Input #gFileNum, IncomingMessage, Player(Other).WasBig

'turn off previous error handling
On Error GoTo 0

'allow older versions of game to be read, and notify player that their opponent
'needs to get the new version of the game

'the following info won't be in saved game files created by older versions of the game!

On Error GoTo NewDataError
'read captured planets data
Input #gFileNum, NumPlanetsCaptured

For i = 0 To 49
    Input #gFileNum, Planet(i).Captured
Next i

Input #gFileNum, NumFailedInvasions
For i = 0 To 49
    Input #gFileNum, Planet(i).FailedInvasion, Planet(i).FailedInvasionTroopLosses, Planet(i).FailedInvasionMechLosses
Next i

'turn off that error handler
On Error GoTo 0

'close the file
Close gFileNum


Exit Sub

FileError:
    'file not in the directory
    PlaySoundEffect "Quiet"
    MsgBox "File not found.", vbOKOnly + vbExclamation, "File Error"
    End
    Exit Sub

NewDataError:
    'error reading from older version
    'close the file
    Dim Msg
    Msg = "The game information file being read by the program is from" + Chr(13)
    Msg = Msg + "an older version of the game." + Chr(13) + Chr(13)
    Msg = Msg + "Please ask your opponent to upgrade to the latest version" + Chr(13)
    Msg = Msg + "of 4000 A.D., as they will not be able to load the file you " + Chr(13) + "send to them." + Chr(13) + Chr(13)
    Msg = Msg + "Upgrades and more available on the internet at:" + Chr(13) + "www.interlog.com/~gordons/4000ad.html"
    
    PlaySoundEffect "Quiet"
    MsgBox Msg, vbOKOnly + vbInformation, "Reading Old Format"
    
    Close gFileNum
    
    Exit Sub
    
End Sub

Private Sub ZlibTester_Progress(ByVal percent_complete As Integer)
'this is included to test if the zlibtool.ocx file was properly registered
'during installation - see details in the sub_main procedure in
'the Declare.bas file

End Sub


Private Sub zlibITester_Progress(ByVal percent_complete As Integer)
'this is included to test if the zlibtool.ocx file was properly registered
'during installation - see details in the sub_main procedure in
'the Declare.bas file

End Sub



Public Sub ReadConfigFile()
Dim Filename As String
Dim ShortName As String

'set up error trapping
On Error GoTo ErrorHandler


ShortName = "\4000cfg.txt"
Filename = App.Path & ShortName

'get a free file number
gFileNum = FreeFile

'open the file
Open Filename For Input As gFileNum

'read galaxysize
Input #gFileNum, DefaultGameSound, DefaultGameSize



'close the file
Close gFileNum


Exit Sub

ErrorHandler:
    'file isn't there - could be missing/deleted, or player is reading game file from
    'an older version of the game, where no 4000cfg.txt file is written
    
    'close the file and continue
    Close gFileNum

    Exit Sub
    
End Sub
