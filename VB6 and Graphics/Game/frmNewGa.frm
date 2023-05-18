VERSION 5.00
Begin VB.Form frmNewGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "                                    Setup A New Game"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   9465
   ControlBox      =   0   'False
   Icon            =   "FRMNEWGA.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Game"
      Height          =   330
      Left            =   4455
      TabIndex        =   22
      Top             =   5040
      Width           =   1200
   End
   Begin VB.PictureBox picChooseSector 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   5595
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2655
      Width           =   1575
      Begin VB.CommandButton cmdChoice 
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
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdChoice 
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
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdChoice 
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
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdChoice 
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
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Frame fraSelect 
      BackColor       =   &H00000000&
      Caption         =   "Choose Home Planet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1800
      Left            =   2685
      TabIndex        =   12
      Top             =   2340
      Width           =   4725
      Begin VB.TextBox txtNamePlanet 
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
         Height          =   330
         Left            =   540
         TabIndex        =   21
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label lblPlanetName 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name Your Home Planet:"
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
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   990
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Choose sector - Player 2 will start in the opposite corner"
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
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   2595
      End
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00000000&
      Caption         =   "Player 1 Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   2670
      TabIndex        =   11
      Top             =   135
      Width           =   3015
      Begin VB.TextBox txtResources 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1380
         TabIndex        =   20
         Text            =   "15"
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox txtTroops 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1380
         TabIndex        =   18
         Text            =   "10"
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtPlayer1Name 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Top             =   375
         Width           =   1620
      End
      Begin VB.Label lblResources 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Resources:"
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
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label lblTroops 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Troops:"
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
         Left            =   180
         TabIndex        =   17
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label lblInitialValues 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Settings:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   150
         TabIndex        =   16
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Left            =   4455
      TabIndex        =   5
      Top             =   4500
      Width           =   1200
   End
   Begin VB.Frame fraDifficulty 
      BackColor       =   &H00000000&
      Caption         =   "Galaxy Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   5775
      TabIndex        =   0
      Top             =   135
      Width           =   1695
      Begin VB.OptionButton optHard 
         BackColor       =   &H00000000&
         Caption         =   "Large (50)"
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
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   1335
         Width           =   1350
      End
      Begin VB.OptionButton optMedium 
         BackColor       =   &H00000000&
         Caption         =   "Medium (40)"
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
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   885
         Width           =   1380
      End
      Begin VB.OptionButton optEasy 
         BackColor       =   &H00000000&
         Caption         =   "Small (30)"
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
         Height          =   375
         Left            =   150
         TabIndex        =   2
         Top             =   435
         Value           =   -1  'True
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EasyResources, EasyTroops
Dim MedResources, MedTroops
Dim HardResources, HardTroops


'for the galaxy size version:
Dim SmallResources, SmallTroops
Dim LargeResources, LargeTroops

Public Counter As Integer   'to keep track of length of player name

Public SectorChoice As Integer 'to set starting home planet sectors

Private Sub cmdChoice_Click(Index As Integer)
'player chooses home planet, other player gets opposite corner
'also sets planet's owner

Select Case Index
Case 0
    SectorChoice = 0
Case 1
    SectorChoice = 1
Case 2
    SectorChoice = 2
Case 3
    SectorChoice = 3
End Select


If SectorChoice = 0 Then
    Player(Current).HomePlanet = 0
    Planet(0).Owner = Current
    Player(Other).HomePlanet = 49
    Planet(49).Owner = Other
    'set other planets neutral
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral
ElseIf SectorChoice = 1 Then
    Player(Current).HomePlanet = 9
    Planet(9).Owner = Current
    Player(Other).HomePlanet = 40
    Planet(40).Owner = Other
    'neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
ElseIf SectorChoice = 2 Then
    Player(Current).HomePlanet = 40
    Planet(40).Owner = Current
    Player(Other).HomePlanet = 9
    Planet(9).Owner = Other
    'neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
ElseIf SectorChoice = 3 Then
    Player(Current).HomePlanet = 49
    Planet(49).Owner = Current
    Player(Other).HomePlanet = 0
    Planet(0).Owner = Other
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral
End If

'after selecting a starting sector,
'OK button is enabled
cmdStartGame.Enabled = True
cmdStartGame.SetFocus

End Sub


Private Sub cmdExit_Click()
PlaySoundEffect "Quiet"

If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Exiting Game") = vbYes Then
    PlaySoundEffect "Abort"
    'deregister help file
    QuitHelp
    End
End If

End Sub

Private Sub cmdStartGame_Click()

'force player 1 to choose a name
If txtPlayer1Name.Text = "" Then
    PlaySoundEffect "Quiet"
    MsgBox "You must choose a name", vbOKOnly + vbExclamation, "Setup Error"
    txtPlayer1Name.SetFocus
    Exit Sub
End If

If GalaxySize = 30 Then
   'goto procedure to reset homeplanets
   ResetforSmallGalaxy
End If
    
'update player stats
With Player(Current)
    .Name = txtPlayer1Name.Text
    .NumTroops = Val(txtTroops.Text)
    .NumAssaultTroops = 0
    .NumResources = Val(txtResources.Text) - 5
    .NumPlanets = 1
End With

With Player(Other)
.NumTroops = Val(txtTroops.Text)
.NumAssaultTroops = 0
.NumResources = Val(txtResources.Text) - 5
.NumPlanets = 1
End With

'Update boolean values for ships
Dim i As Integer

For i = 0 To 1
   Player(i).Ship(i).Launched = False
Next
        

'set homeplanet name if player does not
If txtNamePlanet = "" Then
    'select one of 5 random names
    Dim Q As Integer
    Q = Int(Rnd * 5) + 1
    Select Case Q
    Case 1
        Planet(Player(Current).HomePlanet).Name = "Glath"
    Case 2
        Planet(Player(Current).HomePlanet).Name = "Pedantak"
    Case 3
        Planet(Player(Current).HomePlanet).Name = "Ritalin IV"
    Case 4
        Planet(Player(Current).HomePlanet).Name = "Belvantas"
    Case 5
        Planet(Player(Current).HomePlanet).Name = "Alt'ngaio"
    End Select
Else
    Planet(Player(Current).HomePlanet).Name = txtNamePlanet.Text
End If

'default player 2 home name
Planet(Player(Other).HomePlanet).Name = "Player2Home"

'update homeplanet troops
Planet(Player(Current).HomePlanet).Troops = Val(txtTroops.Text)
Planet(Player(Other).HomePlanet).Troops = Val(txtTroops.Text)
Planet(Player(Other).HomePlanet).AssaultTroops = 0

'update homeplanet resources
Planet(Player(Current).HomePlanet).Resources = 5
Planet(Player(Other).HomePlanet).Resources = 5

'update homeplanet combatstrength (default of 5)
Dim X, Y
X = Player(Current).HomePlanet
Y = Player(Other).HomePlanet
SetCombatStrength (X)
SetCombatStrength (Y)

'adjust resources around home planets so players don't
'get screwed surrounded by planets with only 1 resource
AdjustResources

'start the game
PlaySoundEffect "Ambient1"

frmGameScreen.Show

Me.Hide



  
End Sub



Private Sub Form_Activate()

DrawStars
    
'clear player name etc - without this, names from last
'game started could appear
txtPlayer1Name.Text = ""
txtNamePlanet.Text = ""

cmdStartGame.Enabled = False

End Sub

Private Sub Form_Load()

'set player one as the first player
Current = 0
Other = 1

TurnNumber = 1

'initialize difficulty values
EasyResources = 25
EasyTroops = 15
MedResources = 15
MedTroops = 15
HardResources = 10
HardTroops = 10

'Large is the default
txtResources.Text = Str(HardResources)
txtTroops.Text = Str(HardTroops)
GalaxySize = 50
optHard.Value = True

'***set default value for GameNumber, otherwise it would be zero and cause an error
GameNumber = 1

'***clear messages and player 2 name, otherwise may 'leak' from last played game
IncomingMessage = ""
Player(0).Name = ""
Player(1).Name = ""


'initialize planet info
InitializePlanets

'set up alien-held planets
SetupAliens



End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'do nothing
Cancel = 1

End Sub




Private Sub optEasy_Click()
'Set easy values
txtResources.Text = Str(EasyResources)
txtTroops.Text = Str(EasyTroops)

GalaxySize = 30

End Sub








Private Sub optHard_Click()
'set hard values
txtResources.Text = Str(HardResources)
txtTroops.Text = Str(HardTroops)

GalaxySize = 50

End Sub





Private Sub optMedium_Click()
'set medium values
txtResources.Text = Str(MedResources)
txtTroops.Text = Str(MedTroops)

GalaxySize = 40

End Sub

Private Sub txtNamePlanet_GotFocus()

txtNamePlanet.Text = ""

End Sub




Private Sub txtPlayer1Name_GotFocus()

txtPlayer1Name.Text = ""

End Sub


Private Sub txtPlayer1Name_KeyDown(KeyCode As Integer, Shift As Integer)
'limit player name to 12 characters
If Counter > 12 Then
    KeyCode = 0
End If

End Sub


Private Sub txtPlayer1Name_KeyPress(KeyAscii As Integer)
'disallow certain keys
'Note: the backspace key erases the contents of the player name text box

Select Case KeyAscii
    Case 8
        'backspace
        txtPlayer1Name.Text = ""
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


Private Sub txtResources_KeyDown(KeyCode As Integer, Shift As Integer)
'disallow input
KeyCode = 0

End Sub


Private Sub txtResources_KeyPress(KeyAscii As Integer)
'disallow input
KeyAscii = 0
End Sub


Private Sub txtTroops_KeyDown(KeyCode As Integer, Shift As Integer)
'disallow input
KeyCode = 0

End Sub


Private Sub txtTroops_KeyPress(KeyAscii As Integer)
'disallow input
KeyAscii = 0

End Sub




Public Sub SetupAliens()
'set up 3 alien-held planets - they expand later in the game...
Dim AlienPlanet1
Dim AlienPlanet2
Dim AlienPlanet3
Dim troops1, troops2, troops3

Randomize

AlienPlanet1 = Int(Rnd * 5) + 12
AlienPlanet2 = Int(Rnd * 5) + 21
AlienPlanet3 = Int(Rnd * 5) + 34

With Planet(AlienPlanet1)
    .Owner = Alien
    .Troops = Int(Rnd * 8) + 2
    .Resources = Int(Rnd * 2) + 1
End With

'call procedure from Declare.bas
SetCombatStrength (AlienPlanet1)


With Planet(AlienPlanet2)
    .Owner = Alien
    .Troops = Int(Rnd * 8) + 2
    .Resources = Int(Rnd * 2) + 1
End With

SetCombatStrength (AlienPlanet2)

With Planet(AlienPlanet3)
    .Owner = Alien
    .Troops = Int(Rnd * 8) + 2
    .Resources = Int(Rnd * 2) + 1
End With

SetCombatStrength (AlienPlanet3)


End Sub

Public Sub ResetforSmallGalaxy()
'hide some of the planets if player 1 chooses a small galaxy

If SectorChoice = 0 Then
    Player(Current).HomePlanet = 11
    Planet(11).Owner = Current
    Player(Other).HomePlanet = 48
    Planet(48).Owner = Other
    'set original planets to neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral
      
ElseIf SectorChoice = 1 Then
    Player(Current).HomePlanet = 18
    Planet(18).Owner = Current
    Player(Other).HomePlanet = 32
    Planet(32).Owner = Other
    'set planets to neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral

ElseIf SectorChoice = 2 Then
    Player(Current).HomePlanet = 32
    Planet(32).Owner = Current
    Player(Other).HomePlanet = 18
    Planet(18).Owner = Other
    'neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral

ElseIf SectorChoice = 3 Then
    Player(Current).HomePlanet = 48
    Planet(48).Owner = Current
    Player(Other).HomePlanet = 11
    Planet(11).Owner = Other
    'neutral
    Planet(9).Owner = Neutral
    Planet(40).Owner = Neutral
    Planet(0).Owner = Neutral
    Planet(49).Owner = Neutral
    
    
End If

End Sub

Public Sub AdjustResources()

'**to readjust the randomly assigned values, to balance play

Select Case Player(Current).HomePlanet
Case 0
    'if any 1's around either homeplanet, change to 2's
    'current player's surrounding planets:
    If Planet(1).Resources = 1 Then Planet(1).Resources = 2
    If Planet(2).Resources = 1 Then Planet(2).Resources = 2
    If Planet(10).Resources = 1 Then Planet(10).Resources = 2
    If Planet(11).Resources = 1 Then Planet(11).Resources = 2
    
    'other player's surrounding planets:
    If Planet(37).Resources = 1 Then Planet(37).Resources = 2
    If Planet(39).Resources = 1 Then Planet(39).Resources = 2
    If Planet(47).Resources = 1 Then Planet(47).Resources = 2
    If Planet(48).Resources = 1 Then Planet(48).Resources = 2

Case 9
    'current player:
    If Planet(7).Resources = 1 Then Planet(7).Resources = 2
    If Planet(8).Resources = 1 Then Planet(8).Resources = 2
    If Planet(18).Resources = 1 Then Planet(18).Resources = 2
    If Planet(19).Resources = 1 Then Planet(19).Resources = 2

    'other player:
    If Planet(31).Resources = 1 Then Planet(31).Resources = 2
    If Planet(32).Resources = 1 Then Planet(32).Resources = 2
    If Planet(41).Resources = 1 Then Planet(41).Resources = 2
    If Planet(42).Resources = 1 Then Planet(42).Resources = 2
    
Case 40
    'current player:
    If Planet(31).Resources = 1 Then Planet(31).Resources = 2
    If Planet(32).Resources = 1 Then Planet(32).Resources = 2
    If Planet(41).Resources = 1 Then Planet(41).Resources = 2
    If Planet(42).Resources = 1 Then Planet(42).Resources = 2

    'other player:
    If Planet(7).Resources = 1 Then Planet(7).Resources = 2
    If Planet(8).Resources = 1 Then Planet(8).Resources = 2
    If Planet(18).Resources = 1 Then Planet(18).Resources = 2
    If Planet(19).Resources = 1 Then Planet(19).Resources = 2
    
Case 49
    If Planet(37).Resources = 1 Then Planet(37).Resources = 2
    If Planet(39).Resources = 1 Then Planet(39).Resources = 2
    If Planet(47).Resources = 1 Then Planet(47).Resources = 2
    If Planet(48).Resources = 1 Then Planet(48).Resources = 2
    
    'other player:
    If Planet(1).Resources = 1 Then Planet(1).Resources = 2
    If Planet(2).Resources = 1 Then Planet(2).Resources = 2
    If Planet(10).Resources = 1 Then Planet(10).Resources = 2
    If Planet(11).Resources = 1 Then Planet(11).Resources = 2
    
Case Else
    'do nothing - player playing in small galaxy
    
End Select

End Sub

Public Sub DrawStars()
Randomize
'draw white stars on the screen
Dim a, X, Y
For a = 1 To 400
    X = Int(Rnd * Me.ScaleWidth)
    Y = Int(Rnd * Me.ScaleHeight)
    Me.PSet (X, Y), vbWhite
Next a
    
'draw dark grey stars
Dim grey
grey = &H808080
For a = 1 To 300
    X = Int(Rnd * Me.ScaleWidth)
    Y = Int(Rnd * Me.ScaleHeight)
    Me.PSet (X, Y), grey
Next a
    
'draw some blue stars
Dim blue
blue = &H800000
For a = 1 To 200
    X = Int(Rnd * Me.ScaleWidth)
    Y = Int(Rnd * Me.ScaleHeight)
    Me.PSet (X, Y), blue
Next a
End Sub

Public Sub InitializePlanets()
'set properties that don't change

Dim a As Integer
For a = 0 To 49
    Planet(a).Owner = Neutral     'default to neutral owner
    Planet(a).Troops = 0
    Planet(a).CombatStrength = 0
    Planet(a).Resources = Int(Rnd * 4) + 1
    Planet(a).Contaminated = False
    Planet(a).HaveMissiles = False
    Planet(a).HaveShields = False
    Planet(a).ImprovedResources = False
    Planet(a).HaveScanner = False
    Planet(a).HaveJammer = False
    Planet(a).Contaminated = False
    Planet(a).NukedResources = False
    Planet(a).Sabotaged = False
Next a


'set up other fixed planet info
'Row One:
With Planet(0)
    .Name = "Regulon"
    .Coordinate = "A1"
    .BackGround = 1
End With
With Planet(1)
    .Name = "Arcturus"
    .Coordinate = "A1"
    .BackGround = 2
End With

With Planet(2)
    .Name = "Pertussis"
    .Coordinate = "B1"
    .BackGround = 3
End With
With Planet(3)
    .Name = "Rigel IV"
    .Coordinate = "B1"
    .BackGround = 4
End With

With Planet(4)
    .Name = "Peklon"
    .Coordinate = "C1"
    .BackGround = 5
End With
With Planet(5)
    .Name = "Smeglor"
    .Coordinate = "C1"
    .BackGround = 4
End With

With Planet(6)
    .Name = "Gabbo"
    .Coordinate = "D1"
    .BackGround = 3
End With
With Planet(7)
    .Name = "Xerxex"
    .Coordinate = "D1"
    .BackGround = 2
End With

With Planet(8)
    .Name = "Bulimus III"
    .Coordinate = "E1"
    .BackGround = 2
End With
With Planet(9)
    .Name = "Irkutsk"
    .Coordinate = "E1"
    .BackGround = 1
End With

'Setup Row Two:
With Planet(10)
    .Name = "Margulon"
    .Coordinate = "A2"
    .BackGround = 2
End With
With Planet(11)
    .Name = "Zigurova"
    .Coordinate = "A2"
    .BackGround = 3
End With

With Planet(12)
    .Name = "Plexus"
    .Coordinate = "B2"
    .BackGround = 4
End With
With Planet(13)
    .Name = "Nexus"
    .Coordinate = "B2"
    .BackGround = 5
End With

With Planet(14)
    .Name = "Plemptor"
    .Coordinate = "C2"
    .BackGround = 4
End With
With Planet(15)
    .Name = "Melnos"
    .Coordinate = "C2"
    .BackGround = 3
End With

With Planet(16)
    .Name = "Dorkasmia"
    .Coordinate = "D2"
    .BackGround = 2
End With
With Planet(17)
    .Name = "Obesios"
    .Coordinate = "D2"
    .BackGround = 1
End With

With Planet(18)
    .Name = "Klathax"
    .Coordinate = "E2"
    .BackGround = 2
End With
With Planet(19)
    .Name = "Troomb"
    .Coordinate = "E2"
    .BackGround = 3
End With

With Planet(20)
    .Name = "Phobos"
    .Coordinate = "A3"
    .BackGround = 4
End With
With Planet(21)
    .Name = "Ebon"
    .Coordinate = "A3"
    .BackGround = 5
End With

With Planet(22)
    .Name = "Orleska"
    .Coordinate = "B3"
    .BackGround = 4
End With
With Planet(23)
    .Name = "Ginsana"
    .Coordinate = "B3"
    .BackGround = 3
End With

With Planet(24)
    .Name = "Algathra"
    .Coordinate = "C3"
    .BackGround = 2
End With
With Planet(25)
    .Name = "Mutalgon"
    .Coordinate = "C3"
    .BackGround = 1
End With

With Planet(26)
    .Name = "Urkel V"
    .Coordinate = "D3"
    .BackGround = 2
End With
With Planet(27)
    .Name = "Cortiska"
    .Coordinate = "D3"
    .BackGround = 3
End With

With Planet(28)
    .Name = "Gath"
    .Coordinate = "E3"
    .BackGround = 4
End With
With Planet(29)
    .Name = "Exevios"
    .Coordinate = "E3"
    .BackGround = 5
End With

With Planet(30)
    .Name = "Ektalbek"
    .Coordinate = "A3"
    .BackGround = 4
End With
With Planet(31)
    .Name = "Intebron"
    .Coordinate = "A4"
    .BackGround = 3
End With

With Planet(32)
    .Name = "Criegor"
    .Coordinate = "B4"
    .BackGround = 2
End With
With Planet(33)
    .Name = "Doolgas"
    .Coordinate = "B3"
    .BackGround = 1
End With

With Planet(34)
    .Name = "Ceti Alpha 5"
    .Coordinate = "C4"
    .BackGround = 2
End With
With Planet(35)
    .Name = "Alteides"
    .Coordinate = "C4"
    .BackGround = 3
End With

With Planet(36)
    .Name = "Baxterion"
    .Coordinate = "D4"
    .BackGround = 4
End With
With Planet(37)
    .Name = "Uilta"
    .Coordinate = "D4"
    .BackGround = 5
End With

With Planet(38)
    .Name = "Goothuzem"
    .Coordinate = "E4"
    .BackGround = 4
End With
With Planet(39)
    .Name = "Eidos"
    .Coordinate = "E4"
    .BackGround = 3
End With

With Planet(40)
    .Name = "Krupaxas"
    .Coordinate = "A5"
    .BackGround = 2
End With
With Planet(41)
    .Name = "Jyzynga"
    .Coordinate = "A5"
    .BackGround = 1
End With

With Planet(42)
    .Name = "Corella"
    .Coordinate = "B5"
    .BackGround = 2
End With
With Planet(43)
    .Name = "Madelos"
    .Coordinate = "B5"
    .BackGround = 3
End With

With Planet(44)
    .Name = "Zantar"
    .Coordinate = "C5"
    .BackGround = 4
End With
With Planet(45)
    .Name = "Solanas III"
    .Coordinate = "C5"
    .BackGround = 5
End With

With Planet(46)
    .Name = "Cerberus"
    .Coordinate = "D5"
    .BackGround = 4
End With
With Planet(47)
    .Name = "Remulak"
    .Coordinate = "D5"
    .BackGround = 3
End With

With Planet(48)
    .Name = "Volenti"
    .Coordinate = "E5"
    .BackGround = 2
End With
With Planet(49)
    .Name = "Rubika"
    .Coordinate = "E5"
    .BackGround = 1
End With

End Sub


