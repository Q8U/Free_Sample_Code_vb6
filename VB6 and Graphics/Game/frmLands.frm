VERSION 5.00
Begin VB.Form frmLandscape 
   BorderStyle     =   0  'None
   Caption         =   "Landscape - landing, attacking"
   ClientHeight    =   4170
   ClientLeft      =   1050
   ClientTop       =   1485
   ClientWidth     =   7545
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
   ForeColor       =   &H0000FFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   945
      ScaleHeight     =   180
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   1455
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Timer tmrToolTips 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   285
      Top             =   1335
   End
   Begin VB.Timer tmrFire 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   6825
      Top             =   2640
   End
   Begin VB.Timer Timer 
      Interval        =   1500
      Left            =   1290
      Top             =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   315
      Left            =   6120
      TabIndex        =   0
      Top             =   3630
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgWreckedJammer 
      Height          =   555
      Left            =   6660
      Picture         =   "FRMLANDS.frx":0000
      Stretch         =   -1  'True
      Top             =   1455
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image5 
      Height          =   15
      Left            =   2820
      Top             =   2235
      Width           =   30
   End
   Begin VB.Image Image2 
      Height          =   4200
      Left            =   7350
      Picture         =   "FRMLANDS.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image imgMissiles2 
      Height          =   600
      Left            =   4980
      Picture         =   "FRMLANDS.frx":41EC
      Stretch         =   -1  'True
      Top             =   3285
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   165
      Picture         =   "FRMLANDS.frx":50B6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7290
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   195
      Picture         =   "FRMLANDS.frx":9388
      Stretch         =   -1  'True
      Top             =   4005
      Width           =   7290
   End
   Begin VB.Image Image1 
      Height          =   4170
      Left            =   0
      Picture         =   "FRMLANDS.frx":D65A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   210
   End
   Begin VB.Image imgFire 
      Height          =   480
      Index           =   2
      Left            =   6780
      Picture         =   "FRMLANDS.frx":10F7C
      Top             =   2115
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFire 
      Height          =   480
      Index           =   1
      Left            =   6405
      Picture         =   "FRMLANDS.frx":11846
      Top             =   2595
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFire 
      Height          =   480
      Index           =   0
      Left            =   6780
      Picture         =   "FRMLANDS.frx":12110
      Top             =   3045
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgScannerFire 
      Height          =   480
      Left            =   1860
      Picture         =   "FRMLANDS.frx":129DA
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMissileFire 
      Height          =   480
      Left            =   4965
      Picture         =   "FRMLANDS.frx":132A4
      Top             =   2685
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   7
      Left            =   3120
      Picture         =   "FRMLANDS.frx":13B6E
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   6
      Left            =   2670
      Picture         =   "FRMLANDS.frx":14438
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   5
      Left            =   2265
      Picture         =   "FRMLANDS.frx":14D02
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   4
      Left            =   1860
      Picture         =   "FRMLANDS.frx":155CC
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   3
      Left            =   1455
      Picture         =   "FRMLANDS.frx":15E96
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   2
      Left            =   1050
      Picture         =   "FRMLANDS.frx":16760
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   1
      Left            =   660
      Picture         =   "FRMLANDS.frx":1702A
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgMech 
      Height          =   465
      Index           =   0
      Left            =   270
      Picture         =   "FRMLANDS.frx":178F4
      Stretch         =   -1  'True
      Top             =   3495
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   9
      Left            =   3270
      Picture         =   "FRMLANDS.frx":181BE
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   8
      Left            =   2955
      Picture         =   "FRMLANDS.frx":18A88
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   7
      Left            =   2595
      Picture         =   "FRMLANDS.frx":19352
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   6
      Left            =   2250
      Picture         =   "FRMLANDS.frx":19C1C
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   5
      Left            =   1950
      Picture         =   "FRMLANDS.frx":1A4E6
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   4
      Left            =   1620
      Picture         =   "FRMLANDS.frx":1ADB0
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   3
      Left            =   1305
      Picture         =   "FRMLANDS.frx":1B67A
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   2
      Left            =   960
      Picture         =   "FRMLANDS.frx":1BF44
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   1
      Left            =   660
      Picture         =   "FRMLANDS.frx":1C80E
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgTroop 
      Height          =   315
      Index           =   0
      Left            =   330
      Picture         =   "FRMLANDS.frx":1D0D8
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgAssault 
      Height          =   495
      Left            =   2235
      Picture         =   "FRMLANDS.frx":1D9A2
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgTroops 
      Height          =   480
      Left            =   1800
      Picture         =   "FRMLANDS.frx":1E26C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgBarracks4 
      Height          =   465
      Left            =   3000
      Picture         =   "FRMLANDS.frx":1EB36
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgBarracks3 
      Height          =   465
      Left            =   3990
      Picture         =   "FRMLANDS.frx":1F400
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgFactory 
      Height          =   645
      Left            =   5700
      Picture         =   "FRMLANDS.frx":1FCCA
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgCommandCentre 
      Height          =   630
      Left            =   3360
      Picture         =   "FRMLANDS.frx":20594
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   780
   End
   Begin VB.Image imgBarracks1 
      Height          =   555
      Left            =   2520
      Picture         =   "FRMLANDS.frx":21D16
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgScanner 
      Height          =   555
      Left            =   1680
      Picture         =   "FRMLANDS.frx":225E0
      Stretch         =   -1  'True
      Top             =   2220
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgBarracks2 
      Height          =   540
      Left            =   4200
      Picture         =   "FRMLANDS.frx":22EAA
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgMissiles 
      Height          =   720
      Left            =   4860
      Picture         =   "FRMLANDS.frx":23774
      Stretch         =   -1  'True
      Top             =   2685
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgBackGround5 
      Height          =   465
      Left            =   3555
      Picture         =   "FRMLANDS.frx":2463E
      Stretch         =   -1  'True
      Top             =   900
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBackGround4 
      Height          =   465
      Left            =   2700
      Picture         =   "FRMLANDS.frx":45E10
      Stretch         =   -1  'True
      Top             =   885
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBackGround3 
      Height          =   465
      Left            =   4365
      Picture         =   "FRMLANDS.frx":675E2
      Stretch         =   -1  'True
      Top             =   375
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBackGround2 
      Height          =   510
      Left            =   3540
      Picture         =   "FRMLANDS.frx":8A624
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgBackGround1 
      Height          =   465
      Left            =   2715
      Picture         =   "FRMLANDS.frx":AB9BE
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgSatellite 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6225
      Picture         =   "FRMLANDS.frx":D007C
      Stretch         =   -1  'True
      Top             =   285
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgJammer 
      Height          =   615
      Left            =   4980
      Picture         =   "FRMLANDS.frx":D17FE
      Stretch         =   -1  'True
      Top             =   1980
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "frmLandscape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Counter As Integer

Public MissileFlag As Boolean       'used in cmdOK, to get rid of missiles on planet
Public SatelliteFlag As Boolean     'to get rid of planetary shield
Public ScannerFlag As Boolean       'to get rid of scanner
Public JammerFlag As Boolean        'get rid of jammer


Public MouseX, MouseY               'for tooltips
Public ObjectSelected As String     'set when mouse lingers over missiles, scanner, jammer, buildings



Private Sub cmdOK_Click()
'flags set in tmrFire procedure...

'get rid of missiles
If MissileFlag Then
    Planet(ActivePlanet).HaveMissiles = False
End If

'get rid of shields
If SatelliteFlag Then
    Planet(ActivePlanet).HaveShields = False
End If

'get rid of scanner
If ScannerFlag Then
    Planet(ActivePlanet).HaveScanner = False
End If

'get rid of jammer
If JammerFlag Then
    Planet(ActivePlanet).HaveJammer = False
End If

Unload frmLandscape

End Sub





Private Sub Form_Activate()
SetBackGround

Select Case Planet(ActivePlanet).Owner

Case Neutral
    'show blank landscape and troops landing there
    ShowTroops
    
Case Current
    'player is viewing planet
    'turn on some of the buildings, if necessary
    ShowTroops
    ShowBuildings
Case Other
    'showing results of attack
    If AttackStrength > DefenceStrength Then
        'current player wins
        ShowVictory
      ElseIf AttackStrength <= DefenceStrength Then
        'current player loses
        ShowDefeat
    End If
Case Alien
    'showing results of attack
    If AttackStrength > DefenceStrength Then
        'current player wins
        ShowVictory
      ElseIf AttackStrength <= DefenceStrength Then
        'current player loses
        ShowDefeat
    End If
End Select



End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        'show quick help
        ShowQuickHelp
End Select
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get rid of tooltips if mouse moves off one of the buildings etc.

picToolTip.Visible = False
tmrToolTips.Enabled = False
picToolTip.CurrentX = 0
picToolTip.CurrentY = 0

'erase text in previous tooltip
picToolTip.Print "                     "

End Sub

Private Sub imgBarracks1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Barracks"
tmrToolTips.Enabled = True
MouseX = imgBarracks1.Left - 250
MouseY = imgBarracks1.Top - 175
End Sub


Private Sub imgBarracks2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Barracks"
tmrToolTips.Enabled = True
MouseX = imgBarracks2.Left - 250
MouseY = imgBarracks2.Top - 175
End Sub


Private Sub imgBarracks3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Barracks"
tmrToolTips.Enabled = True
MouseX = imgBarracks3.Left - 250
MouseY = imgBarracks3.Top - 175
End Sub


Private Sub imgBarracks4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Barracks"
tmrToolTips.Enabled = True
MouseX = imgBarracks4.Left - 250
MouseY = imgBarracks4.Top - 175
End Sub


Private Sub imgCommandCentre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Command Centre"
tmrToolTips.Enabled = True
MouseX = imgCommandCentre.Left - 325
MouseY = imgCommandCentre.Top - 250



End Sub


Private Sub imgFactory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Factory"
tmrToolTips.Enabled = True
MouseX = imgFactory.Left + 125
MouseY = imgFactory.Top - 250

End Sub


Private Sub imgJammer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Jammer"
tmrToolTips.Enabled = True
MouseX = imgJammer.Left - 250
MouseY = imgJammer.Top - 250

End Sub


Private Sub imgmissiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Missiles"
tmrToolTips.Enabled = True
MouseX = imgMissiles.Left + 125
MouseY = imgMissiles.Top + 600
End Sub


Private Sub imgSatellite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Planetary Shield"
tmrToolTips.Enabled = True
MouseX = imgSatellite.Left - 400
MouseY = imgSatellite.Top + 500
End Sub


Private Sub imgScanner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ObjectSelected = "Scanner"
tmrToolTips.Enabled = True
MouseX = imgScanner.Left - 325
MouseY = imgScanner.Top - 250

End Sub


Private Sub Timer_Timer()
'slight delay before button appears to unload form

cmdOK.Visible = True
cmdOK.Enabled = True

End Sub



Public Sub SetBackGround()
'set background according to planet attributes

Select Case Planet(ActivePlanet).BackGround

Case 1
    frmLandscape.Picture = imgBackGround1.Picture
Case 2
    frmLandscape.Picture = imgBackGround2.Picture
Case 3
    frmLandscape.Picture = imgBackGround3.Picture
Case 4
    frmLandscape.Picture = imgBackGround4.Picture
Case 5
    frmLandscape.Picture = imgBackGround5.Picture
End Select

End Sub

Public Sub ShowBuildings()
'show appropriate bldgs for #troops, research etc.
'command building is always shown, even if no troops/mechs

If Planet(ActivePlanet).Troops > 5 Or Planet(ActivePlanet).AssaultTroops > 5 Then
    imgBarracks1.Visible = True
End If

If Planet(ActivePlanet).Troops > 10 Or Planet(ActivePlanet).AssaultTroops > 10 Then
    imgBarracks2.Visible = True
End If

If Planet(ActivePlanet).Troops > 20 Or Planet(ActivePlanet).AssaultTroops > 15 Then
    imgBarracks3.Visible = True
End If

If Planet(ActivePlanet).Troops > 30 Or Planet(ActivePlanet).AssaultTroops > 20 Then
    imgBarracks4.Visible = True
End If

'set buildings for research done and improvements paid for
If Planet(ActivePlanet).HaveShields Then
    imgSatellite.Visible = True
End If

If Planet(ActivePlanet).HaveMissiles Then
    If Player(Current).Missile2Researched Then
        imgMissiles.Picture = imgMissiles2.Picture
        imgMissiles.Visible = True
    Else
        imgMissiles.Visible = True
    End If
End If

If Planet(ActivePlanet).ImprovedResources Then
    imgFactory.Visible = True
End If

If Planet(ActivePlanet).HaveScanner Then
    imgScanner.Visible = True
End If

If Planet(ActivePlanet).HaveJammer Then
    imgJammer.Visible = True
End If

End Sub

Public Sub ShowVictory()
Dim Msg(5) As String
Dim choice As Integer

Msg(0) = "Your Forces are Victorious!"
Msg(1) = "A Rousing Victory!"
Msg(2) = Player(Other).Name + "'s forces put to the sword!"
Msg(3) = "The Invasion Was A Complete Success!"
Msg(4) = "Enemy Forces Annihiliated!"
Msg(5) = "The Planet Is Yours!"

'Current player wins
'show ship landed

'show buildings
ShowBuildings
'test
ShowTroops

'show satellite/missile icons burning
tmrFire.Enabled = True

'show message
frmLandscape.CurrentX = 300
frmLandscape.CurrentY = 250
If Planet(ActivePlanet).Owner = Other Then
    choice = Int(Rnd * 5) + 1
    frmLandscape.FontSize = 12
    frmLandscape.Print Msg(choice)
ElseIf Planet(ActivePlanet).Owner = Alien Then
    frmLandscape.FontSize = 12
    frmLandscape.Print "Death to the Melnikons!"
End If

frmLandscape.Print
frmLandscape.FontSize = 10

If TroopLosses > 0 Then
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Troop Losses: ", TroopLosses & "%"
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Troops Remaining: ", Planet(ActivePlanet).Troops
ElseIf TroopLosses = 0 And Planet(ActivePlanet).Troops > 0 Then
    frmLandscape.CurrentX = 700
    frmLandscape.Print "No Troops Lost"
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Troops Available: ", Planet(ActivePlanet).Troops
End If

If AssaultLosses > 0 Then
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Assault Troop Losses: ", AssaultLosses & "%"
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Mechs Remaining: ", Planet(ActivePlanet).AssaultTroops
ElseIf AssaultLosses = 0 And Planet(ActivePlanet).AssaultTroops > 0 Then
    frmLandscape.CurrentX = 700
    frmLandscape.Print "No Mechs Lost"
    frmLandscape.CurrentX = 700
    frmLandscape.Print "Mechs Available: ", Planet(ActivePlanet).AssaultTroops
End If

End Sub

Public Sub ShowDefeat()
'player lost the battle
'choose one of a number of messages

Dim Msg(5) As String
Dim choice As Integer

Msg(0) = "Your Forces Are In Ruin!"
Msg(1) = "A Crushing Defeat - All Troops Lost!"
Msg(2) = Player(Other).Name + "'s Forces Overwhelm Your Attack Squad!"
Msg(3) = "The Invasion Was A Total Failure!"
Msg(4) = "Your Forces Were Annihiliated!"
Msg(5) = "A Humiliating Rout!"



'Current player loses
ShowBuildings
ShowTroops

'show message at top
frmLandscape.CurrentX = 300
frmLandscape.CurrentY = 200

If Planet(ActivePlanet).Owner = Other Then
    choice = Int(Rnd * 5) + 1
    frmLandscape.FontSize = 12
    Print Msg(choice)
ElseIf Planet(ActivePlanet).Owner = Alien Then
    frmLandscape.FontSize = 12
    Print "Your Forces Were Wiped Out By The Melnikons!"
End If

frmLandscape.FontSize = 10

Print
frmLandscape.CurrentX = 300

Print "Last transmission from ship's sensors returned "
frmLandscape.CurrentX = 300
Print "the following planetary data:"

frmLandscape.CurrentX = 700
Print "Number of Troops: ", Planet(ActivePlanet).Troops
frmLandscape.CurrentX = 700
Print "Number of Mechs: ", Planet(ActivePlanet).AssaultTroops
frmLandscape.CurrentX = 700
Print "Missile Defenses: ",
If Planet(ActivePlanet).HaveMissiles Then
    Print "Active"
Else
    Print "None"
End If
frmLandscape.CurrentX = 700
Print "Planetary Shields: ",
If Planet(ActivePlanet).HaveShields Then
    Print "Active"
Else
    Print "None"
End If


End Sub

Public Sub ShowTroops()
'turn on images of troops/mechs

Dim NumTroopBoxes
Dim NumAssaultBoxes
Dim Counter

NumTroopBoxes = Int(Planet(ActivePlanet).Troops / 4)
NumAssaultBoxes = Int(Planet(ActivePlanet).AssaultTroops / 4)

If Planet(ActivePlanet).Troops = 0 Then
    NumTroopBoxes = 0
End If
If Planet(ActivePlanet).AssaultTroops = 0 Then
    NumAssaultBoxes = 0
End If

If Planet(ActivePlanet).Troops > 0 And Planet(ActivePlanet).Troops < 4 Then
    NumTroopBoxes = 1
End If

If NumTroopBoxes > 10 Then
   NumTroopBoxes = 10
End If

If Planet(ActivePlanet).AssaultTroops > 0 And Planet(ActivePlanet).AssaultTroops < 4 Then
    NumAssaultBoxes = 1
    
End If

If NumAssaultBoxes > 8 Then
    NumAssaultBoxes = 8
End If

If NumTroopBoxes > 0 Then
    For Counter = 0 To NumTroopBoxes - 1
        imgTroop(Counter).Visible = True
    Next
End If

If NumAssaultBoxes > 0 Then
    For Counter = 0 To NumAssaultBoxes - 1
        imgMech(Counter).Visible = True
    Next
End If

End Sub

Private Sub tmrFire_Timer()
'rotate the fire icons
'only if the buildings are there...

If Planet(ActivePlanet).HaveMissiles Then
    MissileFlag = True
    imgMissileFire.Visible = True
    imgMissileFire.Picture = imgFire(Counter).Picture
End If

Counter = Counter + 1

If Counter = 3 Then
    Counter = 0
End If

If Planet(ActivePlanet).HaveShields Then
    'start really lame animation to show satellite drifting away
    SatelliteFlag = True
    imgSatellite.Top = imgSatellite.Top + 15
    imgSatellite.Left = imgSatellite.Left + 25
    If imgSatellite.Left > 7065 Then
        'image has moved off the edge of the screen
        imgSatellite.Visible = False
    End If
End If

If Planet(ActivePlanet).HaveScanner Then
    ScannerFlag = True
    imgScannerFire.Visible = True
    imgScannerFire.Picture = imgFire(Counter).Picture
End If

If Planet(ActivePlanet).HaveJammer Then
    JammerFlag = True
    imgJammer.Picture = imgWreckedJammer.Picture
End If

End Sub


Private Sub tmrToolTips_Timer()
'if mouse lingers over an object, show tooltip
ShowTip

End Sub



Public Sub ShowTip()
'sort-of tooltips when mouse moved over buildings

Select Case ObjectSelected

Case "Barracks"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 685
    End With
    picToolTip.Picture = LoadPicture()

    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Barracks"

Case "Factory"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 590
    End With
    picToolTip.Picture = LoadPicture()
    
    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Factory"

Case "Jammer"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 1225
    End With
    picToolTip.Picture = LoadPicture()

    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Jamming Device"

Case "Missiles"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 920
    End With
    picToolTip.Picture = LoadPicture()
    
    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Missile Base"


Case "Planetary Shield"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 1200
    End With
    picToolTip.Picture = LoadPicture()
    
    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Planetary Shield"

Case "Scanner"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 685
    End With
    picToolTip.Picture = LoadPicture()
    
    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Scanner"
    
Case "Command Centre"
    With picToolTip
        .Top = MouseY
        .Left = MouseX
        .Visible = True
        .CurrentX = 0
        .CurrentY = 0
        .Width = 1275
     End With
    picToolTip.Picture = LoadPicture()
    
    picToolTip.CurrentX = 0
    picToolTip.CurrentY = 0
    picToolTip.Print "Command Centre"
    
End Select

End Sub
