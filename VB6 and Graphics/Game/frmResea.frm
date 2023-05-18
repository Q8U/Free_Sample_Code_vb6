VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmResearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technology Research Centre"
   ClientHeight    =   5460
   ClientLeft      =   495
   ClientTop       =   1425
   ClientWidth     =   8850
   ControlBox      =   0   'False
   Icon            =   "FRMRESEA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   8850
   Begin VB.PictureBox Picture1 
      Height          =   3660
      Left            =   4695
      Picture         =   "FRMRESEA.frx":0ECA
      ScaleHeight     =   3600
      ScaleWidth      =   3975
      TabIndex        =   15
      Top             =   135
      Width           =   4035
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Select Technology To Research:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   4785
      TabIndex        =   2
      Top             =   255
      Width           =   3825
      Begin VB.OptionButton optCloaking 
         Caption         =   "Cloaking Device"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1890
         TabIndex        =   21
         Top             =   2820
         Width           =   1680
      End
      Begin VB.OptionButton optJammer 
         Caption         =   "Anti-Scanning Jammer"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1305
         TabIndex        =   20
         Top             =   1290
         Width           =   2115
      End
      Begin VB.OptionButton optDeepScanner 
         Caption         =   "Deep-Space Scanner"
         Enabled         =   0   'False
         Height          =   210
         Left            =   1305
         TabIndex        =   19
         Top             =   1005
         Width           =   1800
      End
      Begin VB.OptionButton optScanner 
         Caption         =   "Space Scanner"
         Height          =   315
         Left            =   765
         TabIndex        =   18
         Top             =   660
         Width           =   2115
      End
      Begin VB.OptionButton optFastShips 
         Caption         =   "Ultra-Warp Ship Engines"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1335
         TabIndex        =   17
         Top             =   2505
         Width           =   2040
      End
      Begin VB.OptionButton optBigShips 
         Caption         =   "Expand Ship Capacity"
         Height          =   285
         Left            =   750
         TabIndex        =   16
         Top             =   2220
         Width           =   2040
      End
      Begin VB.OptionButton optResources 
         Caption         =   "Improved Resource Production"
         Height          =   285
         Left            =   750
         TabIndex        =   6
         Top             =   3150
         Width           =   2520
      End
      Begin VB.OptionButton optShips 
         Caption         =   "Ship Weaponry"
         Height          =   285
         Left            =   750
         TabIndex        =   5
         Top             =   1890
         Width           =   2040
      End
      Begin VB.OptionButton optAssault 
         Caption         =   "Mechanized Assault Troops"
         Height          =   270
         Left            =   765
         TabIndex        =   4
         Top             =   1560
         Width           =   2250
      End
      Begin VB.OptionButton optPlanetary 
         Caption         =   "Planetary Shields/Missiles"
         Height          =   315
         Left            =   750
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   6795
      TabIndex        =   1
      Top             =   4455
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5505
      TabIndex        =   0
      Top             =   4455
      Width           =   1200
   End
   Begin VB.Frame fraAmount 
      Caption         =   "Research Teams Allocated:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   75
      TabIndex        =   7
      Top             =   2985
      Width           =   4515
      Begin VB.TextBox txtTotal 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1650
         Width           =   585
      End
      Begin VB.TextBox txtUnitCost 
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   585
      End
      Begin VB.HScrollBar hsbAmount 
         Height          =   255
         LargeChange     =   2
         Left            =   975
         Max             =   0
         TabIndex        =   8
         Top             =   960
         Width           =   2520
      End
      Begin VB.Label lblQuantity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3150
         TabIndex        =   14
         Top             =   1020
         Width           =   315
      End
      Begin VB.Label lblName 
         Height          =   195
         Left            =   660
         TabIndex        =   13
         Top             =   735
         Width           =   2430
      End
      Begin VB.Label lblTotal 
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1725
         Width           =   660
      End
      Begin VB.Label lblUnitCost 
         Caption         =   "Cost Per Team:"
         Height          =   240
         Left            =   750
         TabIndex        =   9
         Top             =   375
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab tabResearch 
      Height          =   2745
      Left            =   75
      TabIndex        =   22
      Top             =   120
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Planetary"
      TabPicture(0)   =   "FRMRESEA.frx":124CC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "optShield"
      Tab(0).Control(1)=   "optBase2"
      Tab(0).Control(2)=   "optBase1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Military"
      TabPicture(1)   =   "FRMRESEA.frx":124E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optLongBioRocket"
      Tab(1).Control(1)=   "optChemical"
      Tab(1).Control(2)=   "optMech"
      Tab(1).Control(3)=   "optPlasma"
      Tab(1).Control(4)=   "optLasers"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Warp Ships"
      TabPicture(2)   =   "FRMRESEA.frx":12504
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optCloak"
      Tab(2).Control(1)=   "optUltra"
      Tab(2).Control(2)=   "optBigger"
      Tab(2).Control(3)=   "optShields2"
      Tab(2).Control(4)=   "optShields1"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Resources"
      TabPicture(3)   =   "FRMRESEA.frx":12520
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "optRegenerate"
      Tab(3).Control(1)=   "optCleanup"
      Tab(3).Control(2)=   "optImprove"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Scanners"
      TabPicture(4)   =   "FRMRESEA.frx":1253C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "optShortScan"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "optLongScan"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "optJam"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "optWarpScan"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.OptionButton optLongBioRocket 
         Caption         =   "Long-Range BioHazard Rocket"
         Height          =   255
         Left            =   -74460
         TabIndex        =   42
         Top             =   2025
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optRegenerate 
         Caption         =   "Regenerate Barren Environments"
         Height          =   240
         Left            =   -74460
         TabIndex        =   41
         Top             =   1315
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.OptionButton optCleanup 
         Caption         =   "BioHazard Detoxification"
         Height          =   225
         Left            =   -74460
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.OptionButton optImprove 
         Caption         =   "Improved Resource Production"
         Height          =   225
         Left            =   -74460
         TabIndex        =   39
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton optCloak 
         Caption         =   "Cloaking Device"
         Height          =   210
         Left            =   -74460
         TabIndex        =   38
         Top             =   2025
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.OptionButton optUltra 
         Caption         =   "Ultra-Warp Ship Engines"
         Height          =   240
         Left            =   -74460
         TabIndex        =   37
         Top             =   1665
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.OptionButton optBigger 
         Caption         =   "Expand Ship Capacity"
         Height          =   210
         Left            =   -74460
         TabIndex        =   36
         Top             =   1315
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.OptionButton optShields2 
         Caption         =   "Shields - Level II"
         Height          =   195
         Left            =   -74460
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.OptionButton optShields1 
         Caption         =   "Shields - Level I"
         Height          =   270
         Left            =   -74460
         TabIndex        =   34
         Top             =   600
         Width           =   2685
      End
      Begin VB.OptionButton optChemical 
         Caption         =   "BioHazard Rocket"
         Height          =   255
         Left            =   -74460
         TabIndex        =   33
         Top             =   1665
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.OptionButton optMech 
         Caption         =   "Mechanized Assault Units"
         Height          =   225
         Left            =   -74460
         TabIndex        =   32
         Top             =   1315
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.OptionButton optPlasma 
         Caption         =   "Plasma Rifle"
         Height          =   240
         Left            =   -74460
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton optLasers 
         Caption         =   "Laser Rifle"
         Height          =   240
         Left            =   -74460
         TabIndex        =   30
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton optWarpScan 
         Caption         =   "Warp Scanner"
         Height          =   225
         Left            =   540
         TabIndex        =   29
         Top             =   1665
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.OptionButton optJam 
         Caption         =   "Scanner Jamming Device"
         Height          =   240
         Left            =   540
         TabIndex        =   28
         Top             =   1315
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.OptionButton optLongScan 
         Caption         =   "Long-Range Scanner"
         Height          =   240
         Left            =   540
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.OptionButton optShortScan 
         Caption         =   "Short-Range Scanner"
         Height          =   195
         Left            =   540
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton optShield 
         Caption         =   "Planetary Shield"
         Height          =   225
         Left            =   -74460
         TabIndex        =   25
         Top             =   1315
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.OptionButton optBase2 
         Caption         =   "Missile Base - Level II"
         Height          =   240
         Left            =   -74460
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.OptionButton optBase1 
         Caption         =   "Missile Base - Level I"
         Height          =   210
         Left            =   -74460
         TabIndex        =   23
         Top             =   600
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'unload the form without doing anything
Unload frmResearch

End Sub




Private Sub cmdOK_Click()
If SoundOn Then
    PlaySoundEffect "Button4"
End If

'calculate the time to research
'these formulas have been changed a lot for play balance, and may not be the same as
'noted in each case statement

PurchasePrice = Val(txtTotal.Text)

Select Case lblName.Caption
Case "Missile Base - Level I"
    '*unit cost of 10--- finish in 2-5 turns
    'set time to finish research
    Player(Current).Missile1ResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats
    
Case "Missile Base - Level II"
    '**unit cost of 15--- finish in 2-5 turns
    'set time to finish research
    Player(Current).Missile2ResearchDone = TurnNumber + Int(Rnd * Int(45 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats
    
Case "Planetary Shield"
    '**unit cost of 15--- finish in 3-5 turns
    'set time to finish research
    Player(Current).ShieldResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 3
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Laser Rifle"
    '**unit cost of 10--- finish in 2-4 turns
    'set time to finish research
    Player(Current).LaserResearchDone = TurnNumber + Int(Rnd * Int(20 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Plasma Rifle"
    '**unit cost of 15--- finish in 2-4 turns
    'set time to finish research
    Player(Current).PlasmaResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Mechanized Assault Troops"
    '**unit cost=15 --- finish in 2-5 turns
    'set time to finish research
    Player(Current).MechResearchDone = TurnNumber + Int(Rnd * Int(45 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "BioHazard Rocket"
    '**unit cost of 25--- finish in 4-6 turns
    'set time to finish research
    Player(Current).BioRocketResearchDone = TurnNumber + Int(Rnd * Int(75 / PurchasePrice)) + 3
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Long-Range BioHazard Rocket"
    '**unit cost of 25--- finish in 2-4 turns
    'set time to finish research
    Player(Current).LongBioResearchDone = TurnNumber + Int(Rnd * Int(50 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Shields - Level I"
    '**unit cost of 10--- finish in 2-4 turns
    'set time to finish research
    Player(Current).ShipShield1ResearchDone = TurnNumber + Int(Rnd * Int(20 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Shields - Level II"
    '**unit cost of 15--- finish in 2-4 turns
    'set time to finish research
    Player(Current).ShipShield2ResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats
    

Case "Expand Ship Capacity"
    '*unit cost 15 - 2-5 turns
    Player(Current).BigShipResearchDone = TurnNumber + Int(Rnd * Int(60 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "UltraWarp Ship Engines"
    '*unit cost 20 - 2-5 turns
    Player(Current).UltraWarpResearchDone = TurnNumber + Int(Rnd * Int(60 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Cloaking Device"
    '*unit cost 20, 3-5 turns
    Player(Current).CloakingResearchDone = TurnNumber + Int(Rnd * Int(60 / PurchasePrice)) + 3
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Improved Resource Production"
    '**unit cost=20---finish in 3-5 turns
    'set time to finish research
    Player(Current).ResourceResearchDone = TurnNumber + Int(Rnd * Int(60 / PurchasePrice)) + 3
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "BioHazard Detoxification"
    '**unit cost=15---finish in 2-4 turns
    'set time to finish research
    Player(Current).BioCleanupResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Regenerate Barren Environments"
    '**unit cost=15---finish in 2-4 turns
    'set time to finish research
    Player(Current).RegenerationResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Short-Range Scanner"
    '**unit cost 10  - finish in 2-4 turns
    'set time to finish research
    Player(Current).ScannerResearchDone = TurnNumber + Int(Rnd * Int(20 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats
    
Case "Long-Range Scanner"
    '**unit cost 15  - finish in 2-5 turns
    'set time to finish research
    Player(Current).DeepScannerResearchDone = TurnNumber + Int(Rnd * Int(45 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Scanner Jamming Device"
    '**unit cost 15  - finish in 2-5 turns
    'set time to finish research
    Player(Current).JammerResearchDone = TurnNumber + Int(Rnd * Int(45 / PurchasePrice)) + 2
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

Case "Warp Scanner"
    '**unit cost 15  - finish in 3-5 turns
    'set time to finish research
    Player(Current).WarpScannerResearchDone = TurnNumber + Int(Rnd * Int(30 / PurchasePrice)) + 3
    Player(Current).NumResources = Player(Current).NumResources - PurchasePrice
    'call update procedure
    frmGameScreen.UpdatePlayerStats

End Select

frmGameScreen.ClearFrame

'unload the form
Unload frmResearch

End Sub



Private Sub Form_Load()

'See which tab items should be visible and/or disabled
If Player(Current).Missile1Researched Then
    optBase2.Visible = True
    optBase1.Enabled = False
End If

'see if Planetary Shield
If Player(Current).Missile2Researched Then
    optShield.Visible = True
    optBase2.Enabled = False
End If

'disable planetary shield if done
If Player(Current).ShieldResearched Then
    optShield.Enabled = False
End If

'see if plasma rifle
If Player(Current).LaserResearched Then
    optPlasma.Visible = True
    optLasers.Enabled = False
End If

'see if mech
If Player(Current).PlasmaResearched Then
    optMech.Visible = True
    optPlasma.Enabled = False
End If

'see if biohazard
If Player(Current).MechResearched Then
    optChemical.Visible = True
    optMech.Enabled = False
End If

'see if long-range biohazard
If Player(Current).BioRocketResearched Then
    optLongBioRocket.Visible = True
    optChemical.Enabled = False
End If

If Player(Current).LongBioResearched Then
    optLongBioRocket.Enabled = False
End If

'see if warp shields 2 and bigger ship
If Player(Current).ShipShield1Researched Then
    optShields2.Visible = True
    optBigger.Visible = True
    optShields1.Enabled = False
End If

'see if ultrawarp choice should be enabled
If Player(Current).BigShipResearched Then
    optUltra.Visible = True
    optShields2.Enabled = False
    optBigger.Enabled = False
End If

'turn off shield2 if done
If Player(Current).ShipShield2Researched Then
    optShields2.Enabled = False
End If

'see if cloaking
If Player(Current).UltraWarpResearched Then
    optCloak.Visible = True
    optUltra.Enabled = False
End If

If Player(Current).CloakingResearched Then
    optCloak.Enabled = False
End If

'see if detox
If Player(Current).BioRocketResearched And Player(Current).ResourcesResearched Then
    optCleanup.Visible = True
    optImprove.Enabled = False
End If

'see if regeneration
If Player(Current).BioCleanupResearched Then
    optRegenerate.Visible = True
    optCleanup.Enabled = False
End If

If Player(Current).RegenerationResearched Then
    optRegenerate.Enabled = False
End If

'see if long-range scanner should be on
If Player(Current).ScannerResearched Then
    optLongScan.Visible = True
    optShortScan.Enabled = False
End If

'see if jammer
If Player(Current).DeepScannerResearched Then
    optJam.Visible = True
    optWarpScan.Visible = True
    optLongScan.Enabled = False
End If

'disable jammer, warpscan if necessary
If Player(Current).JammerResearched Then
    optJam.Enabled = False
End If

If Player(Current).WarpScannerResearched Then
    optWarpScan.Enabled = False
End If

If Player(Current).RegenerationResearched Then
    optRegenerate.Enabled = False
End If

If Player(Current).ResourcesResearched Then
    optImprove.Enabled = False
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
cmdCancel_Click

End Sub


Private Sub hsbAmount_Change()
'use horizontal scroll bar to set amount of resources being allocated
'also done in Scroll method
lblQuantity.Caption = Str(hsbAmount.Value)

'calculate total cost:
txtTotal.Text = Str(hsbAmount.Value * UnitCost)

If hsbAmount.Value = 0 Then
    cmdOK.Enabled = False
End If

If hsbAmount.Value > 0 Then
    cmdOK.Enabled = True
End If

End Sub

Private Sub hsbAmount_Scroll()
lblQuantity.Caption = Str(hsbAmount.Value)
txtTotal.Text = Str(hsbAmount.Value * UnitCost)

End Sub


Private Sub optAssault_Click()
'see if mech research already done
If Player(Current).MechResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).MechResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Assault Troops"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        lblName.Caption = "Assault Troops"
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Funding Shortfall"
        ClearFrame
    End If

End If


End Sub

Private Sub optBase1_Click()
'see if missile base 1 research already done
If Player(Current).Missile1Researched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).Missile1ResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Missile Base - Level I"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        'lblName.Caption = "Planetary Defenses"
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optBase2_Click()
'see if missile base 2 research already done
If Player(Current).Missile2Researched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).Missile2ResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Missile Base - Level II"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        'lblName.Caption = "Planetary Defenses"
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub


Private Sub optBigger_Click()
'see if bigger ship research already done
If Player(Current).BigShipResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).BigShipResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Expand Ship Capacity"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        lblName.Caption = "Expand Ship Capacity"
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub









Private Sub optChemical_Click()
'see if biohazard rocket research already done
If Player(Current).BioRocketResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).BioRocketResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 25
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "BioHazard Rocket"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optCleanup_Click()
'see if planet detoxification research already done
If Player(Current).BioCleanupResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).BioCleanupResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "BioHazard Detoxification"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If


End Sub

Private Sub optCloak_Click()
'see if cloaking device research already done
If Player(Current).CloakingResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).CloakingResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 20
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Cloaking Device"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub





Private Sub optDeepScanner_Click()
'see if long-range scanner research already done
If Player(Current).DeepScannerResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).DeepScannerResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Deep-Space Scanner"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        lblName.Caption = "Deep-Space Scanner"
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Funding Shortfall"
        ClearFrame
    End If
End If

End Sub

Private Sub optImprove_Click()
'see if improved resource production research already done
If Player(Current).ResourcesResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).ResourceResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 20
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Improved Resource Production"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If


End Sub

Private Sub optJam_Click()
'see if scanner jamming research already done
If Player(Current).JammerResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).JammerResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Scanner Jamming Device"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub


Private Sub optLasers_Click()
'see if laser rifle research already done
If Player(Current).LaserResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).LaserResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Laser Rifle"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optLongBioRocket_Click()
'see if long-range biorocket research already done
If Player(Current).LongBioResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).LongBioResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 25
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Long-Range BioHazard Rocket"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optLongScan_Click()
'see if long-range scanner research already done
If Player(Current).DeepScannerResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).DeepScannerResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Long-Range Scanner"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optMech_Click()
'see if mech research already done
If Player(Current).MechResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).MechResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Mechanized Assault Troops"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If

End If

End Sub



Private Sub optPlasma_Click()
'see if plasma rifle research already done
If Player(Current).PlasmaResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).PlasmaResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Plasma Rifle"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optRegenerate_Click()
'see if planet regeneration research already done
If Player(Current).RegenerationResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).RegenerationResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Regenerate Barren Environments"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If


End Sub




Private Sub optShield_Click()
'see if planetary shield research already done
If Player(Current).ShieldResearched Then
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).ShieldResearchDone > 0 Then
    'research already underway
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Planetary Shield"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If


End Sub

Private Sub optShields1_Click()
'see if ships shield 1 research already done
If Player(Current).ShipShield1Researched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).ShipShield1ResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Shields - Level I"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optShields2_Click()
'see if ships shield 2 research already done
If Player(Current).ShipShield2Researched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).ShipShield2ResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Shields - Level II"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub


Public Sub ClearFrame()
'clear out the text, value caption etc from research frame
'whenever user selects another choice, or has insufficient
'resources, or project underway
hsbAmount.Value = 0
txtUnitCost.Text = ""
txtTotal.Text = ""
lblName.Caption = ""
fraAmount.Enabled = False

'reset values of option buttons
optBase1.Value = False
optBase2.Value = False
optShield.Value = False
'***
optLasers.Value = False
optPlasma.Value = False
optMech.Value = False
optChemical.Value = False
optLongBioRocket.Value = False
'***
optShields1.Value = False
optShields2.Value = False
optBigger.Value = False
optUltra.Value = False
optCloak.Value = False
'**
optImprove.Value = False
optCleanup.Value = False
optRegenerate.Value = False
'**
optShortScan.Value = False
optLongScan.Value = False
optJam.Value = False
optWarpScan.Value = False

End Sub

Private Sub optShortScan_Click()
'see if scanner research already done
If Player(Current).ScannerResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).ScannerResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 10
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Short-Range Scanner"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If


End Sub

Private Sub optUltra_Click()
'see if ultrawarp engine research already done
If Player(Current).UltraWarpResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).UltraWarpResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub

Else
    'check if enough money
    UnitCost = 20
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "UltraWarp Ship Engines"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If

End Sub

Private Sub optWarpScan_Click()
'see if warp path scanner research already done
If Player(Current).WarpScannerResearched Then
    PlaySoundEffect "Quiet"
    MsgBox "You already have this technology"
    ClearFrame
    Exit Sub
ElseIf Player(Current).WarpScannerResearchDone > 0 Then
    'research already underway
    PlaySoundEffect "Quiet"
    MsgBox "This project is already underway"
    ClearFrame
    Exit Sub
Else
    'check if enough money
    UnitCost = 15
    txtUnitCost.Text = Str(UnitCost)
    lblName.Caption = "Warp Scanner"

    If Player(Current).NumResources >= UnitCost Then
        'enable the amount frame
        fraAmount.Enabled = True
        hsbAmount.Value = 0
        'set label,text, and unit cost
        txtUnitCost.Text = Str(UnitCost)

        'set max value for scrollbar
        hsbAmount.Max = Int(Player(Current).NumResources / UnitCost)
    Else
        'not enough money
        PlaySoundEffect "Quiet"
        MsgBox "Insufficient resources", vbExclamation, "Project Cancelled"
        ClearFrame
    End If
End If
End Sub




Private Sub tabResearch_Click(PreviousTab As Integer)
'reset research variables
ClearFrame

End Sub

