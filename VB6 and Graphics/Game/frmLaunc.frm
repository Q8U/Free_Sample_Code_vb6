VERSION 5.00
Begin VB.Form frmLaunch 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Launching..."
   ClientHeight    =   3435
   ClientLeft      =   1500
   ClientTop       =   1890
   ClientWidth     =   6315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCloak 
      BackColor       =   &H00400000&
      Caption         =   " Cloaking Device"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   705
      TabIndex        =   13
      Top             =   2115
      Width           =   1830
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3900
      TabIndex        =   12
      Top             =   2475
      Width           =   1230
   End
   Begin VB.CommandButton cmdLaunchShip 
      Caption         =   "&Launch"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3900
      TabIndex        =   11
      Top             =   1980
      Width           =   1230
   End
   Begin VB.Frame fraAdvancedShip 
      BackColor       =   &H00400000&
      Caption         =   "Advanced Tactics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1305
      Left            =   435
      TabIndex        =   10
      Top             =   1725
      Width           =   2460
      Begin VB.CheckBox chkSabotage 
         BackColor       =   &H00400000&
         Caption         =   " Sabotage Mission"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   270
         TabIndex        =   14
         Top             =   810
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Select Ship"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1245
      Left            =   450
      TabIndex        =   7
      Top             =   360
      Width           =   1485
      Begin VB.CommandButton cmdShip2 
         Caption         =   "Ship 2"
         Height          =   285
         Left            =   195
         TabIndex        =   9
         Top             =   780
         Width           =   1000
      End
      Begin VB.CommandButton cmdShip1 
         Caption         =   "Ship 1"
         Height          =   285
         Left            =   195
         TabIndex        =   8
         Top             =   255
         Width           =   1000
      End
   End
   Begin VB.Frame fraTactical 
      BackColor       =   &H00400000&
      Caption         =   "Personnel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1245
      Left            =   2235
      TabIndex        =   0
      Top             =   360
      Width           =   3600
      Begin VB.HScrollBar hsbTroopsOnShip 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   2
         Left            =   960
         Max             =   1000
         TabIndex        =   2
         Top             =   270
         Width           =   2040
      End
      Begin VB.HScrollBar hsbAssaultTroopsOnShip 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   2
         Left            =   930
         Max             =   1000
         TabIndex        =   1
         Top             =   795
         Width           =   2070
      End
      Begin VB.Label lblRegularTroops 
         BackColor       =   &H00400000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblNumRegularTroops 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   3135
         TabIndex        =   5
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblAssaultTroops 
         BackStyle       =   0  'Transparent
         Caption         =   "Mechs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   165
         TabIndex        =   4
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblNumAssaultTroops 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   3135
         TabIndex        =   3
         Top             =   795
         Width           =   360
      End
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   195
      Picture         =   "FRMLAUNC.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   5850
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   195
      Picture         =   "FRMLAUNC.frx":42D2
      Stretch         =   -1  'True
      Top             =   15
      Width           =   5880
   End
   Begin VB.Image Image2 
      Height          =   3480
      Left            =   6075
      Picture         =   "FRMLAUNC.frx":85A4
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   15
      Picture         =   "FRMLAUNC.frx":BEC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "frmLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCloak_Click()
Dim CloakingCost As Integer
CloakingCost = 3

'see if player can afford to use the cloaking device
If Player(Current).NumResources >= CloakingCost Then
    'do nothing
Else
    PlaySoundEffect "Quiet"
    MsgBox "The Cloaking Device costs 3 resources to power up", , "Device Not Installed"
    chkCloak.Value = 0
End If

End Sub




Private Sub chkSabotage_Click()

If chkSabotage.Value = 0 Then
    'clean up
    Frame1.Enabled = True
    If Planet(ActivePlanet).Troops = 0 Then
        hsbTroopsOnShip.Enabled = False
    Else
        hsbTroopsOnShip.Value = 0
        lblNumRegularTroops.Caption = 0
    End If
    
    If Planet(ActivePlanet).AssaultTroops = 0 Then
        hsbAssaultTroopsOnShip.Enabled = False
    Else
        lblNumAssaultTroops.Caption = 0
        hsbAssaultTroopsOnShip.Value = 0
    End If

    chkSabotage.Value = 0
    chkCloak.Value = 0
    
End If

'make sure player has at least 1 mech to go on sabotage mission
If chkSabotage.Value = 1 Then
    If Planet(ActivePlanet).AssaultTroops < 1 Then
        PlaySoundEffect "Quiet"
        MsgBox "You must have at least one mech to" + Chr(13) + "carry out this mission.", vbOKOnly, "Mission Failure"
        chkSabotage.Value = 0
        Exit Sub
    Else
        'set up the mission
        Frame1.Enabled = False
        hsbTroopsOnShip.Value = 0
        hsbTroopsOnShip.Enabled = False
        hsbAssaultTroopsOnShip.Value = 1
        hsbAssaultTroopsOnShip.Enabled = False
        cmdLaunchShip.Enabled = True
    End If
End If

End Sub







Private Sub cmdCancel_Click()
'unload the form without doing anything
Unload frmLaunch

End Sub




Private Sub cmdLaunchShip_Click()

Dim CloakingCost As Integer
CloakingCost = 3

If hsbTroopsOnShip.Value = 0 And hsbAssaultTroopsOnShip = 0 Then
    PlaySoundEffect "Quiet"
    MsgBox "At least one troop or mech is required to operate the ship.", vbOKOnly + vbInformation, "Empty Ship"
    cmdLaunchShip.Enabled = False
    Exit Sub
End If

'update the ship with troops etc
With Player(Current).Ship(activeship)
     .Launched = True
     .Troops = hsbTroopsOnShip.Value
     .AssaultTroops = hsbAssaultTroopsOnShip.Value
     .Coordinate = Planet(ActivePlanet).Coordinate
     .WarpPosition = 1
     .CenterX = frmGameScreen.picPlanet(ActivePlanet).Left + (frmGameScreen.picPlanet(ActivePlanet).Width / 2)
     .CenterY = frmGameScreen.picPlanet(ActivePlanet).Top + (frmGameScreen.picPlanet(ActivePlanet).Height / 2)
     .ShipNumber = activeship    'set ship number to 0 or 1
End With

If chkSabotage.Value = 1 Then
    Player(Current).Ship(activeship).Sabotage = True
End If

'determine combatstrength
Player(Current).Ship(activeship).CombatStrength = 0
'laser and plasma for troops?
    
Dim X, Y
X = hsbTroopsOnShip.Value
Y = hsbAssaultTroopsOnShip.Value
    
If Player(Current).LaserResearched Then
    X = Int(X * 1.15)
End If

If Player(Current).PlasmaResearched Then
    X = Int(X * 1.3)
End If
    
'****Main Formula:
Player(Current).Ship(activeship).CombatStrength = X + (Y * 5)
'****

'shields
Dim Q As Integer
Q = 0
If Player(Current).ShipShield1Researched Then
    Q = 5
End If
    
If Player(Current).ShipShield2Researched Then
    Q = Q + 5
End If
    
Player(Current).Ship(activeship).CombatStrength = Player(Current).Ship(activeship).CombatStrength + Q
'end shields
      
       
'cloaking device
If chkCloak.Value = 1 Then
    Player(Current).Ship(activeship).HaveCloakingDevice = True
    Player(Current).NumResources = Player(Current).NumResources - CloakingCost
End If

'update player stats
frmGameScreen.UpdatePlayerStats

'update planet stats
With Planet(ActivePlanet)
    .Troops = Planet(ActivePlanet).Troops - hsbTroopsOnShip.Value
    .AssaultTroops = Planet(ActivePlanet).AssaultTroops - hsbAssaultTroopsOnShip.Value
    'adjust combat strength
     SetCombatStrength (ActivePlanet)
End With

'disable the planet management frame
frmGameScreen.ClearFrame
    
PlaySoundEffect "Launch"
    
'unload the launch form and return to game
Unload frmLaunch


End Sub

Private Sub cmdShip1_Click()
'set ship(0) as the one to be launched
activeship = 0

'***Scroll Bars*****
'see if regular troop scroll bar should be enabled
If Planet(ActivePlanet).Troops > 0 Then
    fraTactical.Enabled = True
    hsbTroopsOnShip.Enabled = True
    If Player(Current).BigShipResearched = False Then
        'set max values for regular troop scroll bar
        'limit max troops to 25
        hsbTroopsOnShip.Max = Planet(ActivePlanet).Troops
        If Planet(ActivePlanet).Troops > 25 Then
            hsbTroopsOnShip.Max = 25
        End If
    ElseIf Player(Current).BigShipResearched Then
        'no limit
        hsbTroopsOnShip.Max = Planet(ActivePlanet).Troops
    End If
End If

'see if assault troops scroll bar should be enabled
If Player(Current).MechResearched And Planet(ActivePlanet).AssaultTroops > 0 Then
    fraTactical.Enabled = True
    hsbAssaultTroopsOnShip.Enabled = True
    hsbAssaultTroopsOnShip.Max = Planet(ActivePlanet).AssaultTroops
    If Player(Current).BigShipResearched = False Then
       'set max value for scrollbar - limit of 5
       If Planet(ActivePlanet).AssaultTroops > 5 Then
          hsbAssaultTroopsOnShip.Max = 5
       End If

    ElseIf Player(Current).BigShipResearched = True Then
        'no limit on mechs on board
        hsbAssaultTroopsOnShip.Max = Planet(ActivePlanet).AssaultTroops

    End If
End If

'*****Advanced Tactics Frame*****
'Sabotage checkbox - only if mechs and shields1 researched already
If Player(Current).MechResearched And Player(Current).ShipShield1Researched = True Then
  chkSabotage.Enabled = True
End If

'Cloaking checkbox
If Player(Current).CloakingResearched Then
    chkCloak.Enabled = True
End If


End Sub

Private Sub cmdShip2_Click()
'set ship(1) as the one to be launched
activeship = 1

'***Scroll Bars*****
'see if regular troop scroll bar should be enabled
If Planet(ActivePlanet).Troops > 0 Then
    fraTactical.Enabled = True
    hsbTroopsOnShip.Enabled = True
    If Player(Current).BigShipResearched = False Then
        'set max values for regular troop scroll bar
        'limit max troops to 25
        hsbTroopsOnShip.Max = Planet(ActivePlanet).Troops
        If Planet(ActivePlanet).Troops > 25 Then
            hsbTroopsOnShip.Max = 25
        End If
    ElseIf Player(Current).BigShipResearched Then
        'no limit
        hsbTroopsOnShip.Max = Planet(ActivePlanet).Troops
    End If
End If

'see if assault troops scroll bar should be enabled
If Planet(ActivePlanet).AssaultTroops > 0 Then
    fraTactical.Enabled = True
    hsbAssaultTroopsOnShip.Enabled = True
    hsbAssaultTroopsOnShip.Max = Planet(ActivePlanet).AssaultTroops
    If Player(Current).BigShipResearched = False Then
       'set max value for scrollbar - limit of 5
       If Planet(ActivePlanet).AssaultTroops > 5 Then
          hsbAssaultTroopsOnShip.Max = 5
       End If

    ElseIf Player(Current).BigShipResearched = True Then
        'no limit on mechs on board
        hsbAssaultTroopsOnShip.Max = Planet(ActivePlanet).AssaultTroops

    End If
End If

'*****Advanced Tactics Frame*****
'Sabotage checkbox - only if mechs and ship shields 1 researched already
If Player(Current).MechResearched And Player(Current).ShipShield1Researched Then
   chkSabotage.Enabled = True
End If

'Cloaking checkbox
If Player(Current).CloakingResearched Then
    chkCloak.Enabled = True
End If

End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        'show quick help
        ShowQuickHelp
End Select
End Sub

Private Sub Form_Load()

'check which ships are available for launching, if any
If Player(Current).Ship(0).Launched Then
    cmdShip1.Enabled = False
End If

If Player(Current).Ship(1).Launched Then
    cmdShip2.Enabled = False
End If

'disable ship 2 if ship 1 is available
If Player(Current).Ship(0).Launched = False Then
    cmdShip2.Enabled = False
End If


'initial values for the labels
lblNumRegularTroops = 0
lblNumAssaultTroops = 0

End Sub




Private Sub hsbAssaultTroopsOnShip_Change()
'adjust label as scroll bar used
lblNumAssaultTroops.Caption = hsbAssaultTroopsOnShip.Value

'enable the launch button if more than 0 troops selected
If hsbAssaultTroopsOnShip.Value >= 1 Then
    cmdLaunchShip.Enabled = True
End If


End Sub

Private Sub hsbAssaultTroopsOnShip_Scroll()
lblNumAssaultTroops.Caption = hsbAssaultTroopsOnShip.Value

End Sub


Private Sub hsbTroopsOnShip_Change()
'update the label caption with the value
lblNumRegularTroops.Caption = hsbTroopsOnShip.Value

'enable the launch button if more than 0 troops selected
If hsbTroopsOnShip.Value >= 1 Then
    cmdLaunchShip.Enabled = True
End If

End Sub


Private Sub hsbTroopsOnShip_Scroll()
'update the label caption
lblNumRegularTroops.Caption = hsbTroopsOnShip.Value

'enable the launch button if more than 0 troops selected
If hsbTroopsOnShip.Value >= 1 Then
    cmdLaunchShip.Enabled = True
End If


End Sub


