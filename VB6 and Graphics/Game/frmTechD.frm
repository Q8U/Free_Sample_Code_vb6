VERSION 5.00
Begin VB.Form frmTechDone 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "TechDone form"
   ClientHeight    =   5025
   ClientLeft      =   615
   ClientTop       =   1305
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5025
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTech 
      BackColor       =   &H00000000&
      Height          =   2670
      Left            =   2025
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   3
      Top             =   645
      Width           =   4275
      Begin VB.Image imgTech 
         Height          =   2640
         Left            =   -30
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   4245
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3450
      TabIndex        =   0
      Top             =   4410
      Width           =   1455
   End
   Begin VB.Image imgShield 
      Height          =   585
      Left            =   1140
      Picture         =   "FRMTECHD.frx":0000
      Stretch         =   -1  'True
      Top             =   4020
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgBioRocket 
      Height          =   525
      Left            =   7635
      Picture         =   "FRMTECHD.frx":CD82
      Stretch         =   -1  'True
      Top             =   2730
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image imgLaser 
      Height          =   495
      Left            =   7635
      Picture         =   "FRMTECHD.frx":2BB90
      Stretch         =   -1  'True
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMissile1 
      Height          =   525
      Left            =   15
      Picture         =   "FRMTECHD.frx":4BD02
      Stretch         =   -1  'True
      Top             =   3930
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgWarpScan 
      Height          =   555
      Left            =   30
      Picture         =   "FRMTECHD.frx":59CEC
      Stretch         =   -1  'True
      Top             =   3270
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgScience2 
      Height          =   555
      Left            =   60
      Picture         =   "FRMTECHD.frx":681D6
      Stretch         =   -1  'True
      Top             =   2625
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgScience1 
      Height          =   585
      Left            =   30
      Picture         =   "FRMTECHD.frx":BC7E4
      Stretch         =   -1  'True
      Top             =   1980
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgCloaking 
      Height          =   555
      Left            =   45
      Picture         =   "FRMTECHD.frx":CBCE6
      Stretch         =   -1  'True
      Top             =   1395
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image imgJam 
      Height          =   510
      Left            =   30
      Picture         =   "FRMTECHD.frx":FD238
      Stretch         =   -1  'True
      Top             =   795
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image imgDeepScanner 
      Height          =   555
      Left            =   30
      Picture         =   "FRMTECHD.frx":10523A
      Stretch         =   -1  'True
      Top             =   180
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgUltraWarp 
      Height          =   570
      Left            =   7635
      Picture         =   "FRMTECHD.frx":114D5C
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Image4 
      Height          =   180
      Left            =   945
      Picture         =   "FRMTECHD.frx":1263EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6465
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   945
      Picture         =   "FRMTECHD.frx":12A6C0
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   6495
   End
   Begin VB.Image Image2 
      Height          =   4995
      Left            =   7380
      Picture         =   "FRMTECHD.frx":12E992
      Stretch         =   -1  'True
      Top             =   15
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   5025
      Left            =   720
      Picture         =   "FRMTECHD.frx":1322B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image imgResource 
      Height          =   555
      Left            =   7620
      Picture         =   "FRMTECHD.frx":135BD6
      Stretch         =   -1  'True
      Top             =   1575
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image imgPlanetShield 
      Height          =   495
      Left            =   7665
      Picture         =   "FRMTECHD.frx":148C18
      Stretch         =   -1  'True
      Top             =   1005
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image imgMech 
      Height          =   525
      Left            =   7605
      Picture         =   "FRMTECHD.frx":159FEA
      Stretch         =   -1  'True
      Top             =   420
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTitle 
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
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   330
      Width           =   4665
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   960
      Left            =   2055
      TabIndex        =   1
      Top             =   3330
      Width           =   4290
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTechDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
'continue with game
Unload Me

End Sub


Private Sub Form_Activate()
'draw lots of stars on the screen
Randomize
Dim a, X, Y, X2, Y2 As Integer

'draw white stars
For a = 1 To 40
   X = Int(Rnd * 900)
   Y = Int(Rnd * Me.ScaleHeight)
   X2 = Int(Rnd * 1000) + 7585
   Y2 = Int(Rnd * Me.ScaleHeight)
   Me.PSet (X, Y), vbWhite
   Me.PSet (X2, Y2), vbWhite
Next a
    
'draw dark grey stars
Dim grey
grey = &H808080
For a = 1 To 50
   X = Int(Rnd * 800)
   Y = Int(Rnd * Me.ScaleHeight)
   X2 = Int(Rnd * 1000) + 7585
   Y2 = Int(Rnd * Me.ScaleHeight)
   Me.PSet (X, Y), grey
   Me.PSet (X2, Y2), grey
Next a
    
'draw some blue stars
Dim blue
blue = &H800000
For a = 1 To 25
   X = Int(Rnd * 800)
   Y = Int(Rnd * Me.ScaleHeight)
   X2 = Int(Rnd * 1000) + 7585
   Y2 = Int(Rnd * Me.ScaleHeight)
   Me.PSet (X, Y), blue
   Me.PSet (X2, Y2), blue
Next a

PlaySoundEffect "Research"

End Sub




Private Sub Form_Load()
'set pictures and text based on what tech has been researched
'Techlevel variable set in form_load of frmGameScreen

Select Case TechLevel

Case 1
    'assault troops researched
    imgTech.Picture = imgMech.Picture
    lblTitle.Caption = "Assault Technology Research Completed!"
    lblDescription.Caption = "Mechanized Assault Troops deliver several times the offensive power of standard troops and can operate in toxic environments unsuited for humans. Mechs are required for Sabotage Missions, once Level I Ship Shields have been researched."

Case 2
    'planetary shield research done
    imgTech.Picture = imgPlanetShield.Picture
    lblTitle.Caption = "Planetary Shield Research Completed!"
    lblDescription.Caption = "An array of space-based laser defenses provides the highest level of planetary defense."

Case 3
    'resource research done
    imgTech.Picture = imgResource.Picture
    lblTitle.Caption = "Improved Resource Production Achieved!"
    lblDescription.Caption = "Planet resource production can be increased up to 80% with the construction of advanced production facilities."

Case 5
    'scanner research done
    imgTech.Picture = imgDeepScanner.Picture
    lblTitle.Caption = "Space Scanner Researched!"
    lblDescription.Caption = "Short-range scanners can be installed on any planet to allow detailed exploration of the galaxy. Long-range scanners and anti-scanner jamming devices can now be researched."
    
Case 6
    'bigger ships research done
    imgTech.Picture = imgScience1.Picture
    lblTitle.Caption = "Ship Capacity Expanded!"
    lblDescription.Caption = "Advances in ship design allow for unlimited expansion of the warp ship's transport capacity. The possibility of ultra-warp travel can now be researched."
    
Case 7
    'UltraWarp research done
    imgTech.Picture = imgUltraWarp.Picture
    lblTitle.Caption = "UltraWarp Engines Developed!"
    lblDescription.Caption = "UltraWarp engines greatly increase the range of warp ships, almost doubling the possible distance travelled."

Case 8
    'deep scanner research done
    imgTech.Picture = imgDeepScanner.Picture
    lblTitle.Caption = "Deep-Space Scanner Researched!"
    lblDescription.Caption = "Advanced optical technology expands the range of your new and existing scanners."
  
Case 9
    'jammer research done
    imgTech.Picture = imgJam.Picture
    lblTitle.Caption = "Anti-Scanner Jamming Researched!"
    lblDescription.Caption = "Planets can now be shielded from enemy scanners with new jamming technologies."
  
Case 10
    'cloaking device done
    imgTech.Picture = imgCloaking.Picture
    lblTitle.Caption = "Cloaking Device Research Complete!"
    lblDescription.Caption = "The Cloaking Device allows warp ships to travel almost completely undetected through the warp path. Ships must be specially fitted with the device for each launch."

Case 11
    'Missile1
    imgTech.Picture = imgMissile1.Picture
    lblTitle.Caption = "Missile Research (Level I) Complete!"
    lblDescription.Caption = "Missiles provide a basic level of defense against invading forces. More advanced missile technology can now be researched."
    
Case 12
    'Missile2
    imgTech.Picture = imgScience1.Picture
    lblTitle.Caption = "Missile Research (Level II) Complete!"
    lblDescription.Caption = "Advanced rocket technology provides greater protection against attempted enemy landings."
    
Case 13
    'Laser rifle
    imgTech.Picture = imgLaser.Picture
    lblTitle.Caption = "Laser Rifle Research Complete!"
    lblDescription.Caption = "Troops are now outfitted with laser rifles for increased combat performance. Even greater firepower is possibile with Plasma Rifles."
    
Case 14
    'plasma rifle
    imgTech.Picture = imgScience2.Picture
    lblTitle.Caption = "Plasma Rifle Research Complete!"
    lblDescription.Caption = "Troops are now outfitted with plasma rifles, and their combat strength is dramatically improved."
    
Case 15
    'Biorocket
    imgTech.Picture = imgScience2.Picture
    lblTitle.Caption = "BioHazard Rocket Research Complete!"
    lblDescription.Caption = "These rockets contaminate target planets with a deadly mixture of radiation and biochemical weapons, destroying resource capacity and rendering planets all but uninhabitable."
    
Case 16
    'long biorocket
    imgTech.Picture = imgScience1.Picture
    lblTitle.Caption = "Long-Range BioHazard Rocket Research Complete!"
    lblDescription.Caption = "The destructive power of the BioHazard Rockets can now be unleashed on a wider range of targets."
    
Case 17
    'shipshield1
    imgTech.Picture = imgShield.Picture
    lblTitle.Caption = "Ship Shields (Level I) Research Complete!"
    lblDescription.Caption = "These shields provide basic protection against a planet's missile or plantary shield defenses. More advanced shielding can now be researched. Sabotage missions can be launched once Mechanized Assault Troop technology has been researched."
    
Case 18
    'shipshield2
    imgTech.Picture = imgScience2.Picture
    lblTitle.Caption = "Ship Shields (Level II) Research Complete!"
    lblDescription.Caption = "Advanced Shields provide the maximum level of protection for ships attacking enemy planets. "
    Player(Current).ShipShield2Researched = True
    'this seems to solve the problem of this research item not being recognized...
    
Case 19
    'biocleanup
    imgTech.Picture = imgScience1.Picture
    lblTitle.Caption = "BioHazard Cleanup Research Complete!"
    lblDescription.Caption = "Contaminated planets can be detoxified and returned to a habitable state. With additional research, a planet's resource capacity can also be regenerated."
    
Case 20
    'regeneration
    imgTech.Picture = imgScience2.Picture
    lblTitle.Caption = "Regeneration Research Complete!"
    lblDescription.Caption = "Previously contaminated planets can have their resource protential rejuvenated, and possibly even increased, through advanced environmental technology."
    
Case 21
    'warp scanner
    imgTech.Picture = imgScience1.Picture
    lblTitle.Caption = "Warp Scanner Research Complete!"
    lblDescription.Caption = "Using the warp scanner, you can preview the possible landing sites of enemy ships in the warp path."
    
End Select


End Sub


















