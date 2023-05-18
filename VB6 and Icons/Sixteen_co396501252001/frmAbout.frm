VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Icon Editor..."
   ClientHeight    =   4305
   ClientLeft      =   675
   ClientTop       =   1725
   ClientWidth     =   7215
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Height          =   372
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   972
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   335.805
      ScaleMode       =   0  'User
      ScaleWidth      =   335.805
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   120
      Picture         =   "frmAbout.frx":0614
      Top             =   1080
      Width           =   8775
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Icon Editor by Gregg Cleveland"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1056
      TabIndex        =   1
      Top             =   240
      Width           =   3888
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version 1.1"
      Height          =   228
      Left            =   1056
      TabIndex        =   2
      Top             =   600
      Width           =   3888
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Me.Caption = "About " & App.Title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'lblTitle.Caption = App.Title

'Adjust size to original
Width = Int(7254 * Screen.TwipsPerPixelX / 12 + 0.5)
Height = Int(4092 * Screen.TwipsPerPixelY / 12 + 0.5)

'Center form
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

cmdOK.Left = 250
cmdOK.Top = 270
picIcon.Left = 20
picIcon.Top = 20
Visible = True
End Sub
