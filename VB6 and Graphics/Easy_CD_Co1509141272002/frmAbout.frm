VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About CD Cover (1.0.0)"
   ClientHeight    =   1845
   ClientLeft      =   3735
   ClientTop       =   3000
   ClientWidth     =   5565
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1273.452
   ScaleMode       =   0  'User
   ScaleWidth      =   5225.823
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Label lblinfo2 
      AutoSize        =   -1  'True
      Caption         =   "Email to:"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblemail 
      AutoSize        =   -1  'True
      Caption         =   "johnge2@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Designed by GaTe^KeEpEr / Drag 'n' Drop Added By MaD^DoGg"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -140.858
      X2              =   5084.026
      Y1              =   942.147
      Y2              =   942.147
   End
   Begin VB.Label lblTitle 
      Caption         =   "Easy CD Cover"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -126.772
      X2              =   5084.026
      Y1              =   952.501
      Y2              =   952.501
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Easy CD Cover v1.0 Source Release
'
' by: G@Te^k3eP3R (johnge2@yahoo.com)
'

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
lblemail = email


End Sub

Private Sub lblemail_Click()
sendemail
End Sub

