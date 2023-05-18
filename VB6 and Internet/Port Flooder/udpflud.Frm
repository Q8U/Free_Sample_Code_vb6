VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDP Flooder"
   ClientHeight    =   2895
   ClientLeft      =   2100
   ClientTop       =   2160
   ClientWidth     =   5175
   Icon            =   "udpflud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Halt"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtLeft 
      BackColor       =   &H00000000&
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox status 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   0
   End
   Begin VB.TextBox txtDelay 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Text            =   "20"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtSends 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   $"udpflud.frx":1782
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "139"
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtHost 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "206.432.2.4"
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox UDP 
      Height          =   480
      Left            =   5400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Times Left To Send"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Delay In Milliseconds"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Number Of times To Send"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Text To Send"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Port To Use"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Host To Send To"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu2 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu3 
         Caption         =   "S&top"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu4 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdClear_Click()
  status.Text = "       ¤¤¤¤¤ UDP Flooder ¤¤¤¤¤        "
  status.Text = status.Text & "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤"
End Sub


Private Sub cmdSend_Click()
  txtLeft.Text = txtSends.Text
  If txtLeft.Text = 0 Then
    status.Text = status.Text & " ¤¤¤¤¤ UDP Flood Complete ¤¤¤¤¤"
    Timer1.Enabled = False
    txtLeft.Text = txtSends.Text
  End If
End Sub


Private Sub cmdStop_Click()
  Timer1.Enabled = False
  status.Text = status.Text & "Flooding has Been Stopped"
End Sub

Private Sub Form_Load()
  status.Text = "       ¤¤¤¤¤ UDP Flooder ¤¤¤¤¤        "
  status.Text = status.Text & "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤"
  txtLeft.Text = txtSends.Text
End Sub

Private Sub Text1_Change()

End Sub

Private Sub mnu2_Click()
  txtLeft.Text = txtSends.Text
  If txtLeft.Text = 0 Then
    status.Text = status.Text & " ¤¤¤¤¤ UDP Flood Complete ¤¤¤¤¤"
    Timer1.Enabled = False
    txtLeft.Text = txtSends.Text
  End If
End Sub

Private Sub mnu3_Click()
  Timer1.Enabled = False
  status.Text = status.Text & "Flooding Has Been Stopped"
End Sub

Private Sub mnu4_Click()
  Unload Me
End Sub

Private Sub mnu6_Click()
  frmAbout.Show
  Unload Me
End Sub

Private Sub Timer1_Timer()
  txtDelay.Text = Timer1.Interval
  txtLeft.Text = txtLeft.Text - 1
  If txtSends = 0 Then GoTo udp_end
udp_end:
  status.Text = status.Text & " ¤¤¤¤¤ UDP Flood Complete ¤¤¤¤¤"
  udpTimer.Enabled = False
  cmdUDPSendData.Enabled = True
  txtLeft.Text = txtSends.Text
End Sub


