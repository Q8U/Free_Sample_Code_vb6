VERSION 5.00
Object = "{C61830C1-8B47-11D4-9F3F-0000B45C4CF6}#1.0#0"; "EASYSOUND.OCX"
Begin VB.Form frmEasySound 
   Caption         =   "Using Sounds"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "frmEasySound.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetGunFreq 
      Caption         =   "Set Frequency"
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtGunFreq 
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtBoinkFreq 
      Height          =   285
      Left            =   120
      MaxLength       =   6
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetBoinkFreq 
      Caption         =   "Set Frequency"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetBoinkFreq 
      Caption         =   "Get Frequency"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetCountFreq 
      Caption         =   "Get Frequency"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetGunFreq 
      Caption         =   "Get Frequency"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayDupGun 
      Caption         =   "Play duplicate gun"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayDupBoink 
      Caption         =   "Play Duplicate boink"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Count"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDupBoink 
      Caption         =   "Duplicate boink"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDupGun 
      Caption         =   "Duplicate Gun"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdStream 
      Caption         =   "&Count down"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdMachine 
      Caption         =   "&Machine Gun"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBoink 
      Caption         =   "&Boink"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin EASYSOUNDLibCtl.ESound ESound 
      Left            =   3240
      OleObjectBlob   =   "frmEasySound.frx":014A
      Top             =   3480
   End
   Begin VB.Label lblGunLenght 
      AutoSize        =   -1  'True
      Caption         =   "Length: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   19
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lblBoinkLenght 
      AutoSize        =   -1  'True
      Caption         =   "Length: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lblBoinkFreq 
      AutoSize        =   -1  'True
      Caption         =   "Frequency: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label lblFreqCount 
      AutoSize        =   -1  'True
      Caption         =   "Frequency: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3240
      TabIndex        =   11
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label lblGunFreq 
      AutoSize        =   -1  'True
      Caption         =   "Frequency: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   3840
      Width           =   840
   End
End
Attribute VB_Name = "frmEasySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StaticMachineGun As Long
Dim StaticBoink As Long
Dim StreamCount As Long

Dim StaticDupGun As Long
Dim StaticDupBoink As Long


Private Sub cmdBoink_Click()

lblBoinkLenght.Caption = "Length: " & ESound.GetStaticByteLength(StaticBoink)

ESound.PlayStaticSound StaticBoink, 0

End Sub

Private Sub cmdDupBoink_Click()

StaticDupBoink = ESound.DuplicateStaticSound(StaticBoink)
cmdPlayDupBoink.Enabled = True


End Sub

Private Sub cmdDupGun_Click()

    StaticDupGun = ESound.DuplicateStaticSound(StaticMachineGun)
    cmdPlayDupGun.Enabled = True
    
End Sub


Private Sub cmdGetBoinkFreq_Click()

lblBoinkFreq.Caption = "Frequency: " & ESound.GetStaticFrequency(StaticBoink)

End Sub

Private Sub cmdGetCountFreq_Click()

lblFreqCount.Caption = "Frequency: " & ESound.GetStreamingFrequency(StreamCount)

End Sub

Private Sub cmdGetGunFreq_Click()

lblGunFreq.Caption = "Frequency: " & ESound.GetStaticFrequency(StaticMachineGun)

End Sub

Private Sub cmdMachine_Click()

lblGunLenght.Caption = "Length: " & ESound.GetStaticByteLength(StaticMachineGun)
ESound.PlayStaticSound StaticMachineGun, 0


End Sub

Private Sub cmdPlayDupBoink_Click()

    ESound.PlayStaticSound StaticDupBoink, 0
    

End Sub

Private Sub cmdPlayDupGun_Click()

ESound.PlayStaticSound StaticDupGun, 0

End Sub

Private Sub cmdSetBoinkFreq_Click()

If CLng(Val(txtBoinkFreq.Text)) > 100 Or CLng(Val(txtBoinkFreq.Text)) < 100000 Then
    
    ESound.SetStaticFrequency StaticBoink, CLng(Val(txtBoinkFreq.Text))

End If


End Sub


Private Sub cmdSetGunFreq_Click()

If CLng(Val(txtGunFreq.Text)) > 100 Or CLng(Val(txtGunFreq.Text)) < 100000 Then
    
   ESound.SetStaticFrequency StaticMachineGun, CLng(Val(txtGunFreq.Text))

End If


End Sub

Private Sub cmdStop_Click()

    ESound.StopStreamingSound StreamCount

End Sub

Private Sub cmdStream_Click()

ESound.PlayStreamingSound StreamCount, 0


End Sub


Private Sub Command1_Click()

Debug.Print ESound.GetCurrentStaticPosition(StaticMachineGun)

End Sub

Private Sub Command2_Click()
 
 ESound.SetCurrentStaticPosition StaticBoink, 5000
 
 Debug.Print ESound.GetCurrentStaticPosition(StaticBoink)
 
End Sub

Private Sub Form_Load()
Dim rt As Long
Dim AppPath As String


'don´t forget this
ESound.Window = Me.hWnd

'Initialize the sound machine
rt = ESound.InitializeSound()

If rt <> EX_OK Then
    
    MsgBox "Could Not initialize sound", vbOKOnly, "Failure"
    Exit Sub
End If

AppPath = App.Path & "\"


StaticMachineGun = ESound.CreateStaticSound(AppPath & "machine.wav")


If StaticMachineGun < 0 Then
    MsgBox "Could not load machine gun sound", vbOKOnly, "Failure"
    cmdMachine.Enabled = False
End If


StaticBoink = ESound.CreateStaticSound(AppPath & "boink.wav")

If StaticBoink < 0 Then
    MsgBox "Could not load boink sound", vbOKOnly, "Failure"
    cmdBoink.Enabled = False
End If

StreamCount = ESound.CreateStreamingSound(AppPath & "countdown.wav")

If StreamCount < 0 Then
    MsgBox "Could not load count sound", vbOKOnly, "Failure"
    cmdStream.Enabled = False
    cmdStop.Enabled = False
End If


End Sub
