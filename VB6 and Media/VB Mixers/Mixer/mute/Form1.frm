VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   1485
      ScaleHeight     =   810
      ScaleWidth      =   5085
      TabIndex        =   0
      ToolTipText     =   "Mute"
      Top             =   1935
      Width           =   5145
      Begin VB.OptionButton Option2 
         BackColor       =   &H000000FF&
         Caption         =   "Silence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Mute"
         Top             =   405
         Width           =   5055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "listen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Unmute"
         Top             =   45
         Value           =   -1  'True
         Width           =   5055
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   135
      TabIndex        =   3
      Top             =   45
      Width           =   7710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unmute As MIXERCONTROL
Dim mute As MIXERCONTROL


Private Sub Form_Load()
Label1 = "Sorry the code is so big but the problem with the mute is that you have to get the line control of the device you wish to mute ie mic/master/TAD. The control then needs a buffer to store the retrieved info then the program reads the buffer and manipulate the contents ie you can mute the buffered control. So I have muted the master control hear for you."
End Sub

Private Sub Option1_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option2_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub
