VERSION 5.00
Begin VB.Form frmMessageConsole 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MessageConsole"
   ClientHeight    =   3030
   ClientLeft      =   1440
   ClientTop       =   1260
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraExitMessage 
      BackColor       =   &H00404040&
      Height          =   750
      Left            =   4350
      TabIndex        =   4
      Top             =   1875
      Width           =   1710
      Begin VB.CommandButton cmdExitMessage 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   285
         TabIndex        =   5
         Top             =   255
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   750
      Left            =   750
      TabIndex        =   1
      Top             =   1875
      Width           =   3240
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1695
         TabIndex        =   3
         Top             =   255
         Width           =   1200
      End
      Begin VB.CommandButton cmdEnterMessage 
         Caption         =   "&Enter Message"
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   255
         Width           =   1335
      End
   End
   Begin VB.TextBox txtMessageBox 
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
      Height          =   1380
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   405
      Width           =   5655
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   255
      Picture         =   "FRMMESSA.frx":0000
      Stretch         =   -1  'True
      Top             =   15
      Width           =   6375
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   255
      Picture         =   "FRMMESSA.frx":42D2
      Stretch         =   -1  'True
      Top             =   2805
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   3030
      Left            =   6645
      Picture         =   "FRMMESSA.frx":85A4
      Stretch         =   -1  'True
      Top             =   15
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   3030
      Left            =   0
      Picture         =   "FRMMESSA.frx":BEC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmMessageConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdEnterMessage_Click()
'unlock the text box to allow typing
txtMessageBox.Locked = False

'clear the text box for typing
txtMessageBox.Text = ""

'enable the "send" button
cmdSend.Enabled = True
cmdSend.Default = True

'move the focus to the textbox
txtMessageBox.SetFocus
End Sub


Private Sub cmdExitMessage_Click()

PlaySoundEffect "Button3"

'get rid of 'incoming message'
frmGameScreen.txtMessages.FontBold = False
frmGameScreen.txtMessages.ForeColor = vbGreen
frmGameScreen.txtMessages.Text = "No New Messages..."

Unload frmMessageConsole


End Sub




Private Sub cmdSend_Click()

PlaySoundEffect "Button4"

'Finished typing message - save as OutgoingMessage
OutgoingMessage = txtMessageBox.Text

'get rid of 'incoming message'
frmGameScreen.txtMessages.FontBold = False
frmGameScreen.txtMessages.ForeColor = vbGreen
frmGameScreen.txtMessages.Text = "*Message Sent*"
frmGameScreen.tmrUpdateMessageBox.Enabled = True

'unload the form
Unload frmMessageConsole


End Sub






Private Sub Form_Load()
txtMessageBox.Text = IncomingMessage

'lock the text box to prevent alteration of text
txtMessageBox.Locked = True


End Sub


