VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   1515
   ClientLeft      =   5100
   ClientTop       =   4725
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdexit 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmHelp.frx":0000
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4245
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Easy CD Cover v1.0 Source Release
'
' by: G@Te^k3eP3R (johnge2@yahoo.com)
'
Private Sub cmdexit_Click()
Unload Me
End Sub

