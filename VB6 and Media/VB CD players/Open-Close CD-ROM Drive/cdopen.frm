VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CD Open/Close By Sam Huggill 17/4/98"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "cdopen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close the CD ROM"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open The CD ROM"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long

Private Sub Command1_Click()
retvalue = mciSendString("set CDAudio door open", _
returnstring, 127, 0)

End Sub

Private Sub Command2_Click()
retvalue = mciSendString("set CDAudio door closed", _
returnstring, 127, 0)

End Sub

