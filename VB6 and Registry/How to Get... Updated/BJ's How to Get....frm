VERSION 5.00
Begin VB.Form frmHowtoGet 
   Caption         =   "BJ's How to Get..."
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3090
   Icon            =   "BJ's How to Get....frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdMakeINI 
      Caption         =   "Make INI"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Information"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdWinRun 
      Caption         =   "Win Run"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrialVersion 
      Caption         =   "Trial Version"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdWindowsDir 
      Caption         =   "Windows Dir"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrayIcon 
      Caption         =   "Tray Icon"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisableX 
      Caption         =   "Disable [X]"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAboutBox 
      Caption         =   "About Box"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmHowtoGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAboutBox_Click()
 frmAboutBox.Show
End Sub

Private Sub cmdDisableX_Click()
 frmDisable_Close_Button.Show
End Sub

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get...&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub cmdInfo_Click()
Dim Msg, Style, Title, Response, MyString

Msg = "The following demos show." & vbCrLf & vbCrLf & _
"1: How to access the systems About Box. - (BJ's About Box.)" & vbCrLf & _
"2: How to disable the [X] in the top right hand corner. - (BJ's Disable Close Button.)" & vbCrLf & _
"3: How to add an Icon to the Icon Tray next to Time at bottom Right. - (BJ's Tray Icon.)" & vbCrLf & _
"4: How to make a Trial Version Application and wright to the Registry. - (BJ's Trial Vresion.)" & vbCrLf & _
"5: How to get the Windows Directory for your applications. - (BJ's Windows Directory.)" & vbCrLf & _
"6: How to get information on how long Windows has been running for. - (BJ's Win Run.)" & vbCrLf & _
"7: How to make an .ini file. - (BJ's Make INI.)" & vbCrLf & vbCrLf & _
"I hope you find a few things useful for your self. Thanks BJ. Any information E-Mail me: bryce3@bigpond.com"

Style = vbOKOnly + vbInformation

Title = "Information about BJ's How to Get..."

Response = MsgBox(Msg, Style, Title)

If Response = vbOK Then
   MyString = "OK"
End If

End Sub

Private Sub cmdMakeINI_Click()
frmMakeINI.Show
End Sub

Private Sub cmdTrayIcon_Click()
 frmTrayIcon.Show
End Sub

Private Sub cmdTrialVersion_Click()
 frmTrialVersion.Show
End Sub

Private Sub cmdWindowsDir_Click()
 frmWindowsDir.Show
End Sub

Private Sub cmdWinRun_Click()
 frmWin_Run.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub
