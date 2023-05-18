VERSION 5.00
Begin VB.Form frmWindowsDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BJ's How to Get... Windows Directory."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "BJ's Windows Directory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmWindowsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copy all this text you are reading. Select all
' by clicking on the text
' hold mouse button down and go to the end of the Text
' Now press CTRL + C to Copy
' Create new EXE project and paste it into Form1
' Now You will need
' 1: Text box named Text1 which is the "Default Name"
' 1: Command Button named Command1. which is the "Default Name"
' Change Caption to Get Windows Directory
' 1: Command Button named Command2. which is the "Default Name"
' Change Caption to Exit
' Now click Run or F5 and Click on Command Button

'Function to get Windows directory
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Variable to store the Windows directory
Dim WinDir As String

'Buffer and constant used for API functions
Dim msBuffer As String * 255
Const BUFFERSIZE = 255

Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Windows Directory.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub Command2_Click() 'Exit Command Button
Unload frmWindowsDir
End Sub


Private Sub Form_Load()
Dim lBytes As Long

lBytes = GetWindowsDirectory(msBuffer, BUFFERSIZE)
WinDir = Left$(msBuffer, lBytes)

Label1.Caption = WinDir  'Which = above Default = C:\Windows

End Sub
