VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extracting data"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   3000
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   340
      Left            =   4320
      TabIndex        =   2
      Top             =   1520
      Width           =   1160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   190
      Left            =   165
      TabIndex        =   1
      Top             =   1515
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -120
      Top             =   1440
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      _Version        =   393216
      FullWidth       =   353
      FullHeight      =   49
   End
   Begin VB.Label Label1 
      Caption         =   "Extracting files..."
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unZipFile As New Unzip
Dim TempPath As String
Dim ExecPath As String
Private Sub Command1_Click()
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 2

'Label1.Caption = unZipFile.GetLastMessage
Me.Refresh
Show
End Sub


Private Sub Form_Load()
Me.Show
Me.Refresh
LoadAVI
End Sub

Function LoadAVI()
Dim Temp As String
Dim MyAVI As String
Dim MyZip As String
Open App.path & "\" & App.EXEName & ".exe" For Binary As #1
    Temp = String(LOF(1), "x")
    Get #1, , Temp
Close #1
pos1 = InStr(1, Temp, "<start AVI>")
pos1 = pos1 + 11
pos2 = InStr(1, Temp, "<end AVI>")
MyAVI = Mid(Temp, pos1, pos2 - pos1)

Randomize Timer
TempPath = "C:\windows\temp\avi" & CStr(Rnd(1000) * 1)
Destroy TempPath
Open TempPath For Binary As #1
Put #1, , MyAVI
Close #1
Animation1.Open TempPath
Animation1.Play
Destroy TempPath
zippos1 = InStr(1, Temp, "<start compressed file>")
zippos1 = zippos1 + 23
zippos2 = InStr(1, Temp, "<end compressed file>")
MyZip = Mid(Temp, zippos1, zippos2 - zippos1)
Destroy "C:\Windows\Temp\Extract.zip"

Open "C:\Windows\Temp\Extract.zip" For Binary As #1
    Put #1, , MyZip
Close #1
xpos1 = InStr(1, Temp, "<start execute command>")
xpos1 = xpos1 + 23
xpos2 = InStr(1, Temp, "<end execute command>")

ExecPath = Replace(Mid(Temp, xpos1, xpos2 - xpos1), "%extractdir%", "C:\Windows\Temp\Install")

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Destroy TempPath
End Sub

Function Destroy(xpath As String)
On Error Resume Next
Kill xpath
End Function

Function MakeDir(path As String)
On Error Resume Next
MkDir path
End Function


Private Sub Timer2_Timer()
Timer2.Enabled = False
Dim unZipFile As New Unzip
MakeDir "C:\Windows\Temp\Install\"
unZipFile.ZipFileName = "C:\Windows\Temp\Extract.zip"
unZipFile.ExtractDir = "C:\Windows\Temp\Install\"
unZipFile.Unzip
Destroy "C:\Windows\Temp\Extract.zip"
ProgressBar1.Value = 100
Shell ExecPath, vbNormalFocus
End
End Sub
