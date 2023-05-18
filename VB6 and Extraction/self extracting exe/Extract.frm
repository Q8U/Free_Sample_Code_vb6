VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.Animation Animation1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      _Version        =   393216
      FullWidth       =   289
      FullHeight      =   65
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempPath As String

Private Sub Form_Load()
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

Dim unZipFile As New Unzip
MakeDir "C:\Windows\Temp\Install\"
unZipFile.ZipFileName = "C:\Windows\Temp\Extract.zip"
unZipFile.ExtractDir = "C:\Windows\Temp\Install\"
unZipFile.Unzip
Destroy "C:\Windows\Temp\Extract.zip"
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

