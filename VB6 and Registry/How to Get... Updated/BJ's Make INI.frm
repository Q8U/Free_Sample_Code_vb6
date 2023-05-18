VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMakeINI 
   Caption         =   "BJ's How to Get... Make INI"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "BJ's Make INI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail me"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbShowINI 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7223
      _Version        =   393217
      TextRTF         =   $"BJ's Make INI.frx":0442
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show INI"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdMakeINI 
      Caption         =   "Make INI"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMakeINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Make INI.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

Private Sub cmdExit_Click()
Unload frmMakeINI
End Sub

Private Sub cmdMakeINI_Click()
    bj = WritePrivateProfileString(SECTION, ENTRY, " How to Get... Entry", INI_FILE)
        bj = GetPrivateProfileString(SECTION, ENTRY, " How to Get... Entry", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION, ENTRY1, " BJ's How to Get... Entry1", INI_FILE)
        bj = GetPrivateProfileString(SECTION, ENTRY1, " BJ's How to Get... Entry1", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION, ENTRY2, " BJ's How to Get... Entry2", INI_FILE)
        bj = GetPrivateProfileString(SECTION, ENTRY2, " BJ's How to Get... Entry2", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION1, ENTRY3, " BJ's How to Get... Entry3", INI_FILE)
        bj = GetPrivateProfileString(SECTION1, ENTRY3, " BJ's How to Get... Entry3", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION2, ENTRY4, " BJ's How to Get... Entry4", INI_FILE)
        bj = GetPrivateProfileString(SECTION2, ENTRY4, " BJ's How to Get... Entry4", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION3, ENTRY5, " BJ's How to Get... Entry5", INI_FILE)
        bj = GetPrivateProfileString(SECTION3, ENTRY5, " BJ's How to Get... Entry5", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION3, ENTRY6, " BJ's How to Get... Entry6", INI_FILE)
        bj = GetPrivateProfileString(SECTION3, ENTRY6, " BJ's How to Get... Entry6", bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION4, ENTRY7, " " & App.Path, INI_FILE)
        bj = GetPrivateProfileString(SECTION4, ENTRY7, " " & App.Path, bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION4, ENTRY8, " " & App.EXEName, INI_FILE)
        bj = GetPrivateProfileString(SECTION4, ENTRY8, " " & App.EXEName, bj, Len(bj), INI_FILE)
    
    bj = WritePrivateProfileString(SECTION4, ENTRY9, " " & App.Major & "." & App.Minor & "." & App.Revision, INI_FILE)
        bj = GetPrivateProfileString(SECTION4, ENTRY9, " " & App.Major & "." & App.Minor & "." & App.Revision, bj, Len(bj), INI_FILE)

End Sub

Private Sub cmdShow_Click()
rtbShowINI.FileName = "C:\Windows\" & INI_FILE
End Sub
