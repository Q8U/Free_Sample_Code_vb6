VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Extracting options"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   4455
      Begin VB.CommandButton ExtrBrowse 
         Caption         =   "Browse"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox ExtrS01 
         Caption         =   "Warn before extracting"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox ExPath 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   690
         Width           =   1935
      End
      Begin VB.OptionButton Extr02 
         Caption         =   "Extract to:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Extr01 
         Caption         =   "Prompt extraction location"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archive viewing options"
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox ArchS02 
         Caption         =   "Archive Grid"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox ArchS01 
         Caption         =   "Archive sort alphabetical"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.OptionButton Arch04 
         Caption         =   "View archive as normal icons (16x16)"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
      End
      Begin VB.OptionButton Arch03 
         Caption         =   "View archive as lists (16x16)"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.OptionButton Arch02 
         Caption         =   "View archive as small icons (16x16)"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
      Begin VB.OptionButton Arch01 
         Caption         =   "View archive as a report (16x16)"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Extr01_Click()
    ExPath.Enabled = False
    ExtrBrowse.Enabled = False
End Sub

Private Sub Extr02_Click()
    ExPath.Enabled = True
    ExtrBrowse.Enabled = True
End Sub

Private Sub ExtrBrowse_Click()
    FrmOptionDir.Show 1, Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    sString = String(100, "*")
    lLength = Len(sString)
    GetPrivateProfileString "Settings", "Report", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Arch01.Value = sString
    GetPrivateProfileString "Settings", "SmallIcons", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Arch02.Value = sString
    GetPrivateProfileString "Settings", "Lists", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Arch03.Value = sString
    GetPrivateProfileString "Settings", "NormalIcons", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Arch04.Value = sString
    GetPrivateProfileString "Settings", "Sort", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    ArchS01.Value = sString
    GetPrivateProfileString "Settings", "Grid", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    ArchS02.Value = sString
    GetPrivateProfileString "Settings", "Prompt", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Extr01.Value = sString
    GetPrivateProfileString "Settings", "SendTo", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    Extr02.Value = sString
    GetPrivateProfileString "Settings", "ExPath", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    If sString = String(100, "*") Or Left(LCase$(sString), 5) = LCase$("False") Then ExPath.Text = "": Extr01.Value = True Else ExPath.Text = sString
    GetPrivateProfileString "Settings", "Warn", vbNullString, sString, lLength, App.Path & "\Settings.ini"
    ExtrS01.Value = sString
End Sub

Private Sub OK_Click()
    If Extr02.Value = True And ExPath.Text = "" Then MessageBox "Please click on the browse button to enter an extraction path.", OKOnly, Critical: Exit Sub
    WritePrivateProfileString "Settings", "Report", CStr(Arch01.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "SmallIcons", CStr(Arch02.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "Lists", CStr(Arch03.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "NormalIcons", CStr(Arch04.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "Sort", CStr(ArchS01.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "Grid", CStr(ArchS02.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "Prompt", CStr(Extr01.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "SendTo", CStr(Extr02.Value), App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "ExPath", ExPath.Text, App.Path & "\Settings.ini"
    WritePrivateProfileString "Settings", "Warn", CStr(ExtrS01.Value), App.Path & "\Settings.ini"
    
    If Arch01.Value = True Then
        frmMain.ListFiles.View = lvwReport
            ElseIf Arch02.Value = True Then
        frmMain.ListFiles.View = lvwSmallIcon
            ElseIf Arch03.Value = True Then
        frmMain.ListFiles.View = lvwList
            ElseIf Arch04.Value = True Then
        frmMain.ListFiles.View = lvwIcon
    End If
    
    If ArchS01.Value = Checked Then frmMain.ListFiles.Sorted = True Else frmMain.ListFiles.Sorted = False
    If ArchS02.Value = Checked Then frmMain.ListFiles.GridLines = True Else frmMain.ListFiles.GridLines = False
    
    If Extr01.Value = True Then ExtractPath = ""
    If Extr02.Value = True Then ExtractPath = ExPath.Text
    If ExtrS01.Value = Checked Then ChkWarningMsg = True Else ChkWarningMsg = False
    
    Unload Me
End Sub
