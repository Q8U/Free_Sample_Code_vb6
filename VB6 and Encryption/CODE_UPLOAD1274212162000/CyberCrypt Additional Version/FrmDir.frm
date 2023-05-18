VERSION 5.00
Begin VB.Form FrmDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save to Directory"
   ClientHeight    =   4245
   ClientLeft      =   1950
   ClientTop       =   1155
   ClientWidth     =   4215
   Icon            =   "FrmDir.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox DirS 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Extraction Directory"
      Top             =   480
      Width           =   3975
   End
   Begin VB.DriveListBox DriveDir 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Drive"
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton ExitCmd 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "FrmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Changes the drive, if the drive doesn't open the sub returns
'an error and sets the drive back to the lastdrive used
Private Sub DriveDir_Change()
    On Error GoTo FinaliseError
    DirS.Path = DriveDir
    Exit Sub
FinaliseError:
    MessageBox "Current drive not avialable.", OKOnly, Critical
    DriveDir = LastDrive
End Sub

Private Sub ExitCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LastDrive = DriveDir.Drive
End Sub

'Click on OK starts the extraction of the selected items to extract
Private Sub OK_Click()

    On Error GoTo FinaliseError
    
    For M = 1 To frmMain.ListFiles.ListItems.Count
        frmMain.Caption = "CyberCrypt Extracted add to (" & DirS.Path & ") from (" & ArchiveName & ")"
        If FileExist(DirS.Path & "\" & frmMain.ListFiles.ListItems(M)) = True Then Kill DirS.Path & "\" & frmMain.ListFiles.ListItems(M)
        frmBusy.Show
        Me.Enabled = False
        Me.MousePointer = 11
        If frmMain.CyTExtract(CyTFile, frmMain.ListFiles.ListItems(M), DirS.Path & "\" & frmMain.ListFiles.ListItems(M)) = False Then MessageBox "An error occured when trying to extract the file(s)!", OKOnly, Critical: Unload frmBusy: Exit For
        Me.Enabled = True
        Me.MousePointer = 0
        Unload frmBusy
    Next M
    Unload Me
    Exit Sub
    
FinaliseError:
    MessageBox "An error occured when trying to extract the file(s)!", OKOnly, Critical
    Unload frmBusy
    Exit Sub
End Sub
