VERSION 5.00
Begin VB.Form FrmOptionDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Directory"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "FrmOptionDir.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton ExitCmd 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.DriveListBox DriveDir 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Drive"
      Top             =   120
      Width           =   3975
   End
   Begin VB.DirListBox DirS 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Extraction Directory"
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "FrmOptionDir"
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

'Click on OK set the extraction path
Private Sub OK_Click()
    FrmOptions.ExPath.Text = DirS.Path
    Unload Me
End Sub

