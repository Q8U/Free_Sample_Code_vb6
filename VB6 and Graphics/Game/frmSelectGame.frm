VERSION 5.00
Object = "{D2D9B7C1-7650-11D1-9481-00A0247B7657}#1.0#0"; "ZLIBOCX2.DLL"
Begin VB.Form frmSelectGame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   555
   ClientTop       =   1275
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5835
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin ZLIBOCX2LibCtl.zlibIF zlibUnzipper 
      Height          =   330
      Left            =   1035
      OleObjectBlob   =   "frmSelectGame.frx":0000
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   4185
      TabIndex        =   9
      Top             =   3900
      Width           =   1200
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2790
      TabIndex        =   8
      Top             =   3900
      Width           =   1200
   End
   Begin VB.DirListBox dirDir1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1890
      Left            =   3915
      TabIndex        =   3
      Top             =   1305
      Width           =   2625
   End
   Begin VB.FileListBox filFile1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2235
      Left            =   1800
      TabIndex        =   2
      Top             =   1305
      Width           =   1965
   End
   Begin VB.DriveListBox drvDrive1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3915
      TabIndex        =   1
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   945
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LOAD GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   3300
      TabIndex        =   7
      Top             =   135
      Width           =   1950
   End
   Begin VB.Label lblDirName2 
      BackColor       =   &H00000000&
      Caption         =   "Directory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3960
      TabIndex        =   6
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label lblDirName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3945
      TabIndex        =   5
      Top             =   945
      Width           =   3960
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00000000&
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1800
      TabIndex        =   4
      Top             =   675
      Width           =   1290
   End
End
Attribute VB_Name = "frmSelectGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Private Sub cmdCancel_Click()
'return to cover form - prevent game crashing
'when there is no game to load
LoadCancelled = True
Unload Me


End Sub


Private Sub cmdStart_Click()
If txtFileName = "" Then
    'no game selected
    PlaySoundEffect "Quiet"
    MsgBox "Please select a game to load"
    Exit Sub
Else
    'decompress selected saved game file
    Dim Path As String
    Path = dirDir1.Path
    
    If Right(Path, 1) <> "\" Then
        Path = Path + "\"
    End If
       
    zlibUnzipper.InputFileName = Path + txtFileName.Text
    zlibUnzipper.OutputFileName = App.Path + "\gameinfo.txt"
    zlibUnzipper.Decompress
    GameNumber = Val(Mid$(txtFileName.Text, 2, 1))
    
    Unload Me  '***program returns to frmCover to read big file, then loads frmGameScreen
    
End If

End Sub

Private Sub dirDir1_Change()
'update file list box with new directory
filFile1.Path = dirDir1.Path

'update dir label
lblDirName.Caption = dirDir1.Path

End Sub

Private Sub drvDrive1_Change()
On Error GoTo DriveError

'change path of dir list box to new drive
dirDir1.Path = drvDrive1.Drive
Exit Sub

DriveError:
PlaySoundEffect "Warning"
MsgBox "Drive Error", , " "

'restore the original drive
drvDrive1.Drive = dirDir1.Path
Exit Sub

End Sub

Private Sub filFile1_Click()
txtFileName.Text = filFile1.Filename

End Sub

Private Sub filFile1_DblClick()
txtFileName.Text = filFile1.Filename
cmdStart_Click

End Sub


Private Sub Form_Activate()
Randomize
'draw white stars on the screen
    Dim a, X, Y
    For a = 1 To 400
        X = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        Me.PSet (X, Y), vbWhite
    Next a
    
    'draw dark grey stars
    Dim grey
    grey = &H808080
    For a = 1 To 400
        X = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        Me.PSet (X, Y), grey
    Next a
    
    'draw blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 200
       X = Int(Rnd * Me.ScaleWidth)
       Y = Int(Rnd * Me.ScaleHeight)
       Me.PSet (X, Y), blue
    Next a

End Sub






Private Sub Form_Load()
'set up drives, set to look only for
'compressed game files ending in .zlb
drvDrive1.Drive = App.Path
dirDir1.Path = App.Path
filFile1.Pattern = "*.zlb"



End Sub


