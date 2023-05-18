VERSION 5.00
Begin VB.Form frmOpenFile 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open File:"
   ClientHeight    =   3180
   ClientLeft      =   90
   ClientTop       =   2190
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmOpen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   432
      Left            =   2400
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   2412
   End
   Begin VB.CommandButton CmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   372
      Left            =   5640
      TabIndex        =   5
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Height          =   372
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1452
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2412
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1890
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2412
   End
   Begin VB.Label lblFormatStatus 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   2052
   End
   Begin VB.Label lblDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drives:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label lblDirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Directories:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File &Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "frmOpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
LoadPath = File1.Path
Unload frmOpenFile
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandler
If picIcon2.Visible = True Then GoTo SkipCheck
'Trim spaces from both ends
Text1.Text = RTrim$(LTrim$(Text1.Text))
If (Text1.Text = ".") Or (Text1.Text = "") Then Exit Sub
If Text1.Text = ".." Then
  ChDir$ ".."
  Dir1.Path = CurDir$
  Exit Sub
End If
If Text1.Text = "\" Then
  Dir1.Path = Left$(Drive1.Drive, 2) + "\"
  Exit Sub
End If
PathText = Text1.Text

'Drive letter in text?
If (Mid$(PathText, 2, 1) = ":") Then
  ChDrive Left$(PathText, 2)
  Drive1.Drive = Left$(PathText, 2)
  If File1.ListCount > 0 Then
    Text1.Text = File1.List(0)
    Call CheckFileFormat(Text1.Text)
    Else
    picIcon2.Visible = False
    frmOpenFile.lblFormatStatus.Caption = ""
    Text1.Text = ""
  End If
  Exit Sub
End If

'If "\" at end of text, cut it off.
If Right$(PathText, 1) = "\" Then PathText = Left$(PathText, Len(PathText) - 1)

'Ignore pattern changes
If InStr(Text1.Text, "*") > 0 Then
  File1.Pattern = Text1.Text
  Text1.SetFocus
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Exit Sub
End If

TempDrive = Left$(Dir1.Path, 3)
TempDirPath = Mid$(Dir1.Path, 4)
If TempDirPath > "" Then TempDirPath = TempDirPath + "\"

Call GetFileTextInfo

'TextStatus:
' 0 = Invalid entry.
' 1 = Valid path only, no filename.
' 2 = Valid path, filename doesn't exists.
' 3 = Valid path, filename exists.

If (TextStatus = 0) Or (TextStatus = 2) Then
  A% = MsgBox("Invalid Entry", 48, "Error")
  Text1.SetFocus
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Exit Sub
End If
If TextStatus = 1 Then
  'Directory only.
  If Left$(CurDir$, 1) <> Left$(Dir1.Path, 1) Then ChDrive Left$(Dir1.Path, 2)
  Dir1.Path = TempDrive + TempDirPath
  Exit Sub
End If
If TextStatus = 3 Then
  'Valid path, filename exists.
  Dir1.Path = TempDrive + TempDirPath
  If Left$(CurDir, 1) <> Left$(Dir1.Path, 1) Then
    ChDrive Left$(Dir1.Path, 1)
  End If
  ChDir TempDrive + TempDirPath
  Text1.Text = TempFileName
  LoadName = TempFileName
  Call CheckFileFormat(Text1.Text)
End If
If FileFormat = 0 Then Text1.SetFocus: Exit Sub
SkipCheck:
LoadPath = CurDir$
LoadName = Text1.Text
Unload frmOpenFile
Call LoadIconFile
Exit Sub

ErrHandler:
e% = Err
A% = MsgBox("Invalid Entry" & Err.Description, 48, "Error")
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
ChDir Dir1.Path
If File1.ListCount > 0 Then
  Text1.Text = File1.List(0)
  Call CheckFileFormat(Text1.Text)
  Else
  picIcon2.Visible = False
  frmOpenFile.lblFormatStatus.Caption = ""
  Text1.Text = ""
End If
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrHandler
TmpDrive$ = Left$(Dir1.Path, 2)
ChDrive Left$(Drive1.Drive, 2)
Dir1.Path = Drive1.Drive
Exit Sub
ErrHandler:
A% = MsgBox("Drive " & Drive1.Drive & " not available.", vbOKOnly, "Oops!")
Drive1.Drive = TmpDrive$
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
Call CheckFileFormat(Text1.Text)
End Sub

Private Sub File1_DblClick()
LoadName = File1.FileName
TempDrive = Left$(Dir1.Path, 3)
TempDirPath = Mid$(Dir1.Path, 4)
If TempDirPath > "" Then TempDirPath = TempDirPath + "\"
TempFileName = ""
Call CheckFileFormat(Text1.Text)
If FileFormat = 1 Then
  FileFoundFlag = True
  Unload frmOpenFile
  frmMain.Enabled = True
  frmMain.SetFocus
  Call LoadIconFile
End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call cmdOK_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  LoadPath = File1.Path
  Unload frmOpenFile
  frmMain.Enabled = True
  frmMain.SetFocus
End If
End Sub
Private Sub Form_Load()
'Load File
FormOpenLoadedFlag = True
FileFoundFlag = False
frmMain.Enabled = False
If TempPattern = "" Then TempPattern = "*.ico;*.bmp"
File1.Pattern = TempPattern
'Adjust size to original
Width = Int(7356 * Screen.TwipsPerPixelX / 12 + 0.5)
Height = Int(3504 * Screen.TwipsPerPixelY / 12 + 0.5)
ScaleWidth = 607
ScaleHeight = 267
'Center form
Left = (Screen.Width - Width) / 2
Top = frmMain.Top + frmMain.cmdRefresh.Top + 75 * Screen.TwipsPerPixelY
lblFileName.Left = 20
lblFileName.Top = 10
lblFileName.Width = 131
lblFileName.Height = 21
lblDirectory.Left = 250
lblDirectory.Top = 10
lblDirectory.Width = 101
lblDirectory.Height = 16
lblDrives.Left = 250
lblDrives.Top = 210
lblDrives.Width = 101
lblDrives.Height = 16
lblFormatStatus.Left = 20
lblFormatStatus.Top = 233
Text1.Left = 20
Text1.Top = 30
Text1.Width = 201
Text1.Height = 24
File1.Left = 20
File1.Top = 60
File1.Width = 201
File1.Height = 162
Dir1.Left = 250
Dir1.Top = 30
Dir1.Width = 201
Dir1.Height = 171
Drive1.Left = 250
Drive1.Top = 230
Drive1.Width = 201
'Drive1.Height = 24 (Read only!)
cmdOK.Left = 470
cmdOK.Top = 10
cmdOK.Width = 121
cmdOK.Height = 31
CmdCancel.Left = 470
CmdCancel.Top = 60
CmdCancel.Width = 121
CmdCancel.Height = 31
picIcon2.Left = 180
picIcon2.Top = 226
picIcon2.Width = 36
picIcon2.Height = 36
picIcon2.ScaleWidth = 32
picIcon2.ScaleHeight = 32

Visible = True
If LoadPath = "" Then
  LoadPath = File1.Path
  Else
  Drive1.Drive = Left$(LoadPath, 2)
  Dir1.Path = LoadPath
End If
FileFormat = 0
If File1.ListCount > 0 Then
  Text1.Text = File1.List(0)
  Call CheckFileFormat(Text1.Text)
  Else: Text1.Text = ""
End If
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormOpenLoadedFlag = False
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FileFormat = 1
  Call cmdOK_Click
  Exit Sub
End If
frmOpenFile.cmdOK.Enabled = True
picIcon2.Visible = False
frmOpenFile.lblFormatStatus.Caption = ""
If KeyAscii = 27 Then
  Unload frmOpenFile
  frmMain.Enabled = True
  frmMain.SetFocus
End If
End Sub

