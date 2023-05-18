VERSION 5.00
Begin VB.Form frmSaveAs 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save As..."
   ClientHeight    =   3180
   ClientLeft      =   4545
   ClientTop       =   4125
   ClientWidth     =   7260
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
   Icon            =   "frmSave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2412
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2412
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
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   2412
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
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
Unload frmSaveAs
ExitAfterSaveFlag = False
LoadAfterSaveFlag = False
NewAfterSaveFlag = False
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandler

'Trim spaces from both ends
Text1.Text = RTrim$(LTrim$(Text1.Text))
If (Text1.Text = ".") Or (Text1.Text = "") Then
  Text1.Text = ""
  Exit Sub
End If
If Text1.Text = ".." Then
  ChDir$ ".."
  Dir1.Path = CurDir$
  Exit Sub
End If
If Text1.Text = "\" Then
  Dir1.Path = Left$(Drive1.Drive, 2) + "\"
  Text1.Text = ""
  Exit Sub
End If
PathText = Text1.Text

'Drive letter in text?
If (Mid$(PathText, 2, 1) = ":") Then
  ChDrive Left$(PathText, 1)
  Drive1.Drive = Left$(PathText, 1)
  Exit Sub
End If

'If "\" at end of text, cut it off.
If Right$(PathText, 1) = "\" Then PathText = Left$(PathText, Len(PathText) - 1)

'Ignore pattern changes
If InStr(Text1.Text, "*") > 0 Then
  Text1.Text = "*.ico"
  Exit Sub
End If

TempDrive = Left$(Dir1.Path, 3)
TempDirPath = Mid$(Dir1.Path, 4)
If TempDirPath > "" Then TempDirPath = TempDirPath + "\"

Call GetFileTextInfo
If TextStatus = 0 Then
  A% = MsgBox("Invalid Entry", 48, "Error")
  Text1.SetFocus
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Exit Sub
End If
If TextStatus > 0 Then
  'No errors, path OK.
  If Left$(CurDir$, 1) <> Left$(Dir1.Path, 1) Then
    'Make current drive same as Dir1
    ChDrive Left$(Dir1.Path, 2)
  End If
  Dir1.Path = TempDrive + TempDirPath
  SavePath = Dir1.Path
  SaveName = TempFileName
End If
If SaveName = "" Then
  SaveName = LoadName
  If LCase$(Right$(SaveName, 4)) = ".bmp" Then
    SaveName = Left$(SaveName, Len(SaveName) - 4) + ".ico"
  End If
  Text1.Text = SaveName
  Text1.SetFocus
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Exit Sub
End If
If Dir(SaveName) > "." Then
  Open SaveName For Binary As #1
  Get #1, 1, FileHeader
  Close
  NumberOfIcons = Asc(Mid$(FileHeader, 5, 1))
  If NumberOfIcons > 1 Then
    A% = MsgBox("Destination file contains multiple icons." & Chr$(13) & "Please choose a different name.", 48, "Save Error")
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Exit Sub
  End If
  A% = MsgBox("Replace Existing File?", 563, "File Already Exists")
  If A% = 7 Then
    'No
    Load frmSaveAs
    Exit Sub
  End If
  If A% = 2 Then
    'Cancel
    LoadAfterSaveFlag = False
    ExitAfterSaveFlag = False
    NewAfterSaveFlag = False
    Unload frmSaveAs
    frmMain.Enabled = True
    frmMain.SetFocus
    Exit Sub
  End If
  'Yes
End If
Call SaveIconFile
Exit Sub

ErrHandler:
e% = Err
A% = MsgBox("Invalid Entry", 48, "Error")
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
ChDir Dir1.Path
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrHandler
Dir1.Path = Drive1.Drive
ErrHandler:
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
End Sub

Private Sub File1_DblClick()
SaveName = File1.FileName
Call SaveIconFile
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call cmdOK_Click
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  LoadAfterSaveFlag = False
  NewAfterSaveFlag = False
  Unload frmSaveAs
  frmMain.Enabled = True
  frmMain.SetFocus
End If
End Sub

Private Sub Form_Load()
'Save As

frmMain.Enabled = False
File1.Pattern = "*.ico"

'Adjust size to original
Width = Int(7356 * Screen.TwipsPerPixelX / 12 + 0.5)
Height = Int(3504 * Screen.TwipsPerPixelY / 12 + 0.5)

ScaleWidth = 605
ScaleHeight = 265

'Center form
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

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
Visible = True
If SavePath = "" Then
  SavePath = File1.Path
  Else
  Drive1.Drive = Left$(SavePath, 2)
  Dir1.Path = SavePath
End If
A% = InStr(SaveName, ".")
If A% = 0 Then SaveName = Left$(SaveName, A% - 1) + ".ico"
If (SaveName > "") And (SaveName <> "UNTITLED.ico") Then
  For A% = 0 To File1.ListCount - 1
  If File1.List(A%) = SaveName Then
    File1.ListIndex = A%
    Exit For
  End If
  Next A%
End If
Text1.Text = SaveName
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call cmdOK_Click
  Exit Sub
End If
If KeyAscii = 27 Then
  Unload frmSaveAs
  frmMain.Enabled = True
  frmMain.SetFocus
End If
End Sub

