VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy CD Cover (1.0.0)"
   ClientHeight    =   6045
   ClientLeft      =   2580
   ClientTop       =   3045
   ClientWidth     =   9810
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9810
   Visible         =   0   'False
   Begin VB.CommandButton btnprint 
      Caption         =   "&Print Covers"
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   5640
      Width           =   1455
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   9000
      Picture         =   "frmmain.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   13
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton cmdfront 
      Caption         =   "Show Picture"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrFront 
      Caption         =   "Print"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdmiddle 
      Caption         =   "Show Picture"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrMiddle 
      Caption         =   "Print"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrBack 
      Caption         =   "Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Show Picture"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox lstfiles 
      Height          =   255
      ItemData        =   "frmmain.frx":0614
      Left            =   120
      List            =   "frmmain.frx":0616
      TabIndex        =   2
      Top             =   1440
      Width           =   9495
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Reset"
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "CD Back Cover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6600
      TabIndex        =   6
      Top             =   1800
      Width           =   3135
      Begin VB.Image Image3 
         Height          =   2325
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2835
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CD Inside Cover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3360
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
      Begin VB.Image Image2 
         Height          =   2325
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2835
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CD Front Cover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3135
      Begin VB.Image Image1 
         Height          =   2325
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2835
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Easy CD Cover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label Label3 
      Caption         =   $"frmmain.frx":0618
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Cover"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpDlg 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Easy CD Cover v1.0 Source Release
'
' by: G@Te^k3eP3R (johnge2@yahoo.com)
'


Private Sub cmdAbout_Click()
    Show frmAbout, vbModal
End Sub

Private Sub btnprint_Click()
Printer.PaintPicture frmmain.Image3, 1200, 1200, 8560, 6730
Printer.EndDoc
End Sub

Private Sub cmdback_Click()
Dim message1
If lstfiles.List(0) = "" Then
message1 = MsgBox("First you have to drop one or more pictures in the dropbox!!!", vbOKOnly, "Drag and drop a picture...")
Else
Image3.Visible = True
Image3.Picture = LoadPicture(lstfiles.List(0))
lstfiles.RemoveItem (0)
cmdback.Visible = False
cmdPrBack.Visible = True
End If

End Sub

Private Sub cmdclear_Click()
' Clears all the contents
cmdfront.Visible = True
cmdmiddle.Visible = True
cmdback.Visible = True
cmdPrFront.Visible = False
cmdPrMiddle.Visible = False
cmdPrBack.Visible = False

lstfiles.Clear
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False

End Sub

Private Sub cmdexit_Click()
' End the program
End
End Sub

Private Sub cmdfront_Click()
Dim message1
If lstfiles.List(0) = "" Then
message1 = MsgBox("First you have to drop one or more pictures in the dropbox!!!", vbOKOnly, "Drag and drop a picture...")
Else

Image1.Visible = True
Image1.Picture = LoadPicture(lstfiles.List(0))
lstfiles.RemoveItem (0)
cmdfront.Visible = False
cmdmiddle.Visible = True
cmdPrFront.Visible = True
End If

End Sub

Private Sub cmdmiddle_Click()
Dim message1
If lstfiles.List(0) = "" Then
message1 = MsgBox("First you have to drop one or more pictures in the dropbox!!!", vbOKOnly, "Drag and drop a picture...")
Else

Image2.Visible = True
Image2.Picture = LoadPicture(lstfiles.List(0))
lstfiles.RemoveItem (0)
cmdfront.Visible = False
cmdmiddle.Visible = False
cmdback.Visible = True
cmdPrMiddle.Visible = True


End If

End Sub

Private Sub cmdPrBack_Click()
Printer.PaintPicture frmmain.Image3, 1200, 1200, 8560, 6730
Printer.EndDoc

End Sub

Private Sub cmdPrFront_Click()
Printer.PaintPicture frmmain.Image1, 1200, 1200, 6800, 6730
Printer.EndDoc

End Sub


Private Sub cmdPrMiddle_Click()
Printer.PaintPicture frmmain.Image2, 1200, 1200, 6800, 6730
Printer.EndDoc
End Sub

Private Sub Command1_Click()
frmAbout.Show

End Sub

Private Sub Form_Load()
Visible = True

DragAcceptFiles frmmain.hwnd, True
WatchForFiles
End Sub



Public Sub WatchForFiles()
' This subrountine watchs for all of your WM_DROPFILES messages
' Dim Variables
Dim FileDropMessage As MSG      ' Msg Type
Dim fileDropped     As Boolean  ' True if Files where dropped
Dim hDrop           As Long     ' Pointer to the dropped file structure
Dim filename        As String * 128  ' the dropped filename
Dim numOfDroppedFiles As Long   ' the amount of dropped files
Dim curFile         As Long     ' the current file number
Dim ret As Long
' loop to keep checking for files
' NOTE : Do any code you want to execute before this set
Do
  ' check for Dropped file messages
  fileDropped = PeekMessage(FileDropMessage, 0, WM_DROPFILES, WM_DROPFILES, PM_REMOVE Or PM_NOYIELD)
  If fileDropped Then
    ' Get the pointer to the dropped file structure
    hDrop = FileDropMessage.wParam
    ' Get the toal number of files
    numOfDroppedFiles = DragQueryFile(hDrop, True, filename, 127)
    For curFile = 1 To numOfDroppedFiles
      ' Get the file name
      ret = DragQueryFile(hDrop, curFile - 1, filename, 127)
      ' at this pointer you can do what you want with the filename
      ' the filename will be a full qalified path
      'lblnumfiles = LTrim$(Str$(numOfDroppedFiles))
    lstfiles.AddItem filename
       Next curFile
    ' We are now done with the structure, tell windows to discard it
    DragFinish hDrop
  End If
  DoEvents
Loop
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelpDlg_Click()
frmHelp.Show
End Sub

Private Sub mnuPrint_Click()
Printer.PaintPicture frmmain.Image3, 1200, 1200, 8560, 6730
Printer.EndDoc
End Sub

Private Sub mnuReset_Click()
' Clears all the contents
cmdfront.Visible = True
cmdmiddle.Visible = True
cmdback.Visible = True
cmdPrFront.Visible = False
cmdPrMiddle.Visible = False
cmdPrBack.Visible = False

lstfiles.Clear
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub picIcon_Click()
frmAbout.Show
End Sub
