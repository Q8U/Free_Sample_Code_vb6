VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TabForm 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register It"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1860
      Index           =   2
      Left            =   90
      ScaleHeight     =   1800
      ScaleWidth      =   2310
      TabIndex        =   3
      Top             =   480
      Width           =   2370
      Begin RichTextLib.RichTextBox RichTe 
         Height          =   6555
         Left            =   60
         TabIndex        =   41
         Top             =   60
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   11562
         _Version        =   393217
         BackColor       =   -2147483638
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"TabForm.frx":0000
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   450
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   0
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   360
      Width           =   1815
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   8220
         TabIndex        =   31
         Text            =   "*."
         Top             =   6150
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3510
         TabIndex        =   30
         Top             =   6480
         Width           =   195
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1545
         TabIndex        =   29
         Top             =   6480
         Width           =   210
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   900
         TabIndex        =   28
         Top             =   6480
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   6480
         Width           =   210
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2190
         TabIndex        =   26
         Top             =   6480
         Width           =   210
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2850
         TabIndex        =   25
         Top             =   6480
         Width           =   210
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Paste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10140
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Saves the contents of the above window to the result window"
         Top             =   6450
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reg It"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Attempts to register the highlighted file with your computers registry."
         Top             =   6150
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unreg It"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6450
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reg All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10140
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6150
         Width           =   915
      End
      Begin VB.ListBox lstFoundFiles 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   4740
         Left            =   3060
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   1350
         Width           =   7965
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   1260
         Left            =   3060
         TabIndex        =   9
         Top             =   60
         Width           =   7965
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   30
         TabIndex        =   7
         Top             =   990
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   4815
         Left            =   30
         TabIndex        =   6
         Top             =   1290
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pre-defined file types to be searched for.These files can be registerd"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   42
         Top             =   6180
         Width           =   4800
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Type that will be searched for"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5730
         TabIndex        =   40
         Top             =   6180
         Width           =   2370
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OCX+DLL+OLE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3780
         TabIndex        =   39
         Top             =   6480
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OLE"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1815
         TabIndex        =   38
         Top             =   6480
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLL"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1170
         TabIndex        =   37
         Top             =   6480
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OCX"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   36
         Top             =   6480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OLB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2460
         TabIndex        =   35
         Top             =   6480
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AX"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   34
         Top             =   6480
         Width           =   210
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7590
         TabIndex        =   33
         Top             =   6420
         Width           =   165
      End
      Begin VB.Label lblfound 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Files Found from search:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5670
         TabIndex        =   32
         Top             =   6450
         Width           =   1845
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter File type and click search or chose a pre-defined file below"
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   30
         TabIndex        =   10
         Top             =   90
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   1245
      Index           =   3
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   1245
      TabIndex        =   46
      Top             =   360
      Width           =   1305
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use this window to register all your audio plugins and find all the files of the specified types on your system. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   735
         Left            =   1200
         TabIndex        =   51
         Top             =   960
         Width           =   8745
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1245
         Left            =   2850
         TabIndex        =   50
         Top             =   3750
         Width           =   5445
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1245
         Left            =   2610
         TabIndex        =   49
         Top             =   5220
         Width           =   6405
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1695
         Left            =   1200
         TabIndex        =   48
         Top             =   1860
         Width           =   8865
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "The Sound Engineers Companion Reg It"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   90
         TabIndex        =   47
         Top             =   150
         Width           =   10935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   6765
      Index           =   1
      Left            =   120
      ScaleHeight     =   6705
      ScaleWidth      =   11085
      TabIndex        =   2
      Top             =   390
      Width           =   11145
      Begin VB.ListBox List2 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   2400
         Left            =   30
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   11010
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Clears list below."
         Top             =   2910
         Width           =   930
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   30
         TabIndex        =   11
         Text            =   "Empty"
         Top             =   2880
         Width           =   11025
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Text            =   "This window displays the results for files that have been registerd and the files that could not be registerd"
         Top             =   6330
         Width           =   11055
      End
      Begin VB.TextBox TypeBox 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1530
         TabIndex        =   19
         Top             =   2610
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1530
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1530
         TabIndex        =   17
         Top             =   2220
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Clears list below."
         Top             =   60
         Width           =   915
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Text            =   "Empty"
         Top             =   30
         Width           =   11025
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   2985
         Left            =   30
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   3270
         Width           =   11040
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7215
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   12726
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reg It"
            Key             =   "picture1p"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Register Results"
            Key             =   "picture2p"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Results"
            Key             =   "picture3p"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "picture4p"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SearchFlag As Integer

Private Sub Picture1p()
End Sub
Private Sub Picture2p()
End Sub
Private Sub Picture3p()
End Sub
Private Sub Picture4p()
End Sub

Private Sub Command10_Click()
    Dim Cancel As Boolean
    On Error GoTo ErrorHandler
    Cancel = False
    CommonDialog1.Filter = "S.E.C Files (*.SEC)|*.sec|RichText Files (*.rtf)|*.rtf|Text Files (*.TXT)|*.txt|All Files (*.*)|*.*|"
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNExplorer
    CommonDialog1.Flags = cdlOFNLongNames
    CommonDialog1.Flags = cdlOFNNoDereferenceLinks
    CommonDialog1.Flags = cdlOFNNoValidate
    CommonDialog1.Flags = cdlOFNReadOnly
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    CommonDialog1.Color = RichTe.BackColor
    CommonDialog1.ShowOpen
    If Not Cancel Then
    'load a file
     If UCase(Right(CommonDialog1.FileName, 3)) = "RTF" Then
     RichTe.LoadFile CommonDialog1.FileName, rtfRTF
     Else
     RichTe.LoadFile CommonDialog1.FileName, rtfText
     End If
     End If
ErrorHandler:
If Err.Number = cdlCancel Then
Cancel = True
Resume Next
End If
End Sub

Private Sub Command6_Click()
   Dim Cancel As Boolean
    On Error GoTo ErrorHandler
    Cancel = False
    CommonDialog1.DefaultExt = ".SEC"
    CommonDialog1.Filter = "S.E.C Files (*.SEC)|*.sec|RichText Files (*.rtf)|*.rtf|Text Files (*.TXT)|*.txt|All Files (*.*)|*.*|"
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNCreatePrompt
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    CommonDialog1.Color = RichTe.BackColor
    CommonDialog1.ShowSave
    If Not Cancel Then
    'load a file
     If UCase(Right(CommonDialog1.FileName, 3)) = "RTF" Then
     RichTe.SaveFile CommonDialog1.FileName, rtfRTF
     Else
     RichTe.SaveFile CommonDialog1.FileName, rtfText
     End If
     End If
ErrorHandler:
If Err.Number = cdlCancel Then
Cancel = True
Resume Next
End If
End Sub

Private Sub Form_Load()
TabStrip1.Tabs("picture1p").Selected = True
Option1.Value = True
Option1_Click
End Sub

Private Sub TabStrip1_Click()
On Error GoTo 10
Picture1(TabStrip1.SelectedItem.Index - 1).Move TabStrip1.ClientLeft _
, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight

Picture1(TabStrip1.SelectedItem.Index - 1).ZOrder
Select Case TabStrip1.SelectedItem.Index

    Case 1
        Picture1p
    Case 2
        Picture2p
    Case 3
        Picture3p
    Case 4
        Picture4p
End Select
10: Exit Sub
End Sub


Private Sub Command1_Click()
If Text1.Text = "" Then
Text3.Text = "Please select a file to be register"
Exit Sub
End If

Dim bRet As Boolean
    bRet = RegisterDLL(Text1.Text, True)
If bRet = True Then
    List1.AddItem (TypeBox) & ".ver." & CheckFileVersion(Text1.Text)
Else
    List2.AddItem (TypeBox) & ".ver." & CheckFileVersion(Text1.Text)
End If
updat

End Sub

Private Sub Command2_Click()
On Error GoTo 10
If Text1.Text = "" Then
Text3.Text = "Please select a file to be removed from your registery."
Exit Sub
End If
Dim bRet As Boolean
    bRet = RegisterDLL(Text1.Text, False)
If bRet = True Then
    Text3.Text = (Text1.Text & " ....Has been removed from your registery")
Else
    Text3.Text = (Text1.Text & " ....Could NOT be removed from your registery")
End If
10: Exit Sub
End Sub
Function updat()
If List1.ListCount <= 1 Then
Text4.Text = List1.ListCount & "  Has been Registerd"
Text5.Text = List2.ListCount & "  Could not be Registerd"
Else
Text4.Text = List1.ListCount & "  Have been Registerd"
Text5.Text = List2.ListCount & "  Could not be Registerd"
End If
End Function
Private Sub Command3_Click()
If File1.ListCount < 1 Then
Check2
Else
check
End If
End Sub
Function Check2()
Dim I As Integer
For I = 0 To lstFoundFiles.ListCount - 1
lstFoundFiles.Selected(I) = True
Com_Click
DoEvents
Next I
End Function
Private Sub Com_Click()
Dim bRet As Boolean
    bRet = RegisterDLL(lstFoundFiles.Text, True)
If bRet = True Then
    List1.AddItem (lstFoundFiles.Text) & ".ver." & CheckFileVersion(lstFoundFiles.Text)
Else
    List2.AddItem (lstFoundFiles.Text) & ".ver." & CheckFileVersion(lstFoundFiles.Text)
End If
updat
End Sub

Private Sub Command4_Click()
List2.Clear
Text5.Text = "Empty"
End Sub

Private Sub Command5_Click()
List1.Clear
Text4.Text = "Empty"
End Sub

Private Sub Command7_Click()
Dim I As Integer
If lstFoundFiles.ListCount < 0 Then Exit Sub Else
RichTe.Text = ""
For I = 0 To lstFoundFiles.ListCount - 1
lstFoundFiles.Selected(I) = True
RichTe.SelStart = RichTe.Width + 1
RichTe.SelRTF = lstFoundFiles.Text & "                                                                                                                                                                                                                                                                                                "
DoEvents
Next I
End Sub

Function check()
Dim I As Integer
For I = 0 To File1.ListCount - 1
File1.Selected(I) = True
Command1_Click
DoEvents
Next I
End Function

Private Sub dir1_Change()
    File1.Path = Dir1.Path
    ChDir Dir1.Path
    Text2.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo 10
    Dir1.Path = Drive1.Drive
    Dir1.Path = Drive1.Drive 'Displays path of drive1 in dir1
    ChDrive Drive1.Drive 'The chrdrive statement changes the directorys
10: Exit Sub
End Sub

Private Sub File1_Click()
    Text2.Text = File1.Path
    TypeBox.Text = File1.FileName
    Text1.Text = (Text2.Text & "\" & TypeBox)
End Sub


Private Sub lstFoundFiles_Click()
Text1.Text = lstFoundFiles.Text
End Sub

Private Sub Option1_Click()
    File1.Pattern = "*.Ocx"
    Text6.Text = "*.Ocx"
End Sub

Private Sub Option2_Click()
    File1.Pattern = "*.Dll"
    Text6.Text = "*.Dll"
End Sub

Private Sub Option3_Click()
    File1.Pattern = "*.Ole"
    Text6.Text = "*.Ole"
End Sub

Private Sub Option4_Click()
    File1.Pattern = "*.Ole;*.Dll;*.Ocx"
    Text6.Text = "*.Ole;*.Dll;*.Ocx"
End Sub

Private Sub Option8_Click()
    File1.Pattern = "*.Olb"
    Text6.Text = "*.Olb"
End Sub

Private Sub Option9_Click()
    File1.Pattern = "*.AX"
    Text6.Text = "*.AX"
End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = "E&xit" Then
        End
    Else                    ' If user chose Cancel, just end Search.
        SearchFlag = False
    End If
End Sub

Private Sub cmdSearch_Click()

Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
lstFoundFiles.Clear
    If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
    NumFiles = 0                       ' Reset found files indicator.
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path
    cmdSearch.SetFocus

End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = Dir1.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = Dir1.Path                      ' Save old path for next recursion.
        Dir1.Path = NewPath
        If Dir1.ListCount > 0 Then
            ' Get to the node bottom.
            Dir1.Path = Dir1.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((Dir1.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If File1.ListCount Then
        If Len(Dir1.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = Dir1.Path                  ' If at root level, leave as is...
        Else
            ThePath = Dir1.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To File1.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + File1.List(ind)
            lstFoundFiles.AddItem entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        Dir1.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function


Private Sub dir1_LostFocus()
    Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub ResetSearch()
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
    Dir1.Path = CurDir: Drive1.Drive = Dir1.Path ' Reset the path.
End Sub

Private Sub txtSearchSpec_Change()
    ' Update file list box if user changes pattern.
    File1.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub


Private Sub Text7_Change()
Text6 = "*." & Text7
File1.Pattern = Text6.Text
End Sub

