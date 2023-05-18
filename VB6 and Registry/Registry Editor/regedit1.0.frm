VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RegEdit v1.0"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5580
   Icon            =   "regedit1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      Caption         =   "Note: You Must Type In The Location Of The Background File And You Must Restart Before The Text Will Appear In Your System Tray!"
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000012&
      Caption         =   "File Location"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "Set Image"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "Change IE Background"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   5760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "Remove Text"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "Insert Text"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   5760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "Insert Text In Your System Tray"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Change IE Window Title"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Default Title"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Programmed By: Vegeta"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Change Window Title"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
On Error Resume Next
dirlist.Path = drvlist.Drive
ChDir dirlist.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo drivehandler
dirlist.Path = drvlist.Drive
drivehandler:
drvlist.Drive = dirlist.Path
End Sub

Private Sub File1_Click()
On Error Resume Next
Text3.Text = filList.Path
Text3 = Text3 + "\" & filList.FileName
End Sub

Private Sub Command1_Click()
 Command1.DialogTitle = "Open Character"
    Command1.Filter = "Diablo 2 Saved Games (*.d2s)|*.d2s|"
End Sub

Private Sub Label1_Click()
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Window Title", Text1.Text
current_tab = 1
End Sub

Private Sub Label10_Click()
RegDeleteKey "HKEY_CURRENT_USER\Control Panel\International", "s1159"
RegDeleteKey "HKEY_CURRENT_USER\Control Panel\International", "s2359"
End Sub

Private Sub Label13_Click()
On Error Resume Next
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", Text3.Text
Text3.Text = ""
End Sub

Private Sub Label14_Click()
On Error Resume Next
Text3.Text = "(Default)"
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", Text3.Text
Text3.Text = ""
MsgBox "Default Background Applied", vbInformation, "Info"
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label2_Click()
frmAbout.Show vbModal
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Window Title", "Internet Explorer"
End Sub

Private Sub Label9_Click()
SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "s1159", Text2.Text
SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "s2359", Text2.Text
End Sub

