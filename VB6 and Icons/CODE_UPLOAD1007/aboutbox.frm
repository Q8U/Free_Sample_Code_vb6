VERSION 5.00
Begin VB.Form AboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About IconWorks"
   ClientHeight    =   2370
   ClientLeft      =   1365
   ClientTop       =   1425
   ClientWidth     =   5640
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABOUTBOX.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   5640
   Begin VB.PictureBox Pic_ApplicationIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   255
      Picture         =   "ABOUTBOX.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   255
      Width           =   480
   End
   Begin VB.CommandButton Cmd_OK 
      Caption         =   "OK"
      Height          =   360
      Left            =   4350
      TabIndex        =   6
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Line lin_HorizontalLine1 
      BorderWidth     =   2
      X1              =   975
      X2              =   5010
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Lbl_IconWorks 
      Caption         =   "IconWorks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   990
      TabIndex        =   1
      Top             =   270
      Width           =   1920
   End
   Begin VB.Label Lbl_Version 
      Caption         =   "Version 1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   2
      Top             =   840
      Width           =   1470
   End
   Begin VB.Label Lbl_Microsoft 
      Caption         =   "Microsoft Visual Basic for Windows 3.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   3
      Top             =   1110
      Width           =   4365
   End
   Begin VB.Label Lbl_Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1005
      TabIndex        =   5
      Top             =   1545
      Width           =   1875
   End
   Begin VB.Label Lbl_InfoValues 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2910
      TabIndex        =   0
      Top             =   1545
      Width           =   1410
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Win32 Then
#Else
Private Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Private Declare Function GetWinFlags Lib "Kernel" () As Long
Const WF_STANDARD = &H10
Const WF_ENHANCED = &H20
Const WF_80x87 = &H400
#End If

Private Sub Cmd_OK_Click()
    Unload AboutBox
End Sub

Private Sub Form_Load()
Dim WinFlags As Long
Dim Mode As String, Processor As String
    ' Dialog Boxes should only have Move and Close items
    ' in their System menus', so remove the others.
    Remove_Items_From_Sysmenu AboutBox
    ' Center the AboutBox on the screen
     Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    ' Get current Windows configuration
#If Win32 Then
#Else
    WinFlags = GetWinFlags()
    If WinFlags And WF_ENHANCED Then Mode = "386 Enhanced Mode" Else Mode = "Standard Mode"
    Lbl_Info.Caption = Mode + CRLF + "Free Memory:" + CRLF + "Math Co-processor:"
    If WinFlags And WF_80x87 Then Processor = "Present" Else Processor = "None"
    Lbl_InfoValues.Caption = CRLF + Format$(GetFreeSpace(0) \ 1024) + " KB" + CRLF + Processor
#End If
End Sub

