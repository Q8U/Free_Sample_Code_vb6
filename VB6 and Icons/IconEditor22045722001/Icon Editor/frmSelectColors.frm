VERSION 5.00
Begin VB.Form frmSelectColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Colors"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   22
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   21
      Top             =   2160
      Width           =   615
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   20
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   4080
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   17
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3240
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3240
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3240
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3240
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2400
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2400
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2400
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2400
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      Caption         =   "Color2"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      Caption         =   "Color1"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblSelectColors 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmSelectColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngColor1, lngColor2 As Long
Public intButtonClicked As Integer
Option Explicit

Private Sub CancelButton_Click()
  intButtonClicked = vbCancel
  Hide
End Sub

Private Sub Form_Load()
  Dim intI As Integer
  
  lngColor1 = vbBlack
  lngColor2 = vbBlack
  
  For intI = 0 To 15
    picColors(intI).BackColor = QBColor(15 - intI)
  Next intI
  
  optColor(1).Value = True
  
  picColor(1).BackColor = lngColor1
  picColor(2).BackColor = lngColor1
  
  intButtonClicked = vbNo
End Sub


Private Sub OKButton_Click()
  intButtonClicked = vbOK
  Hide
End Sub

Private Sub optColor_Click(Index As Integer)
  optColor(Index).Value = True
End Sub

Private Sub picColors_Click(Index As Integer)
  If (optColor(1).Value = True) Then
    lngColor1 = picColors(Index).BackColor
    picColor(1).BackColor = lngColor1
  Else
    lngColor2 = picColors(Index).BackColor
    picColor(2).BackColor = lngColor2
  End If
End Sub
