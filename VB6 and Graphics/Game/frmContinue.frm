VERSION 4.00
Begin VB.Form frmContinue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   1575
   ClientTop       =   405
   ClientWidth     =   5985
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Height          =   6615
   Left            =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Top             =   60
   Width           =   6105
   Begin VB.Label lblChoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   1
      Left            =   2430
      TabIndex        =   3
      Top             =   3915
      Width           =   1275
   End
   Begin VB.Label lblChoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   1935
      TabIndex        =   2
      Top             =   3105
      Width           =   2115
   End
   Begin VB.Label lblGameName 
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3330
      TabIndex        =   1
      Top             =   1125
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Saved As:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1215
      TabIndex        =   0
      Top             =   1170
      Width           =   2115
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   135
      Picture         =   "frmContinue.frx":0000
      Stretch         =   -1  'True
      Top             =   5985
      Width           =   5640
   End
   Begin VB.Image Image4 
      Height          =   180
      Left            =   180
      Picture         =   "frmContinue.frx":42D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5550
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   5715
      Picture         =   "frmContinue.frx":85A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   6180
      Left            =   0
      Picture         =   "frmContinue.frx":BEC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmContinue"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Activate()
lblGameName.Caption = GameName

'draw 1000 stars on the screen
    Dim a, x, Y
    For a = 1 To 300
        x = Int(Rnd * Me.ScaleWidth)
        Y = Int(Rnd * Me.ScaleHeight)
        Me.PSet (x, Y), vbWhite
    Next a

    'draw darker stars
    Dim grey
    grey = &H808080
    For a = 1 To 300
       x = Int(Rnd * Me.ScaleWidth)
       Y = Int(Rnd * Me.ScaleHeight)
       Me.PSet (x, Y), grey
    Next a
       
    'draw some blue stars
    Dim blue
    blue = &H800000
    For a = 1 To 300
       x = Int(Rnd * Me.ScaleWidth)
       Y = Int(Rnd * Me.ScaleHeight)
       Me.PSet (x, Y), blue
    Next a
End Sub

Private Sub Form_Load()
'Me.Top = frmCompress.Top
'Me.Left = frmCompress.Left

End Sub


Private Sub Label2_Click(Index As Integer)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim i
  For i = 0 To 1
    lblChoice(i).ForeColor = vbRed
  Next i
End Sub

Private Sub lblChoice_Click(Index As Integer)
Select Case Index
Case 0
    'back to main menu
    Load frmCover
    frmCover.StarsDrawn = False 'make sure stars are drawn again
    Unload frmGameScreen
    Me.Hide
    frmCover.Show
    Unload Me
    
Case 1
    'quit the game
    PlaySoundEffect "Abort"
    'deregister help file
    QuitHelp
    End
    
End Select

End Sub

Private Sub lblChoice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim Counter
  For Counter = 0 To 1
    lblChoice(Counter).ForeColor = vbRed
  Next Counter
  
  lblChoice(Index).ForeColor = vbBlue
End Sub


