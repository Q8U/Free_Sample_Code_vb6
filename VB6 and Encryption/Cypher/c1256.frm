VERSION 5.00
Begin VB.Form c1256 
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox output1 
      Height          =   1425
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.TextBox input1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Input text"
      Top             =   2280
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crack it"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Shift --------------->"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "c1256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
cypher input1

End Sub

Function cypher(text As TextBox)
For f = 1 To 255
DoEvents
strng$ = text.text
  For c = 1 To Len(strng$)
  DoEvents
  g = Mid(strng$, c, 1)
    If Val(Asc(g)) + 1 > 255 Then
    fin = Chr(1)
    Else:
    fin = Chr(Val(Asc(g) + 1))
    End If
    ex = ex & fin
  Next c
  output1.AddItem ex, (a)
  input1.text = ex
  ex = ""
Next f
End Function

Private Sub Command2_Click()
output1.Clear
End Sub

Private Sub Form_Load()
Me.Show
End Sub

Private Sub output1_Click()
d = 0
For x = 1 To output1.ListCount
If output1.List(x) = output1.text Then Label1.Caption = d: d = 0: Exit Sub
d = Val(d) + 1
Next x
Label1.Caption = d
End Sub
