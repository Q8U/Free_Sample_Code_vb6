VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Basic Encoder/Decoder"
   ClientHeight    =   465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form2"
   ScaleHeight     =   465
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2610
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decode"
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   5070
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cyph = 1
Private Sub Command1_Click()
For i = 1 To Len(Text1)
nxt = Left(Text1, i)
Text2 = Text2 + Chr(Asc(Right(nxt, 1)) + cyph)
Next i
End Sub

Private Sub Command2_Click()
For i = 1 To Len(Text1)
nxt = Left(Text2, i)
Text3 = Text3 + Chr(Asc(Right(nxt, 1)) - cyph)
Next i
End Sub

