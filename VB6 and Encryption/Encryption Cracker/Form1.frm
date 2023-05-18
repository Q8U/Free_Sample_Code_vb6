VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Basic Character Replacement Cracker"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "basic encoder/decoder"
      Height          =   345
      Left            =   3900
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Step 3:"
      Height          =   3135
      Left            =   0
      TabIndex        =   6
      Top             =   3750
      Width           =   6495
      Begin VB.TextBox Text3 
         Height          =   2205
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "Form1.frx":0000
         Top             =   840
         Width           =   6255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Decrypt using method #2"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Decrypt using method #1"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step 2:"
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   6495
      Begin VB.TextBox Text2 
         Height          =   2205
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form1.frx":0002
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter Encrypted Text:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 1:"
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3705
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Encrypted Spacekey:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1950
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this program "cracks" character replacement encoders
'this is similar to quig's decoder availiable at:http://www.planet-source-code.com/xq/ASP/txtCodeId.11763/lngWId.1/qx/vb/scripts/ShowCode.htm
'the problem with his is that i cant figure out how it works
'this is much more basic im guessing
'how it works:
'step 1 : takes the encrypted space character(its easy to find its the most common character)
'step 2 : subtract/add 32 to its ascii value(the spacebars ascii value)
'step 3 : decode the string using the newfound "algorythm"
'please use this code, but don't pass it off as your own
'themba
'ps plz vote 4 me! thanx
Private Sub Command1_Click()
On Error Resume Next
Text3 = "" 'hhhm
alg = Asc(Text1) - 32 'method #1 subtract the ascii value of space from the encrypted value
For i = 1 To Len(Text2) 'this statement gets the next charater
nxt = Left(Text2, i)
Text3 = Text3 + Chr(Asc(Right(nxt, 1)) - alg) 'decode using newfound "algorythm"
Next i
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text3 = ""
alg = Asc(Text1) + 32 'method #2 add the ascii value of space from the encrypted value
For i = 1 To Len(Text2)
nxt = Left(Text2, i)
Text3 = Text3 + Chr(Asc(Right(nxt, 1)) - alg)
Next i
End Sub

Private Sub Command3_Click()
Form2.Visible = True
End Sub

