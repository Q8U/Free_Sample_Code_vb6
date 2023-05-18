VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5988
   LinkTopic       =   "Form1"
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Com 
      Left            =   5040
      Top             =   1080
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Index           =   1
      Left            =   5040
      ScaleHeight     =   252
      ScaleWidth      =   732
      TabIndex        =   2
      Top             =   2160
      Width           =   732
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   192
      Index           =   0
      Left            =   3960
      Picture         =   "Source.frx":0000
      ScaleHeight     =   192
      ScaleWidth      =   192
      TabIndex        =   1
      Top             =   2280
      Width           =   192
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AddButton"
      Height          =   252
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fre As Integer
Dim prev As Integer


Private Sub Command1_Click()
Com.ShowOpen
If Com.filename = "" Then Exit Sub
fre = fre + 1
If fre > 1 Then Load Picture1(fre)
Picture1(fre).Visible = True
Picture1(fre).Picture = LoadPicture(Com.filename)
Picture1(fre).Top = Picture1(fre - 1).Top
Picture1(fre).Left = Picture1(fre - 1).Left + 35
End Sub

Private Sub Form_Load()
fre = 0
Call spixelset(10, 10, 480, 10, Form1, 0)
Call spixelset(10, 40, 480, 40, Form1, 0)
Call spixelset(10, 11, 10, 40, Form1, 1)
Call spixelset(480, 10, 480, 41, Form1, 1)
Call DrawLine(15, 12, 15, 38, Form1, 0)
Call DrawLine(18, 12, 18, 38, Form1, 0)
Picture1(fre).BorderStyle = 0
Picture1(fre).Left = 32
Picture1(fre).Top = 17
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
End Sub

