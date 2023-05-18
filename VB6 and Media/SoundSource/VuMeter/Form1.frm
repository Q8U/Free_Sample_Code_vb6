VERSION 5.00
Object = "{8BC69DBD-5B39-11D4-9012-91733EB02076}#1.0#0"; "Dancer.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BrainFusion Sound Reader"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   5250
      Picture         =   "Form1.frx":044A
      ScaleHeight     =   705
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   480
      Width           =   1185
      Begin VB.Shape Shape10 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   100
         Left            =   30
         Shape           =   3  'Circle
         Top             =   450
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   100
         Left            =   990
         Shape           =   3  'Circle
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   1000
         X2              =   540
         Y1              =   180
         Y2              =   570
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5730
      Top             =   3870
   End
   Begin Dancer_Ocx.Grapher Grapher2 
      Height          =   1455
      Left            =   90
      Top             =   1500
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2566
      BackColor       =   0
      ForeColor       =   255
      Max             =   100
      BarWidth        =   1
      Flat            =   -1  'True
      Inverted        =   -1  'True
      Bstyle          =   2
      Timer           =   50
   End
   Begin Dancer_Ocx.Grapher Grapher1 
      Height          =   1455
      Left            =   90
      Top             =   90
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2566
      BackColor       =   0
      ForeColor       =   255
      Max             =   100
      BarWidth        =   1
      Flat            =   -1  'True
      Inverted        =   0   'False
      Bstyle          =   2
      Timer           =   50
   End
   Begin Dancer_Ocx.Dancer Dancer1 
      Height          =   1455
      Left            =   4650
      TabIndex        =   0
      Top             =   90
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   2566
      Max             =   100
      Color           =   0
      Timer           =   50
   End
   Begin Dancer_Ocx.Dancer Dancer3 
      Height          =   1515
      Left            =   4650
      TabIndex        =   1
      Top             =   1500
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   2672
      Max             =   100
      Orientation     =   2
      Color           =   0
      Timer           =   50
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   225
      Left            =   5130
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   225
      Left            =   5130
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   225
      Left            =   5130
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   225
      Left            =   5130
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   225
      Left            =   5130
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   705
      Left            =   5250
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Signal Level"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   2985
      Left            =   5010
      Top             =   60
      Width           =   1635
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   2985
      Left            =   60
      Top             =   60
      Width           =   4875
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Very Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5130
      TabIndex        =   4
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5130
      TabIndex        =   5
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5130
      TabIndex        =   6
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5130
      TabIndex        =   7
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5130
      TabIndex        =   8
      Top             =   1440
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vol As Long
Dim Max As Long

Private Sub Dancer1_OnTimer(VolumeValue As Long, VolumeMax As Long)
Vol = VolumeValue
Max = VolumeMax
End Sub

Private Sub Dancer3_OnTimer(VolumeValue As Long, VolumeMax As Long)
Dim X As Integer
X = (VolumeValue / VolumeMax) * 100

Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack
Label5.ForeColor = vbBlack
Label6.ForeColor = vbBlack
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False

If X <> 0 Then
  Label6.Caption = "Very Low"
Else
  Label6.Caption = "Zero"
End If
If (X > 80 And X < 101) Then
  Label2.ForeColor = vbRed
  Shape3.Visible = True
  Exit Sub
End If
If X > 60 And X < 81 Then
  Label3.ForeColor = vbRed
  Shape4.Visible = True
  Exit Sub
End If
If X > 40 And X < 61 Then
  Shape5.Visible = True
  Label4.ForeColor = vbRed
  Exit Sub
End If
If X > 20 And X < 41 Then
  Shape6.Visible = True
  Label5.ForeColor = vbRed
  Exit Sub
End If
If X > -1 And X < 21 Then
  Shape7.Visible = True
  Label6.ForeColor = vbRed
  Exit Sub
End If

End Sub

Private Sub Form_Load()
Dancer1.Dance = True
Dancer3.Dance = True
Grapher1.Dance = True

End Sub

Private Sub Grapher1_OnTimer(VolumeValue As Long, VolumeMax As Long)
Grapher1.ForeColor = RGB(75 + (VolumeValue / VolumeMax) * 180, 0, 0)
Grapher2.Update (VolumeValue / VolumeMax) * Grapher2.Max
Grapher2.ForeColor = RGB(0, 0, 75 + (VolumeValue / VolumeMax) * 180)
End Sub

Private Sub Timer1_Timer()
Dim t As Long

t = (Vol / (Max + 1)) * 900
Line1.X1 = 100 + t
t = (Vol / (Max + 1)) * 150
Line1.BorderColor = RGB(100 + t, 0, 0)

If Vol = Max Then
  Shape9.Visible = True
  Shape10.Visible = False
ElseIf Vol = 0 Then
  Shape10.Visible = True
  Shape9.Visible = False
Else
  Shape9.Visible = False
  Shape10.Visible = False
End If
  
End Sub
