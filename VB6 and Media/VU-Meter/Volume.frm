VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VU-Meter"
   ClientHeight    =   5370
   ClientLeft      =   5025
   ClientTop       =   345
   ClientWidth     =   4920
   FillStyle       =   0  'Solid
   Icon            =   "Volume.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   4920
   Begin VB.PictureBox Picture9 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3420
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   25
      Top             =   4890
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3060
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   24
      Top             =   4890
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   3390
      Picture         =   "Volume.frx":030A
      ScaleHeight     =   4665
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   180
      Width           =   225
      Begin VB.PictureBox Picture7 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   30
         ScaleHeight     =   4665
         ScaleWidth      =   165
         TabIndex        =   23
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   3030
      Picture         =   "Volume.frx":9A6E
      ScaleHeight     =   4665
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   180
      Width           =   225
      Begin VB.PictureBox Picture5 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   30
         ScaleHeight     =   4665
         ScaleWidth      =   165
         TabIndex        =   22
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1620
      Top             =   5460
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   660
      Top             =   5460
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1140
      Top             =   5460
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   180
      Top             =   5460
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "VU-Meter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4785
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2715
      Begin VB.Frame Frame3 
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   120
         TabIndex        =   12
         Top             =   1260
         Width           =   915
         Begin VB.CommandButton Command1 
            Caption         =   "Chiudi"
            Default         =   -1  'True
            Height          =   360
            Left            =   120
            TabIndex        =   17
            Top             =   3000
            Width           =   705
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   315
            Index           =   2
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   14
            Text            =   "Volume.frx":131D2
            Top             =   2520
            Width           =   615
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   1305
            Index           =   2
            Left            =   180
            Picture         =   "Volume.frx":131D4
            ScaleHeight     =   1305
            ScaleWidth      =   540
            TabIndex        =   13
            Top             =   840
            Width           =   540
            Begin VB.Image Image3 
               Height          =   285
               Index           =   2
               Left            =   120
               Picture         =   "Volume.frx":156CA
               Stretch         =   -1  'True
               Top             =   780
               Width           =   315
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "L+R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Massimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   16
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Top             =   2220
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Volume rec."
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
         Height          =   3435
         Left            =   1140
         TabIndex        =   3
         Top             =   1260
         Width           =   1455
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   315
            Index           =   1
            Left            =   750
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "Volume.frx":159CC
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   315
            Index           =   0
            Left            =   100
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "Volume.frx":159CE
            Top             =   2520
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "L = R"
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
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   3120
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   1305
            Index           =   0
            Left            =   180
            Picture         =   "Volume.frx":159D0
            ScaleHeight     =   1305
            ScaleWidth      =   540
            TabIndex        =   5
            Top             =   840
            Width           =   540
            Begin VB.Image Image3 
               Height          =   285
               Index           =   0
               Left            =   120
               Picture         =   "Volume.frx":17EC6
               Stretch         =   -1  'True
               Top             =   780
               Width           =   315
            End
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   1305
            Index           =   1
            Left            =   720
            Picture         =   "Volume.frx":181C8
            ScaleHeight     =   1305
            ScaleWidth      =   540
            TabIndex        =   4
            Top             =   840
            Width           =   540
            Begin VB.Image Image3 
               Height          =   285
               Index           =   1
               Left            =   120
               Picture         =   "Volume.frx":1A6BE
               Stretch         =   -1  'True
               Top             =   780
               Width           =   315
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "L - Volume - R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Massimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   7
            Top             =   2220
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   915
         Left            =   1380
         Picture         =   "Volume.frx":1A9C0
         ScaleHeight     =   855
         ScaleWidth      =   1185
         TabIndex        =   2
         Top             =   240
         Width           =   1250
         Begin VB.Image Image2 
            Height          =   150
            Left            =   510
            Picture         =   "Volume.frx":1B32A
            Stretch         =   -1  'True
            Top             =   240
            Width           =   150
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   540
            X2              =   180
            Y1              =   600
            Y2              =   180
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   915
         Left            =   120
         Picture         =   "Volume.frx":1B3FC
         ScaleHeight     =   855
         ScaleWidth      =   1185
         TabIndex        =   1
         Top             =   240
         Width           =   1250
         Begin VB.Image Image1 
            Height          =   150
            Left            =   510
            Picture         =   "Volume.frx":1BD66
            Stretch         =   -1  'True
            Top             =   240
            Width           =   150
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   540
            X2              =   180
            Y1              =   600
            Y2              =   180
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   "By Mauro Di Leonardo mdileo@silvi.net"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Volume e vu
Dim vol2 As Long
Dim vol1 As Long
Dim hmixer As Long ' mixer handle
Dim volCtrl4 As MIXERCONTROL ' wavein mute control
Dim volCtrl3 As MIXERCONTROL ' waveout volume control
Dim volCtrl2 As MIXERCONTROL ' vu control play
Dim volCtrl1 As MIXERCONTROL ' vu control rec
Dim volCtrl As MIXERCONTROL  ' wavein volume control
Dim rc As Long               ' return code
Dim ok As Boolean            ' boolean return code
Dim volMin As Long
Dim volMax As Long
Dim Llevelmeter As Long
Dim Rlevelmeter As Long

'Cursore
Dim Mouse_Button As Integer
Dim Mouse_Y As Single
Dim cursore As Integer
Public Sub sposta_Rlevel(aaa As Long)

Line2.X1 = 540 + (aaa / 32768) * 130
Picture5.Top = Picture5.Height - 4665 + (aaa / 32768) * 2350
Line2.X2 = 180 + (aaa / 32768) * 900
Picture7.Top = Picture7.Height - 4665 + (aaa / 32768) * 2350
If aaa < 32768 / 2 Then Line2.Y2 = 200 - (aaa / 32768) * 200
If aaa > 32768 / 2 Then Line2.Y2 = 50 + (aaa / 32768) * 150

If aaa > 31000 Then
Image2.Visible = True
Picture8.Visible = True
Picture9.Visible = True
Timer4.Enabled = True
End If

End Sub


Public Sub sposta_Llevel(aaa As Long)

Line1.X1 = 540 + (aaa / 32768) * 130
Line1.X2 = 180 + (aaa / 32768) * 900
If aaa < 32768 / 2 Then Line1.Y2 = 200 - (aaa / 32768) * 200
If aaa > 32768 / 2 Then Line1.Y2 = 50 + (aaa / 32768) * 150

If aaa > 31000 Then
Image1.Visible = True
Timer4.Enabled = True
End If

End Sub
Private Sub Command1_Click()
    
    
   
    Unload Me
    
End Sub
Private Sub Form_Load()

Form1.Move Screen.Width - Form1.Width, S0

Dim a As Integer

   InitVolumein
   
   For a = 0 To 1
   Image3(a).Top = Picture4(a).Height - Image3(a).Height - (Text1(a).Text / volMax * (Picture4(a).Height - Image3(a).Height))
   Next a
   
   initvolumeout
   Image3(2).Top = Picture4(2).Height - Image3(2).Height - (Text1(2).Text / volMax * (Picture4(2).Height - Image3(2).Height))
   
End Sub

Private Sub InitVolumein()

Dim a As Long

    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Non posso aprire il mixer."
        Exit Sub
    End If
   
    ok = GetInitVolume(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINE, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
    
    If (ok = True) Then
        volMin = volCtrl.lMinimum
        volMax = volCtrl.lMaximum
        a = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINE, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
        Text1(0).Text = Lvol
        Text1(1).Text = Rvol
    End If
    
'setta tutti i seleziona di registrazione del mixer su on se questi sono su off
    
    ok = GetWaveInMute(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT, volCtrl4)
    
    If (ok = True) Then
        a = setmuteon(hmixer, volCtrl4)
    End If
    
End Sub
Private Sub SettaVolumein()

    volL = CLng(Text1(0).Text)
    volR = CLng(Text1(1).Text)
    SetVolumeControl hmixer, volCtrl, volL, volR
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim a As Long
a = mixerClose(hmixer)

End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 cursore = Index
 Timer3.Enabled = True
 Timer2.Enabled = False
 Mouse_Button = Button
End Sub

Private Sub Image3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Mouse_Button = Button
   Mouse_Y = Image3(Index).Top + Y
   
   If Button = 1 Then
   If Index = 2 Then
   settavolumeout
   Else
   SettaVolumein
   End If
   End If
   
End Sub


Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Timer3.Enabled = False
  Mouse_Button = Button
  Timer2.Enabled = True
End Sub

Private Sub timer1_Timer()

Dim a As Long
Dim b As Long
Dim d As Long
Dim e As Long
Dim f As Long

a = Abs(GetVuControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, volCtrl2))
a = Abs(getrecvucontrol(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, MIXERCONTROL_CONTROLTYPE_PEAKMETER, volCtrl1))

b = Abs(Lvu)
d = Abs(Rvu)
e = Abs(Lrecvu)
f = Abs(Rrecvu)

If b > Llevelmeter Then
Llevelmeter = b
sposta_Llevel (Llevelmeter)
End If

If e > Llevelmeter Then
Llevelmeter = e
sposta_Llevel (Llevelmeter)
End If

If d > Rlevelmeter Then
Rlevelmeter = d
sposta_Rlevel (Rlevelmeter)
End If

If f > Rlevelmeter Then
Rlevelmeter = f
sposta_Rlevel (Rlevelmeter)
End If

If Rlevelmeter > 0 Then Rlevelmeter = Rlevelmeter - 1000
If Llevelmeter > 0 Then Llevelmeter = Llevelmeter - 1000
If Rlevelmeter < 0 Then Rlevelmeter = 0
If Llevelmeter < 0 Then Llevelmeter = 0
sposta_Rlevel (Rlevelmeter)
sposta_Llevel (Llevelmeter)

End Sub

Private Sub Timer2_Timer()

Dim a As Long
Dim b As Long
b = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINE, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
Text1(1).Text = Lvol
Text1(0).Text = Rvol
For a = 0 To 1
Image3(a).Top = Picture4(a).Height - Image3(a).Height - (Text1(a).Text / volMax * (Picture4(a).Height - Image3(a).Height))
Next a
b = getoutvolumecontrol(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl3)
Text1(2).Text = volumeout
Image3(2).Top = Picture4(2).Height - Image3(2).Height - (Text1(2).Text / volMax * (Picture4(2).Height - Image3(2).Height))

End Sub

Private Sub Timer3_Timer()
Dim TempTop As Integer
   If Mouse_Button = 1 Then
      TempTop = Mouse_Y - Image3(cursore).Height / 2
      If TempTop + Image3(cursore).Height > Picture4(cursore).Height Then
         TempTop = Picture4(cursore).Height - Image3(cursore).Height
      End If
      If TempTop < 0 Then TempTop = 0
   If Image3(cursore).Top <> TempTop Then Image3(cursore).Top = TempTop
 
   Text1(cursore).Text = Fix(Abs(CDbl(volMax * (Image3(cursore).Top / (Picture4(cursore).Height - Image3(cursore).Height))) - volMax))
   End If
If Check1.Value = 1 Then
If cursore = 0 Then Image3(1).Top = Image3(0).Top
If cursore = 0 Then Text1(1).Text = Text1(0).Text
If cursore = 1 Then Image3(0).Top = Image3(1).Top
If cursore = 1 Then Text1(0).Text = Text1(1).Text
End If

End Sub


Private Sub Timer4_Timer()
Timer4.Enabled = False
Image1.Visible = False
Image2.Visible = False
Picture8.Visible = False
Picture9.Visible = False
End Sub



Public Sub settavolumeout()

        vol = CLng(Text1(2).Text)
        setoutvolumecontrol hmixer, volCtrl3, vol
    
End Sub

Private Sub initvolumeout()

Dim a As Long
  
    ok = getinitvolumeout(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl3)
    
    If (ok = True) Then
        volMin = volCtrl.lMinimum
        volMax = volCtrl.lMaximum
        a = getoutvolumecontrol(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl3)
        Text1(2).Text = volumeout
        
    End If
    
End Sub
