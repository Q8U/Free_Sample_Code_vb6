VERSION 4.00
Begin VB.Form AboutBox1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "AboutBox"
   ClientHeight    =   4890
   ClientLeft      =   1395
   ClientTop       =   1725
   ClientWidth     =   4965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Height          =   5340
   Icon            =   "aboutbx1.frx":0000
   Left            =   1335
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4965
   Top             =   1335
   Width           =   5085
   Begin Threed.SSPanel SSPanel1 
      Height          =   4885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4970
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   8625
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BevelWidth      =   3
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4320
         Top             =   4260
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   1
         Top             =   4400
         Width           =   1215
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1931
         _StockProps     =   15
         Caption         =   "VB MIDI Piano"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            name            =   "Times New Roman"
            charset         =   0
            weight          =   700
            size            =   18
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Version 1.2"
         ForeColor       =   4210752
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1815
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   3201
         _StockProps     =   15
         Caption         =   "SSPanel2"
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   1635
            Left            =   120
            Picture         =   "aboutbx1.frx":030A
            Top             =   120
            Width           =   4140
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   3480
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "SSPanel5"
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "PO Box 28, Waterford, WI  53185-0028  USA"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   0
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   4
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Copyright © 1993, Arthur V. Edstrom"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   0
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   3
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   360
         Picture         =   "aboutbx1.frx":3F28
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "AboutBox1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    AboutBox1.Hide
End Sub

Private Sub Form_Load()
    ' Center the AboutBox on the screen
    '
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    X% = sndPlaySound(App.Path & "\MIDPIANO.WAV", 3)
    Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
    Screen.MousePointer = 11
    Piano.Show
    Timer1.Enabled = False
    AboutBox1.Command1.Enabled = True
    Screen.MousePointer = 0
End Sub

