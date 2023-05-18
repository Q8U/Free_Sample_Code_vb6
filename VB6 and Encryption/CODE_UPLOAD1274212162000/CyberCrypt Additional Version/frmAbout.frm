VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Versions 
      Caption         =   "License"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3120
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   3060
      ScaleWidth      =   4935
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4995
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please click on License for the license agreement."
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   3555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":200BC
         Height          =   585
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   5055
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4950
         Y1              =   1920
         Y2              =   1920
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
    Unload Me
End Sub

Private Sub Versions_Click()
    On Error Resume Next
    FrmVersions.Show 1, Me
End Sub
