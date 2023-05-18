VERSION 5.00
Begin VB.Form VBSYXMID 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "VB SYSEX Output Devices"
   ClientHeight    =   1995
   ClientLeft      =   3075
   ClientTop       =   3030
   ClientWidth     =   3390
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1995
   ScaleWidth      =   3390
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton OK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "VBSYXMID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    ' Text.Text = "N"
    Hide
End Sub

Private Sub List1_Click()
    ' form1.Print list1.ListIndex
End Sub

Private Sub OK_Click()
    Text.Text = Str$(List1.ListIndex)
    If Text.Text <> -1 Then
        Cancel.Enabled = True
        Hide
    End If
    'Form1.Print List1.ListIndex - 1
End Sub

