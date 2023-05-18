VERSION 5.00
Begin VB.Form frmMessageBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMessageBox.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctBtnNormal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      Picture         =   "frmMessageBox.frx":030A
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pctBtnDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      Picture         =   "frmMessageBox.frx":15A0
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   2685
      Picture         =   "frmMessageBox.frx":1FFE
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   3720
      Picture         =   "frmMessageBox.frx":3294
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   2
      Left            =   1110
      Picture         =   "frmMessageBox.frx":452A
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   2
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   80
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   690
      TabIndex        =   1
      Top             =   90
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "frmMessageBox.frx":57C0
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   30
      Picture         =   "frmMessageBox.frx":608A
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "frmMessageBox.frx":6954
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "frmMessageBox.frx":721E
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblCaption_Click(Index As Integer)
    pctButton_Click (Index)
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctButton(Index).Picture = pctBtnDown.Picture
    lblCaption(Index).ForeColor = RGB(255, 255, 255)
End Sub

Private Sub lblCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctButton(Index).Picture = pctBtnNormal.Picture
    lblCaption(Index).ForeColor = RGB(0, 0, 0)
End Sub

'Following sub checks and sets the button captions which are returned
'in (Result) as integer
Private Sub pctButton_Click(Index As Integer)
    If lblCaption(Index).Caption = "OK" Then
        Result = 0
    ElseIf lblCaption(Index).Caption = "Yes" Then
        Result = 1
    ElseIf lblCaption(Index).Caption = "No" Then
        Result = 2
    ElseIf lblCaption(Index).Caption = "Cancel" Then
        Result = 3
    ElseIf lblCaption(Index).Caption = "Retry" Then
        Result = 4
    ElseIf lblCaption(Index).Caption = "Ignore" Then
        Result = 5
    ElseIf lblCaption(Index).Caption = "Abort" Then
        Result = 6
    Else
        Result = 7
    End If
    
    Unload Me
End Sub

Private Sub pctButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctButton(Index).Picture = pctBtnDown.Picture
    lblCaption(Index).ForeColor = RGB(255, 255, 255)
End Sub

Private Sub pctButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctButton(Index).Picture = pctBtnNormal.Picture
    lblCaption(Index).ForeColor = RGB(0, 0, 0)
End Sub
