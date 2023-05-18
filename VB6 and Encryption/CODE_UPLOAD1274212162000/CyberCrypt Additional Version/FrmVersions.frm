VERSION 5.00
Begin VB.Form FrmVersions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product License"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "FrmVersions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKCmd 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FrmVersions.frx":030A
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmVersions.frx":0A70
      Top             =   2400
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   2070
      Left            =   120
      Top             =   120
      Width           =   7230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed to:"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label UserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is being run on:"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   2040
   End
   Begin VB.Label OsLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   3480
      Width           =   45
   End
End
Attribute VB_Name = "FrmVersions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    'The following two lines loads the user and windows OS into the program versions
    UserName.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "RegisteredOwner") & " - " & GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "VersionNumber")
    OsLabel.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "ProductName") & " - " & GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "ProductId")
End Sub

Private Sub OKCmd_Click()
    Unload Me
End Sub
