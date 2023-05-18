VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regestry Editor Version 1.0"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Open Regestry"
      Height          =   270
      Left            =   2835
      TabIndex        =   25
      Top             =   4350
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Write Binary to regestry"
      Height          =   1860
      Left            =   105
      TabIndex        =   17
      Top             =   2400
      Width           =   4440
      Begin VB.CommandButton Command3 
         Caption         =   "Set DWORD Value"
         Height          =   270
         Left            =   195
         TabIndex        =   21
         Top             =   1455
         Width           =   4020
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   210
         TabIndex        =   20
         Top             =   1080
         Width           =   4005
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2250
         TabIndex        =   19
         Top             =   495
         Width           =   1965
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   495
         Width           =   1965
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DWORD input:"
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   825
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DWORD value name:"
         Height          =   195
         Left            =   2340
         TabIndex        =   23
         Top             =   285
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regestry folder to write to:"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   270
         Width           =   1845
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   270
      Left            =   105
      TabIndex        =   16
      Top             =   4350
      Width           =   2670
   End
   Begin VB.Frame Frame2 
      Caption         =   "Write Binary to regestry"
      Height          =   2265
      Left            =   2340
      TabIndex        =   8
      Top             =   30
      Width           =   2190
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   105
         TabIndex        =   12
         Top             =   495
         Width           =   1965
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   105
         TabIndex        =   11
         Top             =   1005
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   105
         TabIndex        =   10
         Top             =   1515
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Set Binary Value"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   1875
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regestry folder to write to:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Binary value name:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Binary input:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1305
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Write strings to regestry"
      Height          =   2265
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2190
      Begin VB.CommandButton Command1 
         Caption         =   "Set String Value"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   1875
         Width           =   1950
      End
      Begin VB.TextBox StringCall 
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   1515
         Width           =   1965
      End
      Begin VB.TextBox StringName 
         Height          =   285
         Left            =   105
         TabIndex        =   4
         Top             =   1005
         Width           =   1965
      End
      Begin VB.TextBox RegFolder 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   495
         Width           =   1965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "String:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1305
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "String value name:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   795
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regestry folder to write to:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1845
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    SetStringValue RegFolder, StringName, StringCall
    
End Sub

Private Sub Command2_Click()

    End

End Sub

Private Sub Command4_Click()

    SetBinaryValue Text3, Text2, Text1
    
End Sub

Private Sub Command3_Click()

    SetBinaryValue Text4, Text5, Text6
    
End Sub

Private Sub Command5_Click()

    Shell "C:\Windows\Regedit.exe", vbNormalFocus

End Sub

Private Sub Form_Load()

    If App.PrevInstance = True Then End

End Sub
