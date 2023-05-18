VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2340
      Width           =   4275
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Key"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set Key"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Key"
      Height          =   375
      Left            =   1860
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   780
      TabIndex        =   4
      Top             =   900
      Width           =   2535
   End
   Begin VB.TextBox txtKey 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Text            =   "My Key"
      Top             =   540
      Width           =   2535
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Text            =   "Software\My Project"
      Top             =   180
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Key"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   330
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
   If DeleteRegKey(HKEY_CURRENT_USER, txtPath.Text) Then
      txtStatus.Text = "Function was successful!"
   Else
      txtStatus.Text = "Function was NOT successful!"
   End If
End Sub

Private Sub cmdGet_Click()
   Dim strValue As String
   If GetRegKey(HKEY_CURRENT_USER, txtPath.Text, txtKey.Text, strValue) Then
      txtValue = strValue
      txtStatus.Text = "Function was successful!"
   Else
      txtStatus.Text = "Function was NOT successful!"
   End If
End Sub

Private Sub cmdSet_Click()
   If SetRegKey(HKEY_CURRENT_USER, txtPath.Text, txtKey.Text, txtValue.Text) Then
      txtStatus.Text = "Function was successful!"
   Else
      txtStatus.Text = "Function was NOT successful!"
   End If
End Sub
