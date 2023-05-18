VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5985
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkKeepData 
      Caption         =   "Keep edge data (for Left, Right, Up and Down Icon)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.HScrollBar hsbShadowSpace 
      Height          =   255
      Left            =   2160
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   240
      Value           =   1
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblShaodwSpaceNumber 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblShadowSpace 
      Caption         =   "Shadow Space:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intButtonClicked As Integer

Private Sub CancelButton_Click()
  intButtonClicked = vbCancel
  Hide
End Sub

Private Sub Form_Load()
  lblShaodwSpaceNumber.Caption = frmMain.intShadowSpace
  hsbShadowSpace.Value = frmMain.intShadowSpace
End Sub

Private Sub hsbShadowSpace_Change()
  lblShaodwSpaceNumber.Caption = hsbShadowSpace.Value
End Sub

Private Sub OKButton_Click()
  intButtonClicked = vbOK
  Hide
End Sub
