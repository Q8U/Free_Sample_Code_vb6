VERSION 5.00
Begin VB.Form Folder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Folder"
   ClientHeight    =   3300
   ClientLeft      =   33
   ClientTop       =   363
   ClientWidth     =   2431
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   2431
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   253
      Left            =   1089
      TabIndex        =   3
      Top             =   3025
      Width           =   616
   End
   Begin VB.CommandButton OK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   253
      Left            =   1815
      TabIndex        =   2
      Top             =   3025
      Width           =   616
   End
   Begin VB.DirListBox Dir1 
      Height          =   2640
      Left            =   0
      TabIndex        =   1
      Top             =   363
      Width           =   2431
   End
   Begin VB.DriveListBox Drive1 
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2431
   End
End
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ChoseFolder As String

Private Sub Drive1_Change()
    'Drive Changes so Change the Path
    Dir1.Path = Mid(Drive1.List(Drive1.ListIndex), 1, InStr(1, Drive1.List(Drive1.ListIndex), ":") - 1)
End Sub

Private Sub Form_Load()
    frmTagger.Skin1.ApplySkin Me.hwnd 'Skin the Form
    'Load Preempt (Imcomplete)
End Sub

Private Sub OK_Click()
    ChoseFolder = Dir1.Path 'Chose Folder used for retrieving before Unloading the Form
    Me.Visible = False 'Hide to Exit Modal
End Sub
