VERSION 5.00
Begin VB.Form errWin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Trapped Errors"
   ClientHeight    =   3234
   ClientLeft      =   44
   ClientTop       =   286
   ClientWidth     =   6842
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3234
   ScaleWidth      =   6842
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tErr 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "errWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmTagger.Skin1.ApplySkin Me.hwnd
    Select Case UnloadMode
        Case 1 'Unloaded with Code
        Case Else 'Unloaded With Button
            Cancel = True
            Me.Hide 'Hide only so Unload can be used later
    End Select
End Sub

Private Sub Form_Resize()
    'Auto Size for Window Sizes
    tErr.Height = Me.ScaleHeight
    tErr.Width = Me.ScaleWidth
End Sub

