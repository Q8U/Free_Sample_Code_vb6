VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Colour Scheme"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Top             =   1680
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command Button"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Collection Frame"
      Height          =   1575
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Text Box"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Text Label"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Tmp As String
Tmp = Dir(App.Path & "\Schemes\")

Do Until Tmp = ""
    If LCase(Right(Tmp, 4)) = ".skn" Then
        List1.AddItem Tmp
    End If
    DoEvents
    Tmp = Dir()
Loop
End Sub

Private Sub Label1_Click()

End Sub
