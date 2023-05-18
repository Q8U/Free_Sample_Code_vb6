VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Parse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parsing Files"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3780
      Top             =   945
   End
   Begin ACTIVESKINLibCtl.SkinLabel FileCount 
      Height          =   255
      Left            =   390
      OleObjectBlob   =   "Working.frx":0000
      TabIndex        =   3
      Top             =   1980
      Width           =   2130
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5295
      OleObjectBlob   =   "Working.frx":0086
      TabIndex        =   2
      Top             =   1980
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel Label1 
      Height          =   1635
      Left            =   390
      OleObjectBlob   =   "Working.frx":011A
      TabIndex        =   1
      Top             =   150
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "Parse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmTagger.CancelIt = True
End Sub

Private Sub Form_Load()
    frmTagger.Skin1.ApplySkin Me.hwnd 'Skin me
    Randomize 'Apply Random
    GetQuote 'Change Quote
End Sub

Private Sub GetQuote()
    'Show the Random Quote
    I = Int(Rnd * frmTagger.List3.ListCount)
    Label1.Caption = frmTagger.List3.List(I)
End Sub

Private Sub Timer1_Timer()
    GetQuote 'Change Quote
End Sub
