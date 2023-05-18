VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   Icon            =   "FrmSplash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":030A
   ScaleHeight     =   3300
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape Shape1 
      Height          =   3300
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'If the program is running already then don't show the splash screen
    If App.PrevInstance = True Then Timer.Enabled = False: Unload Me: frmMain.Show
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    Unload Me
    frmMain.Show
End Sub
