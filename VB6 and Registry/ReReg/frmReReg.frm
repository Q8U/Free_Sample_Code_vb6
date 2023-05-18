VERSION 5.00
Begin VB.Form frmReReg 
   Caption         =   "ReReg"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "frmReReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin prjReReg.ReReg ReReg 
      Left            =   120
      Top             =   360
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton cmdTerminatePID 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Terminate a given ProcessID and al it's handles"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtPID 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowWinDir 
      Caption         =   "Show Windows Dir"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Display the Windows installation folder"
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdReReg 
      Caption         =   "Reload Registry"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Reload the Windows Registry"
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "PID"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "frmReReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReReg_Click()

    ReReg.ReloadRegistry

End Sub

Private Sub cmdShowWinDir_Click()

    MsgBox ReReg.WinDir

End Sub

Private Sub cmdTerminatePID_Click()

    If Me.txtPID Is Not Null Then
        ReReg.TerminatePID (Me.txtPID)
    Else
        MsgBox "No ProcessID to Terminate", vbCritical, "Error !"
    End If

End Sub
