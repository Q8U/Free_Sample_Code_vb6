VERSION 5.00
Begin VB.Form DirectorySort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save file as"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   FillStyle       =   0  'Solid
   ForeColor       =   &H00400000&
   Icon            =   "DirectorySort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox FileLocation 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Your computers directory."
      Top             =   405
      Width           =   6180
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Select files to be registerd or unregisterd."
      Top             =   3870
      Width           =   6165
   End
   Begin VB.DriveListBox Drive1 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3990
      TabIndex        =   0
      Top             =   30
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Select a folder for the backup file"
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label Label3 
      Caption         =   "Look in :"
      Height          =   255
      Left            =   3150
      TabIndex        =   7
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Save as :"
      Height          =   255
      Left            =   270
      TabIndex        =   6
      Top             =   2670
      Width           =   1515
   End
   Begin VB.Label LabNavControl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4140
      MouseIcon       =   "DirectorySort.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2670
      Width           =   945
   End
   Begin VB.Label LabNavControl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5190
      MouseIcon       =   "DirectorySort.frx":0594
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2670
      Width           =   945
   End
   Begin VB.Shape BkLabNavControl 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   1
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Shape BkLabNavControl 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   4110
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1005
   End
End
Attribute VB_Name = "DirectorySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
    File1 = Dir1
    ChDir Dir1
End Sub

Private Sub Drive1_Change()
  On Error GoTo 10
    Dir1 = Drive1
    ChDrive Drive1
10: Exit Sub
End Sub

Private Sub File1_Click()
    FileLocation = File1.Path & "\" & File1
End Sub

Private Sub LabNavControl_Click(Index As Integer)
Form1.Text1.Tag = FileLocation.Text
Unload DirectorySort

End Sub
