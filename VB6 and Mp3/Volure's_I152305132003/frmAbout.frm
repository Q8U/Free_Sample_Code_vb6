VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Master LoKi's ID3 Machine"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Master LoKi Revealed"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   253
         Left            =   242
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   7
         Top             =   3630
         Width           =   737
      End
      Begin ACTIVESKINLibCtl.SkinLabel Description 
         Height          =   1705
         Left            =   2783
         OleObjectBlob   =   "frmAbout.frx":0000
         TabIndex        =   6
         Top             =   1694
         Width           =   2673
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   858
         Left            =   242
         OleObjectBlob   =   "frmAbout.frx":005E
         TabIndex        =   5
         Top             =   1694
         Width           =   2310
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   979
         Left            =   242
         OleObjectBlob   =   "frmAbout.frx":0162
         TabIndex        =   4
         Top             =   847
         Width           =   5214
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   616
         Left            =   242
         OleObjectBlob   =   "frmAbout.frx":0384
         TabIndex        =   3
         Top             =   242
         Width           =   5214
      End
      Begin VB.CommandButton Command1 
         Caption         =   "http://volure.zapto.org"
         Height          =   375
         Left            =   1089
         TabIndex        =   2
         Top             =   3600
         Width           =   4335
      End
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "frmAbout.frx":0524
         Left            =   240
         List            =   "frmAbout.frx":0534
         TabIndex        =   1
         Top             =   2640
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
    'Open WebPage
    ShellExecute Me.hwnd, "open", Command1.Caption, "-nohome", "C:\", 1
End Sub

Private Sub Form_Load()
    frmTagger.Skin1.ApplySkin Me.hwnd 'Skin
    Picture1.Print "App.path"
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor
End Sub

Private Sub Picture1_DblClick()
    'Open the program Directory (Quick Access for Files therin)
    Shell "Explorer " & App.Path, vbNormalFocus
End Sub

Private Sub List1_Click()
    Select Case List1.List(List1.ListIndex)
        Case "Volure DarkAngel"
            Description = "Volure DarkAngel" & vbCrLf & _
                          "         Thats Me.. I am the Programmer of this Tagger... I used a couple things from PSCode.com and I have to Especially thank the maker of the Volume Control I used on this tagger.. With minor editing I make it work perfectly for my use" & vbCrLf & _
                          "" & vbCrLf & _
                          ""
        Case "pfcGentry"
            Description = "pfcGentry" & vbCrLf & _
                          "         A friend of mine online whom I chat to when Im not heavy into programming. Thanks for the Wasted time. Sometimes you gotta do something when you get tired of coding" & vbCrLf & _
                          "" & vbCrLf & _
                          ""
        Case "http://volure.zapto.org"
            Description = "http://volure.zapto.org" & vbCrLf & _
                          "         My Site to be coming very soon... Please check back frequently to find Other programs I make and Find all Updates" & vbCrLf & _
                          ""
        Case "^DragonFire^"
            Description = "^DragonFire^" & vbCrLf & _
                          "         Dragon has Helped me Debug the program and came up with Other Ideas such as Drag and Drop and the Tag Search" & vbCrLf & _
                          ""
        
    End Select
End Sub

