VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CyberCrypt"
   ClientHeight    =   5385
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7410
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5085
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2594
            MinWidth        =   2594
            Text            =   "Archive *"
            TextSave        =   "Archive *"
            Object.ToolTipText     =   "Archive"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2594
            MinWidth        =   2594
            Text            =   "Offset *"
            TextSave        =   "Offset *"
            Object.ToolTipText     =   "Selected file Offset"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2594
            MinWidth        =   2594
            Text            =   "File Size *"
            TextSave        =   "File Size *"
            Object.ToolTipText     =   "Selected file Size"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2594
            MinWidth        =   2594
            Text            =   "Files in archive 0"
            TextSave        =   "Files in archive 0"
            Object.ToolTipText     =   "Files in archive"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2594
            MinWidth        =   2594
            Text            =   "Sel file num 0"
            TextSave        =   "Sel file num 0"
            Object.ToolTipText     =   "Selected file number"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList FilePics 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":219E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3756
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4432
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":510E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6112
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7842
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8696
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9372
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A04E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C6E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CFBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E976
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F652
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1032E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1100A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12152
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":147E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1563A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":188D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":195AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A28A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A75E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B526
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E19E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED72
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21526
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22832
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23352
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23986
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":244E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2514A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25802
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":259C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":260BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26512
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26966
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2956E
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":299C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A26A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   2280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1800
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.pak"
      DialogTitle     =   "Open PAK"
      Filter          =   "PAK File (*.pak)|*.pak"
      Flags           =   38930
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B39A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BC76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C952
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D62E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E482
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ED5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FA3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FD56
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30202
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33692
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListGray 
      Left            =   600
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":344E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":351C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3677A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37456
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":386AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AC5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B936
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E0EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListFiles 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "FilePics"
      SmallIcons      =   "FilePics"
      ColHdrIcons     =   "FilePics"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Image OptionsPic 
      Height          =   495
      Left            =   4320
      Top             =   120
      Width           =   495
   End
   Begin VB.Image HelpTopics 
      Height          =   495
      Left            =   5160
      Top             =   120
      Width           =   495
   End
   Begin VB.Image AddPic 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      ToolTipText     =   "Add file"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image FileInfoPic 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      ToolTipText     =   "Selected file information"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image ExitPic 
      Height          =   480
      Left            =   6840
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image AboutPic 
      Height          =   495
      Left            =   6000
      ToolTipText     =   "About CyberCrypt"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image ExtractPic 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      ToolTipText     =   "Extract"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image OpenPic 
      Height          =   495
      Left            =   960
      ToolTipText     =   "Open Archive"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image NewPic 
      Height          =   495
      Left            =   120
      ToolTipText     =   "New Archive"
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu Menu01 
      Caption         =   "Extract"
      Visible         =   0   'False
      Begin VB.Menu Click01 
         Caption         =   "Selected file..."
      End
      Begin VB.Menu Click02 
         Caption         =   "All to directory..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\'
    '                                                            '
    '     ****                                                   '
    '   *   ||\*  |||  |||      ||| |\ | / ||| ||| |\ | /|| |||  '
    '  *    |  |* | |  ||| ---   |  | \| \  |   |  | \| |    |   '
    '  *    ||/ * |\   |   ---   |  | \|  \ |   |  | \| |    |   '
    '   *   |  *  | \  |||      ||| |  \ /  |  ||| |  \ \||  |   '
    '     ****                                                   '
    '           Software®                                        '
    '                                                            '
    '                                                            '
    '  Licensed Product                                          '
    '  Copyright © 1999-2001                                     '
    '  CyberCrypt Aditional Studio V2.0                          '
    '                                                            '
    '                                                            '
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\'

Private Sub AboutPic_Click()
    On Error Resume Next
    frmAbout.Show
End Sub

Private Sub AboutPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AboutPic.Picture = ImageListGray.ListImages(8).Picture
End Sub

Private Sub AddPic_Click()
    On Error GoTo FinaliseError
    If CommonDialog.FileName = "" Then
        MessageBox "You haven't opened any new or saved archive, do you want create a new archive?", YesNo, Question
        If Result = 1 Then
            CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
            CommonDialog.DialogTitle = "Save CyT file"
            CommonDialog.Filter = "CyT File (*.CyT)|*.CyT"
            CommonDialog.DefaultExt = ".CyT"
            CommonDialog.ShowSave
            If CommonDialog.FileName = "" Then Exit Sub
            Me.Caption = "CyberCrypt Added (" & CommonDialog.FileTitle & ") to (" & ArchiveName & ")"
            If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
            If CyTCreate(CommonDialog.FileName) = False Then Err.Raise 1
            ChkIfLoad = False
            ChkFastLoad = False
        ElseIf Result = 2 Then
            Exit Sub
        End If
        Exit Sub
    End If
    
    CommonDialog.Flags = &H1000 + &H4 + &H8 + &H800
    CommonDialog.DialogTitle = "ADD files to CyT"
    CommonDialog.Filter = "All files (*.*)|*.*"
    CommonDialog.DefaultExt = ""
    CommonDialog.ShowOpen
    frmBusy.Show
    Me.Enabled = False
    Me.MousePointer = 11
    CyTAdd CyTFile, CommonDialog.FileName, CommonDialog.FileTitle
    Unload frmBusy
    Me.MousePointer = 0
    Me.Enabled = True
    ChkIfLoad = False
    Exit Sub

FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If

End Sub

Private Sub AddPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AddPic.Picture = ImageListGray.ListImages(4).Picture
End Sub

Private Sub Click01_Click()
    
    On Error GoTo FinaliseError
    If ListFiles.SelectedItem = "" Then
        Exit Sub
    Else
        
        If ExtractPath <> "" Then
            
            If ChkWarningMsg = True Then
                MessageBox "Are you sure you want to extract the file selected?", YesNo, Question
                If Result = 1 Then
                    ExtractSelFileNoDlg ExtractPath & "\" & ListFiles.SelectedItem.Text
                    Exit Sub
                        ElseIf Result = 2 Then
                    Exit Sub
                End If
                    Else
                ExtractSelFileNoDlg ExtractPath & "\" & ListFiles.SelectedItem.Text
            End If
                
        Else
            
            If ChkWarningMsg = True Then
                MessageBox "Are you sure you want to extract the file selected?", YesNo, Question
                If Result = 1 Then
                    ExtractSelFile
                    Exit Sub
                        ElseIf Result = 2 Then
                    Exit Sub
                End If
                    Else
                ExtractSelFile
            End If
        
        End If
    End If
                        
    Exit Sub
                    
FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
                    
End Sub

Private Sub ExtractSelFileNoDlg(Path As String)
    On Error GoTo FinaliseError
    ArchiveName = RemoveBackSlash(CyTFile)
    CommonDialog.FileName = Path
    CommonDialog.FileTitle = ListFiles.SelectedItem.Text
    If CommonDialog.FileName = "" Then Exit Sub
    Me.Caption = "CyberCrypt Extracted (" & CommonDialog.FileTitle & ") from (" & ArchiveName & ")"
    If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
    frmBusy.Show
    Me.Enabled = False
    Me.MousePointer = 11
    If CyTExtract(CyTFile, ListFiles.SelectedItem.Text, CommonDialog.FileName) = False Then MessageBox "An error occured when trying to extract the file!", OKOnly, Critical
    Me.Enabled = True
    Me.MousePointer = 0
    Unload frmBusy
    
    If LoadArchive = True Then
        CyTOpen CommonDialog.FileName
        LoadArchive = False
        Me.Caption = "CyberCrypt Extracted (" & CommonDialog.FileTitle & ")"
    End If
    
    Exit Sub

FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
End Sub

Private Sub ExtractSelFile()
    On Error GoTo FinaliseError
    ArchiveName = RemoveBackSlash(CyTFile)
    CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
    CommonDialog.DialogTitle = "Save file"
    CommonDialog.Filter = "All files (*.*)|*.*"
    CommonDialog.DefaultExt = ""
    CommonDialog.FileName = ListFiles.SelectedItem.Text
    CommonDialog.ShowSave
    If CommonDialog.FileName = "" Then Exit Sub
    Me.Caption = "CyberCrypt Extracted (" & CommonDialog.FileTitle & ") from (" & ArchiveName & ")"
    If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
    frmBusy.Show
    Me.Enabled = False
    Me.MousePointer = 11
    If CyTExtract(CyTFile, ListFiles.SelectedItem.Text, CommonDialog.FileName) = False Then MessageBox "An error occured when trying to extract the file!", OKOnly, Critical
    Me.Enabled = True
    Me.MousePointer = 0
    Unload frmBusy
    
    If LoadArchive = True Then
        CyTOpen CommonDialog.FileName
        LoadArchive = False
        Me.Caption = "CyberCrypt Extracted (" & CommonDialog.FileTitle & ")"
    End If
    
    Exit Sub

FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
End Sub

Private Sub ExtractAllTODir(Path As String)
    On Error GoTo FinaliseError
    For M = 1 To frmMain.ListFiles.ListItems.Count
        frmMain.Caption = "CyberCrypt Extracted file to (" & Path & ") from (" & ArchiveName & ")"
        If FileExist(Path & "\" & frmMain.ListFiles.ListItems(M)) = True Then Kill Path & "\" & frmMain.ListFiles.ListItems(M)
        frmBusy.Show
        Me.Enabled = False
        Me.MousePointer = 11
        If frmMain.CyTExtract(CyTFile, frmMain.ListFiles.ListItems(M), Path & "\" & frmMain.ListFiles.ListItems(M)) = False Then MessageBox "An error occured when trying to extract the file(s)!", OKOnly, Critical:  Me.Enabled = True: Me.MousePointer = 0: Unload frmBusy: Exit For
        Me.Enabled = True
        Me.MousePointer = 0
        Unload frmBusy
    Next M
    Exit Sub
    
FinaliseError:
    MessageBox "An error occured when trying to extract the file(s)!", OKOnly, Critical
    Unload frmBusy
    Exit Sub
End Sub

Private Sub Click02_Click()
    
    On Error GoTo FinaliseError
    If ExtractPath <> "" Then
        
        If ChkWarningMsg = True Then
            MessageBox "Are you sure you want to extract all?", YesNo, Question
            If Result = 1 Then
                ExtractAllTODir ExtractPath
                Exit Sub
                    ElseIf Result = 2 Then
                Exit Sub
            End If
                Else
            ExtractAllTODir ExtractPath
        End If
            
    Else
        
        If ChkWarningMsg = True Then
            MessageBox "Are you sure you want to extract all?", YesNo, Question
            If Result = 1 Then
                FrmDir.Show , Me
                Exit Sub
                    ElseIf Result = 2 Then
                Exit Sub
            End If
                Else
            FrmDir.Show , Me
        End If
    
    End If
                        
    Exit Sub
                    
FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
        
End Sub

Private Sub ExitPic_Click()
    End
End Sub

Private Sub GetFileData()
    On Error GoTo FinaliseError
    If ListFiles.SelectedItem = "" Then
        Exit Sub
    Else
        If FileExist(SystemRootS & "\Temp\" & ListFiles.SelectedItem) = True Then Kill SystemRootS & "\Temp\" & ListFiles.SelectedItem
        frmBusy.Show
        Me.Enabled = False
        Me.MousePointer = 11
        CyTExtract CyTFile, ListFiles.SelectedItem, SystemRootS & "\Temp\" & ListFiles.SelectedItem
        Me.Enabled = True
        Me.MousePointer = 0
        Unload frmBusy
        Select Case Right(LCase$(ListFiles.SelectedItem), 3)
            'Text files
            'Case LCase$("txt"): <Text Viewer Form>: Exit Sub
            'Case LCase$("ini"): <Text Viewer Form>: Exit Sub
            'Case LCase$("inf"): <Text Viewer Form>: Exit Sub
            'Case LCase$("cfg"): <Text Viewer Form>: Exit Sub
            'Case LCase$("log"): <Text Viewer Form>: Exit Sub
            'Case LCase$("bat"): <Text Viewer Form>: Exit Sub
        End Select
        FrmFileInfo.Show 1, Me
    End If
    
    Exit Sub

FinaliseError:

End Sub

Private Sub ExitPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ExitPic.Picture = ImageListGray.ListImages(7).Picture
End Sub

Private Sub ExtractPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ExtractPic.Picture = ImageListGray.ListImages(3).Picture
End Sub

Private Sub ExtractPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PopupMenu Menu01
    End If
End Sub

Private Sub FileInfoPic_Click()
    Me.Caption = "CyberCrypt Opened fileInfo (" & ListFiles.SelectedItem & ") in (" & ArchiveName & ")"
    GetFileData
End Sub

Private Sub FileInfoPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
End Sub

Private Sub GetInIData()
    Dim LoadResult As String
    Dim LoadResultB As String
    On Error Resume Next
    LoadResult = String(100, "*")
    lLength = Len(LoadResult)
    
    GetPrivateProfileString "Settings", "SmallIcons", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then frmMain.ListFiles.View = lvwSmallIcon
    
    GetPrivateProfileString "Settings", "Lists", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then frmMain.ListFiles.View = lvwList

    GetPrivateProfileString "Settings", "NormalIcons", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then frmMain.ListFiles.View = lvwIcon
    
    GetPrivateProfileString "Settings", "Report", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then frmMain.ListFiles.View = lvwReport
    
    LoadResult = String(100, "*")
    lLength = Len(LoadResult)
        
    GetPrivateProfileString "Settings", "SendTo", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then
              
        sString = String(100, "*")
        lLength = Len(sString)
        
        GetPrivateProfileString "Settings", "ExPath", vbNullString, sString, lLength, App.Path & "\Settings.ini"
        
        'This codes here because when the string loads in a veriable
        'it adds on one blank (chr(0)) and a return of lLength (*).
        'So this code takes the end and the chr(0) away. This
        'code is needed only because you are loading a string that
        'has been written into the ini without speech marks around the
        'it. This is easy to change but it is safer to use as
        'the code might think that theirs a speech mark in path
        'that loads into the veraible (ExtractPath).
        
        sString = Mid(sString, 1, InStr(1, sString, Chr$(0)) - 1)
        ExtractPath = sString
        
        'For M = 1 To Len(sString)
        '    getchr1 = Left(sString, M)
        '    getchr2 = Right(getchr1, 1)
        '    If Asc(getchr2) = 0 Then ExtractPath = Left(sString, M - 1): Exit For
        'Next M
        
        'If the user has not yet entered the settings or if the prompt
        'option is used and the extract path has not been specified then
        'the veriable (ExtractPath) may become "False" so this code checks
        'the case and sees if the veriable string is "False" and sets it
        'to "" as "False" is an invalid folder of drive name on it's own.
        
        If LCase$(ExtractPath) = LCase$("False") Then ExtractPath = ""
        
    End If
    
    LoadResult = String(100, "*")
    lLength = Len(LoadResult)

    GetPrivateProfileString "Settings", "Prompt", vbNullString, LoadResult, lLength, App.Path & "\Settings.ini"
    If LCase$(LoadResult) = True Then ExtractPath = ""
    
    LoadResultB = String(100, "*")
    lLength = Len(LoadResultB)

    GetPrivateProfileString "Settings", "Sort", vbNullString, LoadResultB, lLength, App.Path & "\Settings.ini"
    If LoadResultB = 1 Then frmMain.ListFiles.Sorted = True Else frmMain.ListFiles.Sorted = False

    GetPrivateProfileString "Settings", "Grid", vbNullString, LoadResultB, lLength, App.Path & "\Settings.ini"
    If LoadResultB = 1 Then frmMain.ListFiles.GridLines = True Else frmMain.ListFiles.GridLines = False
    
    GetPrivateProfileString "Settings", "Warn", vbNullString, LoadResultB, lLength, App.Path & "\Settings.ini"
    If LoadResultB = 1 Then ChkWarningMsg = True Else ChkWarningMsg = False
    If LoadResultB = String(100, "*") Then ChkWarningMsg = False
        
End Sub

Private Sub Form_Load()
    
    FolIndex = 0
    ChkFastLoad = False
    
    GetInIData

    SystemRootS = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION", "SystemRoot")
    ListFiles.ColumnHeaders.Add , , "Filename", 2400
    ListFiles.ColumnHeaders.Add , , "File type", 2100
    ListFiles.ColumnHeaders.Add , , "Size", 2100
    ListFiles.ColumnHeaders.Add , , "Offset", 1500
    ListFiles.ColumnHeaders.Add , , "File number", 1000
    NewPic.Picture = ImageList.ListImages(1).Picture
    OpenPic.Picture = ImageList.ListImages(2).Picture
    AddPic.Picture = ImageListGray.ListImages(4).Picture
    ExtractPic.Picture = ImageListGray.ListImages(3).Picture
    FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
    AboutPic.Picture = ImageList.ListImages(8).Picture
    HelpTopics.Picture = ImageList.ListImages(10).Picture
    OptionsPic.Picture = ImageList.ListImages.Item(11).Picture
    ExitPic.Picture = ImageList.ListImages(7).Picture
        
    ChkIfLoad = False

    If Command <> "" Then
        ChkIfLoad = True
        On Error GoTo FinaliseError
        CommonDialog.FileName = Command
        
        For M = 1 To Len(Command)
            GetChr0 = Right(Command, M)
            getchr1 = Left(GetChr0, 1)
            If getchr1 = "\" Or getchr1 = "/" Then
                ArchiveName = Right(GetChr0, M - 1): Exit For
            End If
        Next M
        Me.Caption = "CyberCrypt (" & ArchiveName & ")"
        StatusBar.Panels(1).Text = "Archive (" & ArchiveName & ")"
        If FileExist(Command) = True Then
            CyTFile = Command
            CyTOpen Command
        End If
        Exit Sub
        
FinaliseError:
        
        If Err = 32755 Then
            Exit Sub
                Else
            MessageBox "An unknown error occured!", OKOnly, Critical
            End
        End If
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMenuPicsDefault
End Sub

Private Sub SetMenuPicsDefault()
    NewPic.Picture = ImageList.ListImages(1).Picture
    OpenPic.Picture = ImageList.ListImages(2).Picture
    If AddPic.Enabled = True Then
        AddPic.Picture = ImageList.ListImages(4).Picture
    End If
    If ExtractPic.Enabled = True Then
        ExtractPic.Picture = ImageList.ListImages(3).Picture
    End If
    If FileInfoPic.Enabled = True Then
        FileInfoPic.Picture = ImageList.ListImages(6).Picture
    End If
    AboutPic.Picture = ImageList.ListImages(8).Picture
    HelpTopics.Picture = ImageList.ListImages(10).Picture
    OptionsPic.Picture = ImageList.ListImages.Item(11).Picture
    ExitPic.Picture = ImageList.ListImages(7).Picture
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If frmMain.Width <= 7530 Then frmMain.Width = 7530
    If frmMain.Height <= 3290 Then frmMain.Height = 3290
    
    For M = 1 To StatusBar.Panels.Count
        StatusBar.Panels(M).Width = frmMain.Width / StatusBar.Panels.Count
    Next M
    
    ListFiles.Width = frmMain.Width - 100
    ListFiles.Height = frmMain.Height - ListFiles.Top - StatusBar.Height - 400

End Sub

Private Sub HelpTopics_Click()
    On Error GoTo FinaliseError
    Shell App.Path & "\HelpTopics.exe", vbNormalFocus
    Exit Sub
FinaliseError:
    MessageBox "Error, Help topics could not be found.", OKOnly, Critical
End Sub

Private Sub HelpTopics_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpTopics.Picture = ImageListGray.ListImages(10).Picture
End Sub

Private Sub ListFiles_Click()
    GetListData
End Sub

Private Sub GetListData()
    On Error Resume Next
    If ArchiveName = "" Then
        StatusBar.Panels(1).Text = "Archive *"
            Else
        StatusBar.Panels(1).Text = "Archive (" & ArchiveName & ")"
    End If
    StatusBar.Panels(5).Text = "Sel file num " & ListFiles.SelectedItem.Index
    If ListFiles.SelectedItem.SubItems(1) <> "" Then
        StatusBar.Panels(2).Text = "Offset (" & ListFiles.SelectedItem.SubItems(3) & ")"
            Else
        StatusBar.Panels(2).Text = "Offset *"
    End If
    If ListFiles.SelectedItem.SubItems(2) <> "" Then
        StatusBar.Panels(3).Text = "File size (" & ListFiles.SelectedItem.SubItems(2) & ")"
            Else
        StatusBar.Panels(3).Text = "File size *"
    End If
End Sub

Private Sub ListFiles_DblClick()
    If CyTFile = "" Then Exit Sub
    If ExtractPic.Enabled = False Then Exit Sub
    If Right(ListFiles.SelectedItem.Text, 3) = "CyT" Then
        MessageBox "You cannot open this type of file form here. You have to extract it first. Would you like to extract now?", OKCancel, Question
        If Result = 3 Then
            LoadArchive = False
            Exit Sub
                Else
            LoadArchive = True
            Click01_Click
        End If
            Else
        LoadArchive = False
        Click01_Click
    End If
End Sub

Private Sub ListFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    GetListData
End Sub

Private Sub ListFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMenuPicsDefault
End Sub

Private Sub ListFiles_OLEDragDrop(Data As MSComCtlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo Erro
    
        For D = 1 To Len(Data.Files(1))
            GetChr0 = Left(Data.Files(1), D)
            getchr1 = Right(GetChr0, 1)
            If Len(GetChr0) = Len(Data.Files(1)) Then
                MessageBox "You cannot drag files, folders into the archive without no file extensions.", OKOnly, Warning
                Exit Sub
            End If
            If getchr1 = "." Then
                If Right(Data.Files(1), 3) = "CyT" Then
                    For M = 1 To Len(Data.Files(1))
                        GetChr0 = Right(Data.Files(1), M)
                        getchr1 = Left(GetChr0, 1)
                        If getchr1 = "\" Or getchr1 = "/" Then
                            TmpFile = Right(GetChr0, M - 1): Exit For
                        End If
                    Next M
                    Me.Caption = "CyberCrypt (" & TmpFile & ")"
                    ArchiveName = TmpFile
                    If FileExist(Data.Files(1)) = True Then
                        CyTOpen Data.Files(1)
                        CommonDialog.FileName = Data.Files(1)
                    End If
                    Exit Sub
                End If
                
                If CyTFile = "" Then
                    MessageBox "You haven't opened any new or saved archive, do you want create a new archive?", YesNo, Question
                    If Result = 1 Then
                        CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
                        CommonDialog.DialogTitle = "Save CyT file"
                        CommonDialog.Filter = "CyT File (*.CyT)|*.CyT"
                        CommonDialog.ShowSave
                        If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
                        CyTCreate CommonDialog.FileName
                        ChkIfLoad = False
                    ElseIf Result = 2 Then
                        Exit Sub
                    End If
                End If
                
                frmBusy.Visible = True
                Me.Enabled = False
                Me.MousePointer = 11
                
                'Adds the file to the archive
                For Files = 1 To Data.Files.Count
                    CyTAdd CyTFile, Data.Files(Files), RemoveBackSlash(Data.Files(Files))
                    Me.Refresh
                    For M = 1 To Len(Data.Files(Files))
                        GetChr0 = Right(Data.Files(Files), M)
                        getchr1 = Left(GetChr0, 1)
                        If getchr1 = "\" Or getchr1 = "/" Then
                            TmpFile = Right(GetChr0, M - 1): Exit For
                        End If
                    Next M
                    Me.Caption = "CyberCrypt Dragged (" & TmpFile & ") into archive."
                Next Files
                frmBusy.lblFile.Caption = "Updating archive..."
                CyTOpen CyTFile
                ExtractPic.Enabled = True: ExtractPic.Picture = ImageList.ListImages(3).Picture
                FileInfoPic.Enabled = True: FileInfoPic.Picture = ImageList.ListImages(6).Picture
                AddPic.Enabled = True: AddPic.Picture = ImageList.ListImages(4).Picture
                Unload frmBusy
                Me.Enabled = True
                Me.MousePointer = 0
                ChkIfLoad = False
                ChkFastLoad = True
                Exit Sub
            End If
        Next D
        
Erro:
    If Err = 32755 Then
        CyTFile = ""
        CommonDialog.FileName = ""
        ExtractPic.Enabled = False: ExtractPic.Picture = ImageListGray.ListImages(3).Picture
        FileInfoPic.Enabled = False: FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
        AddPic.Enabled = False: AddPic.Picture = ImageListGray.ListImages(4).Picture
        ListFiles.ListItems.Clear
        Me.Caption = "CyberCrypt (No new file)"
        Exit Sub
    Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
End Sub

Private Sub NewPic_Click()
    On Error GoTo FinaliseError
    CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
    CommonDialog.DialogTitle = "Save CyT file"
    CommonDialog.Filter = "CyT File (*.CyT)|*.CyT"
    CommonDialog.DefaultExt = ".CyT"
    CommonDialog.ShowSave
    If CommonDialog.FileName = "" Then Exit Sub
    ArchiveName = CommonDialog.FileTitle
    Me.Caption = "CyberCrypt (" & ArchiveName & ")"
    StatusBar.Panels(1).Text = "Archive (" & ArchiveName & ")"
    StatusBar.Panels(2).Text = "Offset *"
    StatusBar.Panels(3).Text = "File size *"
    StatusBar.Panels(4).Text = "Files in archive 0"
    StatusBar.Panels(5).Text = "Sel file num 0"
    If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
    If CyTCreate(CommonDialog.FileName) = False Then Err.Raise 1
    ExtractPic.Enabled = False: ExtractPic.Picture = ImageListGray.ListImages(3).Picture
    FileInfoPic.Enabled = False: FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
    AddPic.Enabled = True: AddPic.Picture = ImageList.ListImages(4).Picture
    ListFiles.ListItems.Clear
    ChkIfLoad = False
    Exit Sub
FinaliseError:
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
End Sub

Private Sub NewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewPic.Picture = ImageListGray.ListImages(1).Picture
End Sub

Private Sub OpenPic_Click()
    On Error GoTo FinaliseError
    CommonDialog.Flags = &H1000 + &H4 + &H8 + &H800
    CommonDialog.Filter = "CyT File (*.CyT)|*.CyT"
    CommonDialog.DialogTitle = "Open CyT file"
    CommonDialog.DefaultExt = ""
    CommonDialog.ShowOpen
    If CommonDialog.FileName = "" Then Exit Sub
    ArchiveName = CommonDialog.FileTitle
    Me.Caption = "CyberCrypt (" & ArchiveName & ")"
    ChkFastLoad = False
    StatusBar.Panels(1).Text = "Archive (" & ArchiveName & ")"
    If FileExist(CommonDialog.FileName) = True Then
        CyTOpen CommonDialog.FileName
        ChkIfLoad = False
    End If
    
    Exit Sub
    
FinaliseError:
    
    If Err = 32755 Then
        Exit Sub
            Else
        MessageBox "An unknown error occured!", OKOnly, Critical
        End
    End If
End Sub

Function CyTOpen(FileName As String) As Boolean
    Dim FileList As String
    Dim Offset As Long
    Dim Size As Long
    Dim Name As String
    Dim LF As ListItem
    Dim LFS As ListSubItem
    
    On Error GoTo Erro
    
    ListFiles.ListItems.Clear
    
    'Check if is a valid CyT file
    If CyTValid(FileName) = True Then
        CyTOpen = True
        FileNumber = FreeFile
        Open FileName For Binary As FileNumber
            'Is a valid CyT file
            'Get the FileList
            Get FileNumber, 7, FileListStart
            
            If FileListStart = 0 Then
                CyTFile = ""
                CyTOpen = False
                CommonDialog.FileName = ""
                Me.Caption = "CyberCrypt (No new file)"
                MessageBox "Empty archive!", OKOnly, Information
                Close FileNumber
                If Command <> "" And ChkIfLoad = True Then End
                Exit Function
            Else
                CyTFile = FileName
                ExtractPic.Enabled = True: ExtractPic.Picture = ImageList.ListImages(3).Picture
                FileInfoPic.Enabled = True: FileInfoPic.Picture = ImageList.ListImages(6).Picture
                AddPic.Enabled = True: AddPic.Picture = ImageList.ListImages(4).Picture
                
                'Add the FileName, OffSet and Size in the ListView control but first clears the ListView
                ListFiles.ListItems.Clear
                
                Do
                    Get FileNumber, FileListStart, Offset
                    FileListStart = FileListStart + 4
                    
                    Get FileNumber, FileListStart, Size
                    FileListStart = FileListStart + 4
                    
                    Name = String$(255, Chr$(0))
                    Get FileNumber, FileListStart, Name
                    Name = Mid(Name, 1, InStr(1, Name, Chr$(0)) - 1)
                    FileListStart = FileListStart + Len(Name) + 1
                                        
                    If Name = "" Or Offset = 0 Or Size = 0 Then
                        CyTFile = ""
                        CyTOpen = False
                        CommonDialog.FileName = ""
                        Me.Caption = "CyberCrypt (No new file)"
                        MessageBox "Empty archive!", OKOnly, Information
                        Close FileNumber
                        If Command <> "" And ChkIfLoad = True Then End
                        Exit Function
                    End If
                
                    FindTypeIcon Name, Offset, Size
                    
                Loop Until FileListStart > LOF(FileNumber)

            End If
            
        Else
            'Is a invalid CyT file
            CyTOpen = False
            CommonDialog.FileName = ""
            CyTFile = ""
            Me.Caption = "CyberCrypt (No new file)"
            StatusBar.Panels(1).Text = "Archive *"
            StatusBar.Panels(2).Text = "Offset *"
            StatusBar.Panels(3).Text = "File size *"
            StatusBar.Panels(4).Text = "Files in archive 0"
            StatusBar.Panels(5).Text = "Sel file num 0"
            ExtractPic.Enabled = False: ExtractPic.Picture = ImageListGray.ListImages(3).Picture
            FileInfoPic.Enabled = False: FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
            AddPic.Enabled = False: AddPic.Picture = ImageListGray.ListImages(4).Picture
            MessageBox "The specified filename is not a valid archive!", OKOnly, Critical
            Close FileNumber
            If Command <> "" And ChkIfLoad = True Then End
            Exit Function
        End If
    Close FileNumber
    Exit Function
    
Erro:
    If Err = 5 Then
        CyTOpen = False
        CyTFile = ""
        CommonDialog.FileName = ""
        ListFiles.ListItems.Clear
        Me.Caption = "CyberCrypt (No new file)"
        StatusBar.Panels(1).Text = "Archive *"
        StatusBar.Panels(2).Text = "Offset *"
        StatusBar.Panels(3).Text = "File size *"
        StatusBar.Panels(4).Text = "Files in archive 0"
        StatusBar.Panels(5).Text = "Sel file num 0"
        ExtractPic.Enabled = False: ExtractPic.Picture = ImageListGray.ListImages(3).Picture
        FileInfoPic.Enabled = False: FileInfoPic.Picture = ImageListGray.ListImages(6).Picture
        AddPic.Enabled = False: AddPic.Picture = ImageListGray.ListImages(4).Picture
        MessageBox "An error occured when trying to read the archive!", OKOnly, Critical
        Close FileNumber
        If Command <> "" And ChkIfLoad = True Then End
        Exit Function
    End If
End Function

Private Sub FindTypeIcon(Name As String, Offset As Long, Size As Long)

    'I've used this code because some programs have rubish icons
    'so i've added my own, so this means that any file anything
    'other than the file extensions below becomes an unknown file
    'icon. The file will still appear though.

    Dim LF As ListItem
    Dim LFS As ListSubItem
    Dim PicIndex As Long
    Dim NameTemp As String

    PicIndex = 36
    
    NameTemp = LCase$(Right(Name, 3))
    
    Select Case NameTemp
        Case LCase$("abt"): PicIndex = 1: GoTo ChkFormat
        Case LCase$("avi"): PicIndex = 5: GoTo ChkFormat
        Case LCase$("bat"): PicIndex = 6: GoTo ChkFormat
        Case LCase$("bmp"): PicIndex = 7: GoTo ChkFormat
        Case LCase$("cyt"): PicIndex = 9: GoTo ChkFormat
        Case LCase$("dll"): PicIndex = 14: GoTo ChkFormat
        Case LCase$("sys"): PicIndex = 14: GoTo ChkFormat
        Case LCase$("top"): PicIndex = 15: GoTo ChkFormat
        Case LCase$("exe"): PicIndex = 16: GoTo ChkFormat
        Case LCase$("com"): PicIndex = 16: GoTo ChkFormat
        Case LCase$("ext"): PicIndex = 17: GoTo ChkFormat
        Case LCase$("zip"): PicIndex = 19: GoTo ChkFormat
        Case LCase$("gif"): PicIndex = 20: GoTo ChkFormat
        Case LCase$("ini"): PicIndex = 23: GoTo ChkFormat
        Case LCase$("inf"): PicIndex = 23: GoTo ChkFormat
        Case LCase$("css"): PicIndex = 23: GoTo ChkFormat
        Case LCase$("dat"): PicIndex = 23: GoTo ChkFormat
        Case LCase$("jpg"): PicIndex = 24: GoTo ChkFormat
        Case LCase$("pwl"): PicIndex = 25: GoTo ChkFormat
        Case LCase$("mid"): PicIndex = 26: GoTo ChkFormat
        Case LCase$("mp3"): PicIndex = 27: GoTo ChkFormat
        Case LCase$("mpg"): PicIndex = 28: GoTo ChkFormat
        Case LCase$("mpe"): PicIndex = 28: GoTo ChkFormat
        Case LCase$("stp"): PicIndex = 33: GoTo ChkFormat
        Case LCase$("wav"): PicIndex = 34: GoTo ChkFormat
        Case LCase$("wma"): PicIndex = 34: GoTo ChkFormat
        Case LCase$("txt"): PicIndex = 35: GoTo ChkFormat
        Case LCase$("log"): PicIndex = 35: GoTo ChkFormat
        Case LCase$("cfg"): PicIndex = 35: GoTo ChkFormat
        Case LCase$("usr"): PicIndex = 37: GoTo ChkFormat
        Case LCase$("hlp"): PicIndex = 38: GoTo ChkFormat
        Case LCase$("ico"): PicIndex = 36: GoTo ChkFormat
        Case LCase$("htm"): PicIndex = 39: GoTo ChkFormat
        Case LCase$("wmf"): PicIndex = 21: GoTo ChkFormat
    End Select

    NameTemp = LCase$(Right(Name, 4))

    Select Case NameTemp
        Case LCase$("jpeg"): PicIndex = 24: GoTo ChkFormat
        Case LCase$("user"): PicIndex = 37: GoTo ChkFormat
        Case LCase$("html"): PicIndex = 39: GoTo ChkFormat
    End Select
    
        PicIndex = 36
        
ChkFormat:

    'Checks if file is recognised as a setup file or an uninstall file, only by name
    NameTemp = LCase$(Left(Name, 5))
    If NameTemp = LCase$("setup") And LCase$(Right(Name, 3)) = "exe" Then PicIndex = 33
    NameTemp = LCase$(Left(Name, 6))
    If NameTemp = LCase$("install") And LCase$(Right(Name, 3)) = "exe" Then PicIndex = 33
    NameTemp = LCase$(Left(Name, 9))
    If NameTemp = LCase$("uninstall") And LCase$(Right(Name, 3)) = "exe" Then PicIndex = 33

    NameTemp = LCase$(Right(Name, 3))

    Set LF = ListFiles.ListItems.Add(, , Name, PicIndex, PicIndex)
    
    'Gets the file type and puts it into the listview
    If GetFileTypeName(NameTemp & "file") = LCase$("<Unknown file type>") Or Len(GetFileTypeName(NameTemp & "file")) >= 88 Then
        If GetFileTypeName("." & NameTemp) = LCase$("<Unknown file type>") Or Len(GetFileTypeName("." & NameTemp)) >= 88 Then
            Set LFS = LF.ListSubItems.Add(, , UCase$(NameTemp) & " file")
                Else
            Set LFS = LF.ListSubItems.Add(, , GetFileTypeName("." & NameTemp))
        End If
            Else
        Set LFS = LF.ListSubItems.Add(, , GetFileTypeName("." & NameTemp))
    End If
    
    Set LFS = LF.ListSubItems.Add(, , CStr(FormatKB(Size)) & "  (" & Size & " bytes)")
    Set LFS = LF.ListSubItems.Add(, , Offset)
    Set LFS = LF.ListSubItems.Add(, , ListFiles.ListItems.Count)
    
    LoadPanelInfo

End Sub

Private Sub LoadPanelInfo()
    StatusBar.Panels(4).Text = "Files in archive " & ListFiles.ListItems.Count
    GetListData
End Sub

Private Sub KillFile(Name As String)
    On Error Resume Next
    Kill SystemRootS & "\Temp\" & Name
End Sub

Private Sub FindTypeIconClone(NameADD As String, OffSetADD As Long, SizeADD As Long)
    CyTOpen CyTFile
    Exit Sub
End Sub

Function CyTCreate(FileName As String) As Boolean
    On Error GoTo Erro
    Dim FileList As String
    
    Header = "CYT1.0"
    FileListStart = 0
    
    If FileExist(FileName) = True Then
        CyTCreate = False
        Exit Function
    Else
        FileNumber = FreeFile
        Open FileName For Binary As FileNumber
            Put FileNumber, 1, Header
            Put FileNumber, Len(Header) + 1, FileListStart
        Close FileNumber
    End If
    CyTFile = FileName
    ExtractPic.Enabled = True: ExtractPic.Picture = ImageList.ListImages(3).Picture
    FileInfoPic.Enabled = True: FileInfoPic.Picture = ImageList.ListImages(6).Picture
    AddPic.Enabled = True: AddPic.Picture = ImageList.ListImages(4).Picture
    CyTCreate = True
    Exit Function
    
Erro:
    If Err <> 0 Then
        CyTCreate = False
        Exit Function
    End If
End Function

Function CyTAdd(FileCyT As String, FileADD As String, NameADD As String) As Boolean
    On Error GoTo Erro
    Dim BytesADD As String
    Dim OffSetADD As Long
    Dim SizeADD As Long
    Dim LF As ListItem
    Dim LFS As ListSubItem
    
    NameADD = NameADD & Chr$(0)
    
    If FileExist(FileCyT) = False Or FileExist(FileADD) = False Then
        CyTAdd = False
        Exit Function
    Else
        'Check if is a valid CyT file
        If CyTValid(FileCyT) = True Then
            'Is a valid CyT file
            
            FileNumberCyT = FreeFile
            Open FileCyT For Binary As FileNumberCyT
            
            'Get the FileList
            Get FileNumberCyT, 7, FileListStart
    
            'Get the FileList and put in the memory
            If FileListStart = 0 Then
                FileListStart = LOF(FileNumberCyT) + 1
                FileList = ""
            Else
                FileList = String(LOF(FileNumberCyT) - FileListStart + 1, Chr$(0))
                Get FileNumberCyT, FileListStart, FileList
            End If
    
            OffSetADD = FileListStart
            SizeADD = FileLen(FileADD)
                
            'Put the file inside of the CyT
            FileNumberADD = FreeFile
            frmBusy.lblFile = "Adding " & RemoveBackSlash(FileADD)
            frmBusy.Refresh
            Open FileADD For Binary As FileNumberADD
                If LOF(FileNumberADD) > 1000000 Then 'Divid the file in parts to use less memory and make less swap
                    'BytesADD = String(LOF(FileNumberADD) / 100, Chr$(0))
                    'For Position = 1 To LOF(FileNumberADD) Step Len(BytesADD)
                        'Get FileNumberADD, Position, BytesADD
                        'Put FileNumberCyT, FileListStart, BytesADD
                        'FileListStart = FileListStart + Len(BytesADD)
                    'Next Position
                    
                    Position = -999999
                    frmBusy.prgFile.Max = LOF(FileNumberADD)
                    Do
                        Position = Position + 1000000
                        If Position + 999999 > LOF(FileNumberADD) Then
                            frmBusy.prgFile.Value = frmBusy.prgFile.Max
                            frmBusy.Refresh
                            BytesADD = String(LOF(FileNumberADD) - Position + 1, Chr$(0))
                        Else
                            frmBusy.prgFile.Value = Position
                            frmBusy.Refresh
                            BytesADD = String(1000000, Chr$(0))
                        End If
                        Get FileNumberADD, Position, BytesADD
                        Put FileNumberCyT, FileListStart, BytesADD
                        FileListStart = FileListStart + Len(BytesADD)
                    Loop Until Position + 999999 > LOF(FileNumberADD)
                    
                Else
                    frmBusy.prgFile.Max = 1
                    frmBusy.prgFile.Value = 0
                    BytesADD = String(LOF(FileNumberADD), Chr$(0))
                    Get FileNumberADD, 1, BytesADD
                    Put FileNumberCyT, FileListStart, BytesADD
                    FileListStart = FileListStart + Len(BytesADD)
                    frmBusy.prgFile.Value = 1
                End If
            Close FileNumberADD
            
            'Add the new file in the FileList
            Put FileNumberCyT, 7, FileListStart
            Put FileNumberCyT, FileListStart, FileList
            Put FileNumberCyT, FileListStart + Len(FileList), OffSetADD
            Put FileNumberCyT, FileListStart + Len(FileList) + 4, SizeADD
            Put FileNumberCyT, FileListStart + Len(FileList) + 8, NameADD
            Close FileNumberCyT
        Else
            CyTAdd = False
            Exit Function
        End If
    End If
    CyTAdd = True
    If ChkFastLoad = True Then
        FindTypeIcon NameADD, OffSetADD, SizeADD
            Else
        FindTypeIconClone NameADD, OffSetADD, SizeADD
    End If
    Exit Function
    
Erro:
    CyTAdd = False
    Exit Function
End Function

Function CyTValid(CyTFileName As String) As Boolean
    Dim Header As String
    Header = String$(6, Chr$(0))
    
    If FileExist(CyTFileName) = False Then
        CyTValid = False
        Exit Function
    Else
        FileNumber = FreeFile
        Open CyTFileName For Binary As FileNumber
            Get FileNumber, 1, Header
            If Header = "CYT1.0" Then
                CyTValid = True
            Else
                CyTValid = False
            End If
        Close FileNumber
    End If
End Function

Function CyTExtract(CyTFile As String, FileToExtract As String, DestinationFile As String) As Boolean
    'On Error GoTo FinaliseError
    Dim BytesExtract As String
    Dim Offset As Long
    Dim Size As Long
    Dim Name As String
    
    If FileExist(CyTFile) = False Or FileExist(DestinationFile) = True Then
        CyTExtract = False
        Exit Function
    Else
        If CyTValid(CyTFile) = True Then
        
            FileNumber = FreeFile
            Open CyTFile For Binary As FileNumber
                'Get the FileList
                Get FileNumber, 7, FileListStart
            
                If FileListStart = 0 Then
                    CyTExtract = False
                    Close FileNumber
                    Exit Function
                Else
                    
    
                    Do
                        Get FileNumber, FileListStart, Offset
                        FileListStart = FileListStart + 4
                    
                        Get FileNumber, FileListStart, Size
                        FileListStart = FileListStart + 4
                    
                        Name = String$(255, Chr$(0))
                        Get FileNumber, FileListStart, Name
                        Name = Mid(Name, 1, InStr(1, Name, Chr$(0)) - 1)
                        FileListStart = FileListStart + Len(Name) + 1
                    
                        If Name = "" Or Offset = 0 Or Size = 0 Then
                            CyTExtract = False
                            Close FileNumber
                            Exit Function
                        ElseIf LCase(Name) = LCase(FileToExtract) Then
                            frmBusy.lblFile = "Extracting " & FileToExtract
                            DestinationNumber = FreeFile
                            Open DestinationFile For Binary As DestinationNumber
                                If Size > 100000 Then 'Divid the file in parts to use less memory and make less swap
                                    'BytesExtract = String(Size / 100, Chr$(0))
                                    'For Position = 1 To Size Step Len(BytesExtract)
                                        'Get FileNumber, Position + OffSet, BytesExtract
                                        'Put DestinationNumber, Position, BytesExtract
                                    'Next Position
                                    
                                    Position = -1000000
                                    frmBusy.prgFile.Max = Size
                                    Do
                                        
                                        Position = Position + 1000000
                                        If Position + 999999 > Size Then
                                            BytesExtract = String(Size - Position, Chr$(0))
                                            frmBusy.prgFile.Value = frmBusy.prgFile.Max
                                            frmBusy.Refresh
                                        Else
                                            BytesExtract = String(1000000, Chr$(0))
                                            frmBusy.prgFile.Value = Position
                                            frmBusy.Refresh
                                        End If
                                        Get FileNumber, Position + Offset, BytesExtract
                                        Put DestinationNumber, Position + 1, BytesExtract
                                    Loop Until Position + 999999 >= Size
                                Else
                                    BytesExtract = String(Size, Chr$(0))
                                    Get FileNumber, Offset, BytesExtract
                                    Put DestinationNumber, 1, BytesExtract
                                End If
                            Close DestinationNumber
                            Close FileNumber
                            CyTExtract = True
                            Exit Function
                        End If
                    Loop Until FileListStart > LOF(FileNumber)
                End If
            Close FileNumber
            CyTExtract = False
        Else
            CyTExtract = False
            Exit Function
        End If
    End If
    Exit Function
FinaliseError:
    'MessageBox "An error occured while trying to extract file(s).", OKOnly, Critical
End Function

Function RemoveBackSlash(FileName As String) As String
    Dim Temp As Integer
    
    Do
        Temp = Slash
        Slash = InStr(Slash + 1, FileName, "\")
        If Slash = 0 Then
            RemoveBackSlash = Mid(FileName, Temp + 1)
            Exit Function
        End If
    Loop
End Function

Private Sub OpenPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenPic.Picture = ImageListGray.ListImages(2).Picture
End Sub

Private Sub OptionsPic_Click()
    FrmOptions.Show 1, Me
End Sub

Private Sub OptionsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OptionsPic.Picture = ImageListGray.ListImages.Item(11).Picture
End Sub
