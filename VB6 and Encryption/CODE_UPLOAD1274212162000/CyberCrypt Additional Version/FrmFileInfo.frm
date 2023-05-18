VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFileInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File information"
   ClientHeight    =   6000
   ClientLeft      =   4215
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "FrmFileInfo.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.PictureBox StrMnu 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   5295
      TabIndex        =   19
      Top             =   720
      Width           =   5295
      Begin VB.Frame Frame3 
         Caption         =   "General"
         Height          =   4095
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   5055
         Begin VB.TextBox AppLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   3600
            Width           =   3015
         End
         Begin VB.TextBox ArchiveSize 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   3240
            Width           =   3015
         End
         Begin VB.TextBox ArchNum 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2880
            Width           =   3255
         End
         Begin VB.TextBox FileType 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox Info02a 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox Info05 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2160
            Width           =   3255
         End
         Begin VB.TextBox Info04 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1800
            Width           =   3255
         End
         Begin VB.TextBox Info03 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1440
            Width           =   3255
         End
         Begin VB.TextBox Info02 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox Info01 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application path:"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   3600
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Archive size KB / MB:"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   3240
            Width           =   1560
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sel file number:"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   2880
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File type:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   2520
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File size Kb / Mb:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset:"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label lblFileExtract0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File extracted from:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblPath0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File path:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   645
         End
         Begin VB.Label lblName0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File name:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   720
         End
         Begin VB.Label LblSize0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File size:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Icon and picture properties"
      Height          =   4215
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
      Begin VB.PictureBox PreviewImage 
         AutoRedraw      =   -1  'True
         Height          =   1695
         Left            =   3000
         ScaleHeight     =   1635
         ScaleWidth      =   1635
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Picture properties"
         Height          =   2175
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   4575
         Begin VB.TextBox PicWidth 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox PicHeight 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox ImageType 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width:"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hight:"
            Height          =   195
            Left            =   720
            TabIndex        =   29
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File type:"
            Height          =   195
            Left            =   480
            TabIndex        =   28
            Top             =   1320
            Width           =   630
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Preview"
         Height          =   1095
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   855
         Begin VB.PictureBox pctIcon 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   540
            Left            =   120
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmFileInfo.frx":030A
         Height          =   1335
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   3495
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.ToolTipText     =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other..."
            Object.ToolTipText     =   "Other..."
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In form load it checks if the Temp folder in the Windows directory
'exists and if the file exists in the Temp folder
Private Sub Form_Load()
    If FileExist(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem) = True Then
        On Error GoTo FinaliseError
        lngIcon = ExtractIcon(App.hInstance, (SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem), 0)
        If lngIcon = 0 Then
            dFileName = SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem
            GetFileInfo
            Exit Sub
        Else
            dFileName = SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem
            GetFileInfo
            pctIcon.Cls
            pctIcon.AutoSize = True
            pctIcon.AutoRedraw = True
            DrawIcon pctIcon.hdc, 0, 0, lngIcon
            pctIcon.Refresh
            DestroyIcon lngIcon
            End If
        Exit Sub
FinaliseError:
        MessageBox "Sorry the file info could not be found.", OKOnly, Critical
        Unload Me
        KillFile
            Else
        MessageBox "Sorry the file info could not be found.", OKOnly, Critical
        Unload Me
    End If
End Sub

Private Sub OK_Click()
    Unload Me
    KillFile
End Sub

'Kills the file after loading the data
Private Sub KillFile()
    On Error Resume Next
    Kill SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem
End Sub

'This function gets the file information and other information like
'the Offset and file type
Private Function GetFileInfo()
    
    On Error GoTo FinaliseError
    
    Me.Caption = "File information (" & frmMain.ListFiles.SelectedItem.Text & ")"
    
    Info01.Text = frmMain.ListFiles.SelectedItem
    Info02.Text = FileLen(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem)
    Info02a.Text = FormatKB(FileLen(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem))
    Info03.Text = SystemRootS & "\Temp\"
    Info04.Text = CyTFile
    Info05.Text = frmMain.ListFiles.SelectedItem.ListSubItems(3).Text
    ArchNum.Text = frmMain.ListFiles.SelectedItem.Index
    ArchiveSize.Text = FormatKB(FileLen(CyTFile))
    AppLocation.Text = App.Path
    
    NameTemp = Right(frmMain.ListFiles.SelectedItem.Text, 3)
    
    'Gets the file type and puts it into the correct text box
    If GetFileTypeName(NameTemp & "file") = LCase$("<Unknown file type>") Or Len(GetFileTypeName(NameTemp & "file")) >= 88 Then
        If GetFileTypeName("." & NameTemp) = LCase$("<Unknown file type>") Or Len(GetFileTypeName("." & NameTemp)) >= 88 Then
            FileType.Text = UCase$(NameTemp) & " file"
                Else
            FileType.Text = GetFileTypeName("." & NameTemp)
        End If
            Else
        FileType.Text = GetFileTypeName("." & NameTemp)
    End If
    
    PreviewImage.AutoRedraw = True
    ImageType.Text = Right(frmMain.ListFiles.SelectedItem.Text, 3)
    CentrePic PreviewImage, LoadPicture(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem.Text)
    PicWidth.Text = LoadPicture(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem.Text).Width
    PicHeight.Text = LoadPicture(SystemRootS & "\Temp\" & frmMain.ListFiles.SelectedItem.Text).Height
    
FinaliseError:
    
    KillFile
        
End Function

Private Sub chkAttrib_GotFocus(Index As Integer)
    OK.SetFocus
End Sub

'This sub selects through the basic tabbed menu
Private Sub TabStrip_Click()
    If TabStrip.SelectedItem = "General" Then
        Frame5.Visible = False
        StrMnu.Visible = True
            ElseIf TabStrip.SelectedItem = "Other..." Then
        StrMnu.Visible = False
        Frame5.Visible = True
    End If
End Sub
