VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Win32 RegDemo"
   ClientHeight    =   6630
   ClientLeft      =   1110
   ClientTop       =   2115
   ClientWidth     =   7590
   Icon            =   "REGDEMO.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6630
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   2445
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "REGDEMO.frx":030A
      Top             =   3330
      Width           =   5565
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   4740
      TabIndex        =   5
      Top             =   5640
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   3672
      Left            =   3360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3675
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   300
      Width           =   60
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   288
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "My Computer"
      Top             =   4560
      Width           =   7272
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4152
      Left            =   3480
      TabIndex        =   2
      Top             =   300
      Width           =   3132
      _ExtentX        =   5530
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   4152
      Left            =   180
      TabIndex        =   1
      Top             =   300
      Width           =   3192
      _ExtentX        =   5636
      _ExtentY        =   7329
      _Version        =   327682
      Indentation     =   2
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   3960
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":0310
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":0422
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":0534
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":0960
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "REGDEMO.frx":0A72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Key Name"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1272
   End
   Begin VB.Menu MunRegOptions 
      Caption         =   "Reg Options"
      Begin VB.Menu MnuCollapse 
         Caption         =   "&Collapse"
      End
      Begin VB.Menu MnuNew 
         Caption         =   "&New"
         Begin VB.Menu MnuKey 
            Caption         =   "&Key"
         End
         Begin VB.Menu mnusep1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuStringValue 
            Caption         =   "&String Value"
         End
         Begin VB.Menu MnuBinaryValue 
            Caption         =   "&Binary Value"
         End
         Begin VB.Menu MnuDword 
            Caption         =   "&DWORD value"
         End
      End
      Begin VB.Menu MnuFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu MnuRenam 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopyKeyName 
         Caption         =   "&Copy Key Name"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OsVers As OsVersionInfo


Private Sub ResizeControls()
    '-----------------------------------------------------
    'Resize controls when the user resizes the form
    'or moves the splitter bar.
    '-----------------------------------------------------
    On Error GoTo ReSizeError
    TreeView1.Move 5 * Screen.TwipsPerPixelX, TreeView1.Top, Picture1.Left - 5 * Screen.TwipsPerPixelX, ScaleHeight - TreeView1.Top - Text1.Height - 5 * Screen.TwipsPerPixelY
    ListView1.Move Picture1.Left + Picture1.Width, TreeView1.Top, ScaleWidth - Picture1.Left - Picture1.Width - 5 * Screen.TwipsPerPixelX, TreeView1.Height
    Picture1.Top = TreeView1.Top
    Picture1.Height = TreeView1.Height
    
    Label1.Left = TreeView1.Left + 60
    
    Text1.Move TreeView1.Left, ScaleHeight - Text1.Height, ScaleWidth - 2 * TreeView1.Left
    Exit Sub
ReSizeError:
    Exit Sub

End Sub


Private Sub Command1_Click()
'Call SaveListBox(App.Path & "\urls.reg", ListView1)
Dim po As Integer
Dim asr As Integer

For po = 0 To TreeView1.SelectedItem.Children - 1
asr = TreeView1.Nodes(TreeView1.SelectedItem.Index + po).Selected = True
Text2.Text = Text2.Text & vbNewLine & TreeView1.Nodes((TreeView1.SelectedItem.Index))
Next po

End Sub

Public Sub SaveListBox(Directory As String, TheList As ListView)
    Dim savelist As Long
    'On Error Resume Next
    Open Directory$ For Output As #1
    Print #1, "REGEDIT4" & vbNewLine & vbNewLine & "[" & Text1.Text & "]" & vbNewLine
    For savelist& = 1 To TheList.ListItems.Count
        Print #1, """" & TheList.ListItems(savelist&).Text & """="; TheList.ListItems(savelist&).SubItems(1)
    Next savelist&
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    
    While Not EOF(1)
        Input #1, MyString$
        If LOF(1) = 0 Then Exit Sub
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    If Source = Picture1 Then
        Picture1.Left = x
        ResizeControls
    End If

End Sub

Private Sub Form_Load()
    Width = 600 * Screen.TwipsPerPixelX
    Picture1.Width = 5 * Screen.TwipsPerPixelX
    Picture1.Left = 250 * Screen.TwipsPerPixelX
    Dim nodX As Node
    Dim ClmHdr As ColumnHeader
    
    '-----------------------------------------------------
    ' add columns to listview control
    '-----------------------------------------------------
    Set ClmHdr = ListView1.ColumnHeaders.Add()
    ClmHdr.Text = "Name"
    ClmHdr.Width = ListView1.Width / 3
    Set ClmHdr = Me.ListView1.ColumnHeaders.Add()
    ClmHdr.Text = "Data"
    ClmHdr.Width = 3 * ListView1.Width / 2
    '-----------------------------------------------------
    'First we find out what Windows is running. There are a
    'couple of registry keys in Win95 that are not present
    'in Windows NT
    '-----------------------------------------------------
    OsVers.dwVersionInfoSize = 148&
    lTempLong = GetVersionEx(OsVers)
    Select Case OsVers.dwPlatform
        Case VER_PLATFORM_WIN32_NT
            iWinVers = WinNt
        Case VER_PLATFORM_WIN32_WINDOWS
            iWinVers = Win32
        Case Else 'Shouldn't happen
            MsgBox "This program is intended only for use with 32-bit Windows versions."
            Unload Form1
    End Select
       
    '-----------------------------------------------------
    'Start the TreeView with a toplevel key
    '-----------------------------------------------------
    
    Set nodX = TreeView1.Nodes.Add(, , "main", "My Computer", 3)
    
    '-----------------------------------------------------
    'Put in Public main keys
    '-----------------------------------------------------
    
    Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT", 1)
    nodX.EnsureVisible 'Forces the tree open to this level
    nodX.ExpandedImage = 2
    nodX.SelectedImage = 2
    RegEnumKeys nodX, False 'Enumerate a single rgeSubKey, to put a + on the key
    Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_CURRENT_USER", "HKEY_CURRENT_USER", 1)
    nodX.ExpandedImage = 2
    nodX.SelectedImage = 2
    RegEnumKeys nodX, False
    Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE", 1)
    nodX.ExpandedImage = 2
    nodX.SelectedImage = 2
    RegEnumKeys nodX, False
    Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_USERS", "HKEY_USERS", 1)
    nodX.ExpandedImage = 2
    nodX.SelectedImage = 2
    RegEnumKeys nodX, False
    
    '-----------------------------------------------------
    'Put in version-specific main keys
    'HKEY_PERFORMANCE_DATA does not respond in a normal fashion
    'to key queries.  Note it does not appear in REGEDT32.
    'If iWinVers = WinNT Then
        'treeview1.AddItem "HKEY_PERFORMANCE_DATA"
    'End If
    '-----------------------------------------------------
    If iWinVers = Win32 Then
        Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG", 1)
        nodX.ExpandedImage = 2
        nodX.SelectedImage = 2
        RegEnumKeys nodX, False
        Set nodX = TreeView1.Nodes.Add("main", tvwChild, "HKEY_DYN_DATA", "HKEY_DYN_DATA", 1)
        nodX.ExpandedImage = 2
        nodX.SelectedImage = 2
        RegEnumKeys nodX, False
    End If
    '-----------------------------------------------------
    'Setup hourglass cursor
    '-----------------------------------------------------
    iWaitCursor = LoadCursor(0, IDC_WAIT)

End Sub

Private Sub Form_Resize()
    ResizeControls
End Sub


Private Sub listview1_DblClick()
    ' if nothing selected get out
    If Not (ListView1.SelectedItem Is Nothing) Then
        '----------------------------------------------------------
        'Call EditRegValue to load the value into the editor.
        '----------------------------------------------------------
        EditRegValue TreeView1.Nodes(TreeView1.SelectedItem.Key), CLng(ListView1.ListItems(ListView1.SelectedItem.Index).Tag)
    End If
    

End Sub

Private Sub ListView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = Picture1 Then
        Picture1.Left = x + ListView1.Left
        ResizeControls
    End If

End Sub


Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
    Text1 = truncatepath(Node.FullPath, Me, Text1)

End Sub

Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = Picture1 Then
        Picture1.Left = x + TreeView1.Left
        ResizeControls
    End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As Node)
If Node.Children > 1 Then Exit Sub
If Node.Text = "My Computer" Then Exit Sub
    RegEnumKeys Node, True
    Node.Sorted = True

End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
'MsgBox TreeView1.SelectedItem.FirstSibling ' shows current selected name
'MsgBox TreeView1.SelectedItem.Child ' shows next one under
'MsgBox TreeView1.SelectedItem.LastSibling ' shows last top most in list
'MsgBox TreeView1.SelectedItem.Root ' returns top most
If LenB(TreeView1.SelectedItem.Child) = 0 Then
MnuCollapse.Enabled = True
End If

If TreeView1.SelectedItem.Index = 2 Or TreeView1.SelectedItem.Index = 4 Or TreeView1.SelectedItem.Index = 6 Or TreeView1.SelectedItem.Index = 8 Or TreeView1.SelectedItem.Index = 10 Or TreeView1.SelectedItem.Index = 12 Then
MnuDelete.Enabled = False
MnuRenam.Enabled = False
Else
MnuDelete.Enabled = True
MnuRenam.Enabled = True
End If

PopupMenu MunRegOptions
End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)

    '-----------------------------------------------------
    'Empty the value list.
    '-----------------------------------------------------
    ListView1.ListItems.Clear
    
    '-----------------------------------------------------
    'Remember which line was clicked when we start
    '   moving up and down the list.
    '-----------------------------------------------------
    
    rgeSubKey = Node.FullPath
    Text1 = truncatepath(Node.FullPath, Me, Text1)
    If rgeSubKey = "My Computer" Then Exit Sub
    '-----------------------------------------------------
    'Now enumerate all values belonging to this key
    '-----------------------------------------------------
    ListView1.Sorted = False
    RegEnumValues
    ListView1.Sorted = True
End Sub


