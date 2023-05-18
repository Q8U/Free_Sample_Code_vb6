VERSION 5.00
Begin VB.Form Viewer 
   Caption         =   "IconWorks Viewer"
   ClientHeight    =   4245
   ClientLeft      =   1455
   ClientTop       =   1875
   ClientWidth     =   5895
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "VIEWICON.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   Tag             =   "IconWrks Viewer"
   Begin VB.PictureBox Pic_SelectedIconLabel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   15
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   16
      Top             =   0
      Width           =   3540
      Begin VB.PictureBox Pic_IconsBitmap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   0
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox Pic_SelectedIcon 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3045
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.PictureBox Pic_VerticalLine 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   3555
      ScaleHeight     =   4245
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   0
      Width           =   15
   End
   Begin VB.PictureBox Pic_AllIcons 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   3570
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.VScrollBar Scrl_AllIcons 
      Height          =   4275
      Left            =   5640
      TabIndex        =   12
      Top             =   -15
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Txt_FileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   795
      TabIndex        =   1
      Top             =   525
      Width           =   2775
   End
   Begin VB.PictureBox Pic_IconCount 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1365
      ScaleHeight     =   210
      ScaleWidth      =   420
      TabIndex        =   13
      Top             =   1335
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.DirListBox Dir_DirectoryList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   -15
      TabIndex        =   3
      Top             =   1560
      Width           =   1800
   End
   Begin VB.FileListBox File_FileList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   1770
      Pattern         =   "*.ico"
      TabIndex        =   5
      Top             =   1560
      Width           =   1800
   End
   Begin VB.DriveListBox Drv_DriveList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   -15
      TabIndex        =   7
      Top             =   3945
      Width           =   3585
   End
   Begin VB.Line line_HorizontalLine 
      X1              =   0
      X2              =   235
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Label Lbl_File 
      Caption         =   "Fi&le:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   8
      Top             =   585
      Width           =   795
   End
   Begin VB.Label Lbl_Directory 
      Caption         =   "Directory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   14
      Top             =   870
      Width           =   1380
   End
   Begin VB.Label Lbl_CurrentDirectory 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   -15
      TabIndex        =   15
      Top             =   1080
      Width           =   3585
   End
   Begin VB.Label Lbl_Directories 
      Caption         =   "&Directories:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   1335
      Width           =   1365
   End
   Begin VB.Label Lbl_Icons 
      Alignment       =   2  'Center
      Caption         =   "&Icons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1785
      TabIndex        =   4
      Top             =   1335
      Width           =   1755
   End
   Begin VB.Label Lbl_Drives 
      Caption         =   "Dri&ves:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   6
      Top             =   3735
      Width           =   1365
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Menu_FileSelection 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Menu_EditCopy 
         Caption         =   "&Copy"
      End
   End
   Begin VB.Menu Menu_Options 
      Caption         =   "&Options"
      Begin VB.Menu Menu_OptionsSelection 
         Caption         =   "&Editor..."
         Index           =   0
         Shortcut        =   {F7}
      End
      Begin VB.Menu Menu_OptionsSelection 
         Caption         =   "&Show all Icons"
         Index           =   1
         Shortcut        =   ^V
      End
      Begin VB.Menu Menu_OptionsSelection 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Menu_OptionsSelection 
         Caption         =   "Show all icons on &DIR change"
         Index           =   3
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Index"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Keyboard"
         Index           =   2
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Commands"
         Index           =   3
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&Using Help"
         Index           =   4
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Menu_HelpSelection 
         Caption         =   "&About..."
         Index           =   6
      End
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



DefInt A-Z

Dim Badicon

Private Sub Adjust_All_Controls()
  
    ' Save the visibility state of the Icon viewing window, since
    ' we resize it whether it is visible or not.
    '
    AllIconsVisible = Pic_AllIcons.Visible

    ' Hide all controls that can be resized, while the actual resizing is
    ' being done.  This prevents uneccessary screen updates.
    '
    Pic_AllIcons.Visible = False
    Scrl_AllIcons.Visible = False
    File_FileList.Visible = False
    Dir_DirectoryList.Visible = False

    ' Calculate number of icon rows and columns for the new Window size,
    ' and the maximum number of icons that can be displayed at once
    ' within the new window size.
    '
    IconRows = ScaleHeight \ ICON_CELL
    IconColumns = (ScaleWidth - Pic_AllIcons.Left) \ ICON_CELL
    MaxIcons = IconColumns * IconRows
  
    ' Set new width for the icon viewing area
    '
    NewAllIconsWidth = ScaleWidth - Pic_AllIcons.Left + 1

    ' Check if there are more icons than can be displayed at once in the viewing window
    '
    If File_FileList.ListCount > MaxIcons Then
        '
        ' All the icons cannot be displayed at once, so the viewing area must be
        ' reset, the Viewing area Scrollbar must be reset, and the number of icon
        ' columns must be reset since the scrollbar now takes up part of the viewing area.
        '
        Scrl_AllIcons.Left = ScaleWidth - Scrl_AllIcons.Width + 1
        NewAllIconsWidth = Scrl_AllIcons.Left - Pic_AllIcons.Left
        IconColumns = NewAllIconsWidth \ ICON_CELL
        MaxIcons = IconColumns * IconRows
    End If
      
    ' Resize and Repostion affected controls
    '
    Pic_AllIcons.Move Pic_AllIcons.Left, Pic_AllIcons.Top, NewAllIconsWidth, ScaleHeight
    Scrl_AllIcons.Height = ScaleHeight + 2
    Pic_VerticalLine.Height = ScaleHeight
    Drv_DriveList.Top = ScaleHeight - Drv_DriveList.Height + 1
    Lbl_Drives.Top = Drv_DriveList.Top - Lbl_Drives.Height
    File_FileList.Height = Lbl_Drives.Top - Dir_DirectoryList.Top - 1
    Dir_DirectoryList.Height = File_FileList.Height

    ' Redisplay controls hidden before resizing and reposition was done
    '
    Pic_AllIcons.Visible = AllIconsVisible
    File_FileList.Visible = True
    Dir_DirectoryList.Visible = True

End Sub

Private Sub Dir_DirectoryList_Change()
  
    ' A new directory has been selected, so Set current directory
    ' to the newly selected directory
    '
    ChDir Dir_DirectoryList.Path
    
    ' Display the newly selected directory
    '
    Lbl_CurrentDirectory.Caption = Dir_DirectoryList.Path

    ' Inform the File ListBox of the PathChange.
    '
    File_FileList.Path = Dir_DirectoryList.Path

    ' Display new filespec in FileName TextBox
    '
    UpDate_FileSpec Viewer

End Sub

Private Sub Dir_DirectoryList_Click()
    
    ' The actual directory has not changed since the Directory ListBox was
    ' only single clicked , so all we need to do is display the new file
    ' spec for the selected directory in the FileName TextBox.
    '
    UpDate_FileSpec Viewer
    VLastChanged = DIR_CHANGED

End Sub

Private Sub Dir_DirectoryList_KeyPress(KeyAscii As Integer)

    ' Pressing Enter when the Directory ListBox has the Focus should
    ' react just as if the Directory ListBox was double clicked, so all we
    ' need to do is set the Path property of the Directory control to the
    ' selected directory.
    '
    If KeyAscii = 13 Then Dir_DirectoryList.Path = Dir_DirectoryList.List(Dir_DirectoryList.ListIndex)

End Sub

Private Sub Drv_DriveList_Change()

    ' Selecting a drive from a Drive control does not generate an error
    ' if the selected drive is not ready, so we verify that the drive is
    ' in fact ready before we accept the drive.
    '
    Validate_And_Change_Drives Viewer

End Sub

Private Sub File_FileList_Click()

    If File_FileList.ListIndex >= 0 Then
        ' When a file is selected from the File Listbox with single click
        ' from the mouse, this routine displays the selected icon just above
        ' the file listbox if it is a valid Icon file.
        '
        Txt_FileName.Text = File_FileList.FileName
        Badicon = Not Valid_Icon((File_FileList.FileName), True)
        If Not Badicon Then
            '
            ' File is valid Icon file
            '
            Menu_EditCopy.Enabled = True
            VLastChanged = FILE_CHANGED
        End If
    End If

End Sub

Private Sub File_FileList_DblClick()
  
    ' Double Clicking a file within the File ListBox signals that an
    ' existing file has been selected, so attempt to open the file.
    '
    If Not Badicon Then
        VLastChanged = FILE_CHANGED
        Open_Selected_Icon
    End If

End Sub

Private Sub File_FileList_KeyPress(KeyAscii As Integer)
    
    ' Pressing Enter when the File ListBox has the Focus should react
    ' just as if the File ListBox was Double Clicked, so all we need
    ' to do is attempt to open the selected file.
    '
    VLastChanged = FILE_CHANGED
    If KeyAscii = 13 Then Open_Selected_Icon

End Sub

Private Sub File_FileList_PathChange()
    
    ShowingAllIcons = False
    
    If (File_FileList.ListCount > 0) And Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked Then
        '
        ' There are icons in the new directory and the user has selected
        ' to automatically display all icons when the directory is changed,
        ' so we simulate selecting the menu item which displays all the icons.
        '
        Menu_OptionsSelection_Click MID_SHOW_ALL_ICONS
    Else
        ' There are no icons in the current directory, or the user does
        ' not want to automatically display all icons when the directory
        ' changes, so we need to get rid of all displayed icons.
        '
        Scrl_AllIcons.Visible = False
        Pic_AllIcons.Visible = False
        Pic_SelectedIcon.Picture = LoadPicture()
    End If

    ' The menu item to show all icons, is enabled if the current directory
    ' contains icons, and disabled if it does not.
    '
    Menu_OptionsSelection(MID_SHOW_ALL_ICONS).Enabled = File_FileList.ListCount > 0

    ' Display the number of icons in the current directory
    '
    Lbl_Icons.Caption = Format$(File_FileList.ListCount) + " &Icons"

End Sub

Private Sub Form_Load()

    Pic_IconsBitmap.Move 0, 0, 34, 34
    Pic_SelectedIcon.Move Pic_SelectedIcon.Left, Pic_SelectedIcon.Top, 32, 32
    Pic_SelectedIcon.BackColor = WHITE
    Pic_IconsBitmap.BackColor = WHITE

    ' Inform rest of Iconworks that the Viewer is loaded.  Viewer.Visible could
    ' be tested but accessing the visible property would cause the Viewer to be
    ' loaded if not already loaded.
    '
    ViewerLoaded = True

    Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked = -GetPrivateProfileInt(APP_NAME, KEY_SHOW_ICONS, 0, INI_FILENAME)

    ' Position Viewer at 0,0, and set Width and Height to 2/3's that of the Screen.
    '
    Move 0, 0, Screen.Width * 0.66, Screen.Height * 0.66

    ' Calculate the Minimum width and Height for the Viewer.  This is done, so
    ' the smallest window allowed will still allow easy access to all controls.
    '
    MinViewerWidth = (Pic_VerticalLine.Left + Scrl_AllIcons.Width + ICON_CELL + 2) * 15 + (Width - ScaleWidth * 15)
    MinViewerHeight = ICON_CELL * 6 * 15 + (Height - ScaleHeight * 15)

    ' Enable the "Show all icons' menu option only if the current directory
    ' contains icons
    '
    Menu_OptionsSelection(MID_SHOW_ALL_ICONS).Enabled = File_FileList.ListCount > 0

    ' Display the number of icons in the current directory, display the
    ' current directory, and set the current file name to the default
    ' file spec of "*.ICO", which was set at design time into the File
    ' ListBox.
    '
    Lbl_Icons.Caption = Format$(File_FileList.ListCount) + " &Icons"
    Lbl_CurrentDirectory.Caption = Dir_DirectoryList.Path
    Txt_FileName.Text = File_FileList.Pattern

    VLastChanged = DIR_CHANGED
    
    ' The Alt+F4 accelerator for Exit, cannot be assigned using the Menu
    ' design Window, so we need to put the accelerator into the caption.
    ' Alt+F4 is actually the System menus Close option.
    '
    Menu_FileSelection(MID_EXIT).Caption = "E&xit" + A_TAB + "Alt+F4"
                                        
    Show
    Refresh
    If Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked And (File_FileList.ListCount <> 0) Then Menu_OptionsSelection_Click MID_SHOW_ALL_ICONS

End Sub

Private Sub Form_Resize()
  
    ' The Form has been resized, so we need to resize and possible reposition
    ' some of the controls on the form, however, we do not want to do anything
    ' if the form is minimized.
    '
    If WindowState <> MINIMIZED Then
        '
        ' Check if new size is less than the minimum Viewer size.
        '
        If (Width < MinViewerWidth) Or (Height < MinViewerHeight) Then
            '
            ' The form is smaller than the minimum size, either in width or
            ' height, so reset the width and/or height to the minimum values.
            '
            If Width < MinViewerWidth Then NewWidth = MinViewerWidth Else NewWidth = Width
            If Height < MinViewerHeight Then NewHeight = MinViewerHeight Else NewHeight = Height
            Move Left, Top, NewWidth, NewHeight
        Else
            ' The form is greater than the minimum width and height values
            ' so adjust any controls that need resizing or repositioning.
            '
            Adjust_All_Controls

            Scrl_AllIcons.Value = 0

            If (File_FileList.ListCount > MaxIcons) And ShowingAllIcons Then
                '
                ' There are more icons that can be displayed at once within
                ' the current new size of the Form, so we need to calculate
                ' new Max and LargeChange values for the scrollbar.
                '
                Diff = File_FileList.ListCount - MaxIcons
                Scrl_AllIcons.Max = Diff \ IconColumns
                If (Diff Mod IconColumns) Then Scrl_AllIcons.Max = Scrl_AllIcons.Max + 1
                Scrl_AllIcons.LargeChange = IconRows
                Scrl_AllIcons.Visible = True ' And (File_FileList.ListCount > MaxIcons)
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Let the rest of IconWorks know that the Viewer is no longer loaded.
    '
    ViewerLoaded = False
    ShowingAllIcons = False
    R = WinHelp(hWnd, dummy$, HELP_QUIT, 0)

    R = WritePrivateProfileString(APP_NAME, KEY_SHOW_ICONS, Format$(Abs(Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked)), INI_FILENAME)

    ' If the Viewer was started up first then we treat it as the main Form.
    ' So, if the Editor is loaded, we should as the user if the Editor should
    ' also be terminated.
    '
    If (MainForm = ICONWORKS_VIEWER) And (EditorLoaded) Then
        '
        ' Viewer was started first and the Editor is loaded so ask the user
        ' if the Editor should also be terminated.
        '
        Text = "Terminate Editor Also?"
        If MsgBox(Text, 36, "IconWorks") = MBYES Then Unload Editor
        MainForm = ICONWORKS_EDITOR
    End If

End Sub

' When a request is made to display all the icons in the current directory
' this routine is called to perform the task.
'
Private Sub Load_All_Icons()
    
    If CurDir$ <> File_FileList.Path Then ChDir File_FileList.Path

    ' Refresh the File listbox to pick up any files that might have been
    ' added to the current directory since this directory was selected.
    '
    File_FileList.Refresh

    ' Determine if the scrollbar is needed.  If there are more icons in the
    ' current directory than can be displayed at once, the scrollbar must
    ' be active to allow viewing of all the icons.
    '
    Scrl_AllIcons.Visible = File_FileList.ListCount > MaxIcons
    Scrl_AllIcons.Value = 0

    ' Display the Icon Viewing window
    '
    Pic_AllIcons.Visible = True
     
    ' When all the icons are displayed, a single bitmap is created and maintained
    ' in memory.  This bitmap contains the images of all the icons in the current
    ' directory.  This bitmap is used to update the Viewing Window when the window
    ' is scrolled or when the Form is resized.  The bitmap is made of the image
    ' of each icon concatenated into one long bitmap.  This makes for a very fast
    ' screen update when the icons need to be redisplayed.  The icons do not have
    ' to be reloaded each time, but simply copied from this bitmap to the viewing
    ' window.
    '
    Pic_IconsBitmap.Width = File_FileList.ListCount * ICON_CELL
    Pic_IconsBitmap.Cls

    ' To build the memory Icon Bitmap above, each icon must be loaded at least
    ' once so as to obtain its image and add it to the memory bitmap.  The
    ' Pic_SelectedIcon picture contol is used for this purpose, but to prevent
    ' uneccessary flashing of this picture control as each icon is loaded,
    ' it is hidden while this is going on.  And since it is hidden while this
    ' is occuring, AutoRedraw must be set to TRUE to allow copying of the image
    ' while it is hidden.  The image is copied using the Windows API routine
    ' BitBlt().
    '
    Pic_SelectedIcon.Visible = False
    Pic_SelectedIcon.AutoRedraw = True

    ' So something is visibly happening while the icons are being loaded and the
    ' bitmap is being created, the File Listbox label's color is changed, and
    ' the Lbl_IconCount is made visible.  These labels count and display the number
    ' of icons loaded as they are being loaded.
    '
    Lbl_Icons.Caption = "Icons Loaded"
    Lbl_Icons.ForeColor = WHITE
    Lbl_Icons.BackColor = RED
    Lbl_Icons.Refresh
    Pic_IconCount.Visible = True
    
    ' Attempt to load all files listed in the File ListBox.  If valid Icon files
    ' add image to memory bitmap.
    '
    Screen.MousePointer = HOURGLASS
    For X = 0 To File_FileList.ListCount - 1
        '
        ' Display current count of Icons loaded
        '
        Pic_IconCount.CurrentX = 0
        Pic_IconCount.Print X + 1; "  ";
        If Valid_Icon((File_FileList.List(X)), False) Then
            '
            ' The file was a valid Icon file, so add its image to the memory Bitamp
            '
            R = BitBlt(Pic_IconsBitmap.hDC, 2 + X * ICON_CELL, 0, 32, 32, Pic_SelectedIcon.hDC, 0, 0, SRCCOPY)
        Else
            ' The file was not a valid Icon file, so display a BLACK square where the icons
            ' image would have been placed within the Memory bitmap.
            '
            R = BitBlt(Pic_IconsBitmap.hDC, 2 + X * ICON_CELL, 0, 32, 32, 0, 0, 0, BLACKNESS)
        End If
    Next X
    Screen.MousePointer = DEFAULT

    ' Re-Display the SelectedIcon picture and disable its AutoRedraw since
    ' it is no longer needed.
    '
    Pic_SelectedIcon.Visible = True
    Pic_SelectedIcon.AutoRedraw = False

    ' Reset the File list Labels to normal, and display the total number of Icons loaded.
    '
    Lbl_Icons.ForeColor = WINDOW_TEXT
    Lbl_Icons.BackColor = WINDOW_BACKGROUND
    Lbl_Icons.Caption = Format$(File_FileList.ListCount) + " &Icons"

    ' Hide the IconCount label since it is not needed except while loading icons.
    '
    Pic_IconCount.Visible = False

End Sub

Private Sub Menu_EditCopy_Click()
  
    ' Can't place an actual Icon into the System clipboard, so place
    ' a bitmap of its image, in response to a copy command.
    '
    Clipboard.Clear
    Clipboard.SetData Pic_SelectedIcon.Image
  
End Sub

Private Sub Menu_File_Click()

    ' Before displaying the file menu, enable or disable the File.Open
    ' command, based on whether or not an Icon is currently selected.
    '
    Menu_FileSelection(MID_OPEN).Enabled = File_FileList.ListIndex >= 0

End Sub

Private Sub Menu_FileSelection_Click(Index As Integer)

    ' One of the 2 File menu items were selected, so determine which one
    ' and perform the corresponding task.
    '
    Select Case Index
        
        Case MID_OPEN
            Open_Selected_Icon
            
        Case MID_EXIT
            Unload Viewer

    End Select

End Sub

Private Sub Menu_HelpSelection_Click(Index As Integer)

    If Index < MID_ABOUT Then
        '
        ' Determine what help topic to display.  The *Index* and *Using Help*
        ' items are the same for both the Viewer and the Editor, but the
        ' items: Keyboard and Commands are different and have
        ' different Help topic ID's, so we add 3 to the Menu item which
        ' will then make the Index correspond to the correct Help topic.
        '
        If (Index >= MID_KEYBOARD) And (Index <= MID_COMMANDS) Then Index = Index + 3
        Get_Help Index
    Else
        ' Display the IconWorks About box
        '
        AboutBox.Show MODAL
    End If

End Sub

Private Sub Menu_OptionsSelection_Click(Index As Integer)
    
    ' One of the 3 Options menu items were selected, so determine which one
    ' and perform the corresponding task.
    '
    Select Case Index
        
        Case MID_EDITOR
            ' Invoke the Editor, but do not open the selected Icon
            '
            Editor.Show MODELESS

        Case MID_SHOW_ALL_ICONS
            ' Check for too man Icons to display.  maximum of 963.
            '
            If File_FileList.ListCount > 900 Then
                MsgBox "Can display upto a maximum of 900 icons", 16, "Too many Icons"
            Else
                '
                ' Let the rest of the Viewer know that all the Icons are currently
                ' begin displayed.
                '
                Temp = ShowingAllIcons
                ShowingAllIcons = True
    
                ' Before showing all the icons, the values for the scrollbar must
                ' be re-calculated based on the number of icons in the current
                ' directory.  Since this is done when the form is resized, we can
                ' accomplish this by calling the Form_Resize event to do this for us.
                '
                If Not Temp Then Form_Resize
                 
                ' We disable the Edit.Copy menu Item, since after all icons are
                ' displayed, no one icon will be selected yet.
                '
                Menu_EditCopy.Enabled = False
    
                ' Load all the icons and then display them
                '
                Load_All_Icons
                Update_Displayed_Icons
            End If

        Case MID_SHOW_ON_DIR_CHANGE
            '
            ' Toggle the Checked property of the *Show all icons* options.
            '
            Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked = Not Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked
            If Menu_OptionsSelection(MID_SHOW_ON_DIR_CHANGE).Checked And Menu_OptionsSelection(MID_SHOW_ALL_ICONS).Enabled Then Menu_OptionsSelection_Click MID_SHOW_ALL_ICONS
    
    End Select

End Sub

Private Sub Open_Selected_Icon()
Dim OldPattern As String

    If VLastChanged = DIR_CHANGED Then
        '
        ' The directory was the last control accessed, so we need only
        ' set its Path to is currently selected item, which will generate
        ' a Change event for the Directory control, which will take care
        ' of updating the other related controls
        '
        Dir_DirectoryList.Path = Dir_DirectoryList.List(Dir_DirectoryList.ListIndex)
    Else
        ' The FileName TextBox or the File ListBox was last accessed.
        '
        ValidName = True

        ' Validate the filename only if the FileName TextBox was the last
        ' control accessed.  We do not need to Validate the Filename if the
        ' File ListBox was last accessed since if the FileName is listed,
        ' then the File exists.
        '
        If VLastChanged = FILENAME_CHANGED Then ValidName = Validate_FileSpec(Viewer, True)
      
        ' The FileName entered into the FileName TextBox many have contained
        ' a new drive and path, so in case it did, we need to inform the
        ' Drive and Directory controls of this change.
        '
        ChDir File_FileList.Path
        Drv_DriveList.Drive = Left$(File_FileList.Path, 2)
        Dir_DirectoryList.Path = File_FileList.Path

        If ValidName Then
            If Valid_Icon((Txt_FileName.Text), True) Then
                '
                ' A file has been selected so invoke the Editor, and load the
                ' icon into the editor.
                '
                Editor.Show MODELESS
                Load_An_Icon
            End If
        End If
    End If

End Sub

Private Sub Pic_AllIcons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Selections are made only with the Left mouse button.
    '
    If Button = LEFT_BUTTON Then
        '
        ' We need to determine if an icon was has actually been selected
        ' by the MouseDown event, since an Icon does not exist at every
        ' location within the Viewing area (Pic_AllIcons).  So, we
        ' need to calculate the icon position selected based on the
        ' mouse coordinates and then check if an Icon exists at that location.
    
        ' Calculate the column of the selected Icon position.
        '
        XIcon = X \ ICON_CELL
        
        ' Determine if any icons exist in that column.
        '
        If XIcon < IconColumns Then
            '
            ' A valid Column has been selected, so we now need to calculate
            ' the selected Row position.  The Scrollbar's value must be
            ' considered when calculating the Row.
            '
            YIcon = Y \ ICON_CELL + Scrl_AllIcons.Value

            ' Using the Column and Row selected, calculate the actual
            ' Icon position selected.
            '
            SelectedIcon = (YIcon * IconColumns) + XIcon

            ' Determine if an Icon exists at the selected location
            '
            If SelectedIcon < File_FileList.ListCount Then
                '
                ' An icon has been selected, so select the Icon in the File ListBox
                '
                File_FileList.ListIndex = -1
                File_FileList.ListIndex = SelectedIcon

                ' If icon is a valid Win 3.0 icon, begin dragging.
                '
                If Not Badicon Then
                    Pic_AllIcons.DragIcon = Pic_SelectedIcon.DragIcon
                    Pic_AllIcons.Drag
                End If
            End If
        End If
    End If
End Sub

Private Sub Pic_AllIcons_Paint()

    ' A portion of the viewing area needs to be updated, so if we
    ' are currently displaying any icons, Update the viewing area.
    '
    If ShowingAllIcons And (File_FileList.ListCount > 0) Then Update_Displayed_Icons

End Sub

Private Sub Pic_SelectedIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If (Button = LEFT_BUTTON) And (File_FileList.ListIndex >= 0) Then
        '
        ' Set the DragIcon to the Selected Icon so we see the actual icon
        ' when dragging, instead of an inverted Frame of the picture control,
        ' and begin dragging the icon.
        '
        'Pic_SelectedIcon.DragIcon = Pic_SelectedIcon.Picture
        Pic_SelectedIcon.Drag
    End If

End Sub

Private Sub Pic_SelectedIconLabel_Paint()

    Text = "Selected Icon:"
    Pic_SelectedIconLabel.CurrentX = (Pic_SelectedIconLabel.ScaleWidth - Pic_SelectedIconLabel.TextWidth(Text)) \ 2
    Pic_SelectedIconLabel.CurrentY = (Pic_SelectedIconLabel.ScaleHeight - Pic_SelectedIconLabel.TextHeight(Text)) \ 2
    Pic_SelectedIconLabel.Print Text

End Sub

Private Sub Scrl_AllIcons_Change()

    ' The Scrollbar was scrolled, so we need to scroll the displayed
    ' icons within the viewing window.  The Update_Displayed_Icons
    ' procedure displays the Icons based on the Value of the scrollbar
    ' if the scrollbar is currently visible.
    '
    Update_Displayed_Icons

End Sub

Private Sub Txt_FileName_Change()
  
    VLastChanged = FILENAME_CHANGED

End Sub

Private Sub Txt_FileName_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
        '
        ' Enter was pressed, so cancel the KeyStroke to prevent a Beep,
        ' and attempt to open the selected file as an Icon.
        '
        KeyAscii = 0
        Open_Selected_Icon
    End If

End Sub

Private Sub Update_Displayed_Icons()
    
    ' When the form is resized, the scrollbar is scrolled, or anything causing the
    ' currently displayed icons to be updated, this routine is called to display
    ' or redisplay the icons in the viewing window.
    '
    ' Clear the viewing window to White.  The .Cls method could be used, but
    ' it causes excessive flashing, so the .Line method is used instead to
    ' draw a filled white box inside the viewing window, which accomplishes
    ' the same thing but a little more efficiently.
    '
    Pic_AllIcons.Line (0, (IconRows - 1) * ICON_CELL)-(Pic_AllIcons.Width, Pic_AllIcons.Height), WHITE, BF
    
    ' Calculate the number of icon rows that need to be displayed.  It could
    ' be all the rows or only a few if all the icons can fit in the current size
    ' of the viewing window.
    '
    NumIconRows = IconRows
    If MaxIcons > File_FileList.ListCount Then NumIconRows = File_FileList.ListCount \ IconColumns

    ' Determine what icon should be the first icon displayed (Upper left hand
    ' corner of viewing window) based on the current value of the Scrollbar.
    '
    FirstIcon = Scrl_AllIcons.Value * IconColumns

    ' An entire row of Icons is displayed at once which is copied from the memory
    ' bitmap of the icon images.  So we need to calculate the width in pixels
    ' of the current with of a row of Icons, since this can change whenever the
    ' form is resized.
    '
    PixelWidth = IconColumns * ICON_CELL
    xSrc = FirstIcon * ICON_CELL
    Y = 1

    ' Copy icons from the memory Bitmap one row at a time to the viewing window
    '
    For row = 0 To NumIconRows
        R = BitBlt(Pic_AllIcons.hDC, 0, Y, PixelWidth, ICON_CELL, Pic_IconsBitmap.hDC, xSrc, 0, SRCCOPY)
        xSrc = xSrc + PixelWidth
        Y = Y + ICON_CELL
    Next row
 
End Sub

Private Function Valid_Icon(FileName As String, Prompt)
    
    On Error Resume Next

    ' Set Err to no Error (FALSE) and attempt to load the selected file
    '
    Err = False
    Pic_SelectedIcon.DragIcon = LoadPicture(FileName)
    If Err And Prompt Then
        '
        ' The file is not a valid Icon file
        '
        Beep
        X = MsgBox(FileName + " is not a valid Win 3.0 .ICO file", 16, "Bad File")
        Pic_SelectedIcon.Picture = LoadPicture()
        Menu_EditCopy.Enabled = False
        Txt_FileName.Text = File_FileList.Pattern
        VLastChanged = DIR_CHANGED
    ElseIf Not Err Then
        Pic_SelectedIcon.Picture = Pic_SelectedIcon.DragIcon
    End If

    Valid_Icon = (Err = 0)
    
    On Error GoTo 0

End Function

