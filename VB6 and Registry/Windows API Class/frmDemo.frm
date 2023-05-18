VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SilverSoft WINAPI Class Demonstrator by Akhil"
   ClientHeight    =   6630
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Remove the Icon in Tray"
      Height          =   375
      Left            =   135
      TabIndex        =   9
      Top             =   6120
      Width           =   5010
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create an Icon in Tray"
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   5685
      Width           =   5010
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   5760
      Picture         =   "frmDemo.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   2010
      Width           =   540
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   105
      TabIndex        =   6
      Top             =   3720
      Width           =   5070
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enumerate Registry Keys"
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   3255
      Width           =   5040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Retrieve Icon of An Application"
      Height          =   375
      Left            =   135
      TabIndex        =   4
      Top             =   2565
      Width           =   4290
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   510
      Left            =   4530
      ScaleHeight     =   450
      ScaleWidth      =   570
      TabIndex        =   3
      Top             =   2505
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send E-Mail to Akhil P (akhiljayaraj@hotmail.com)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2025
      Width           =   5040
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   2640
      TabIndex        =   1
      Top             =   135
      Width           =   2520
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find all files in C:\"
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   2355
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code is coded by Akhil
'This Code is just for information and so most of the
'code is in commented form.Remove it to view the results.






Option Explicit
Dim CurrentPath As String ' for regenumerationkey
Dim clsWinApi As CWinAPI 'Declare a class variable

Private Sub cmdFind_Click()
Dim FileCount As Integer, DirCount As Integer
clsWinApi.FindFilesAPI "c:\", "*.*", FileCount, DirCount, List1.hWnd, False, True
MsgBox "No. of Files Found: " & FileCount, , "WINAPI Class Demonstrator"
End Sub

Private Sub Command1_Click()
clsWinApi.SendEmail "akhiljayaraj@hotmail.com", Me.hWnd
End Sub

Private Sub Command2_Click()
clsWinApi.RetrieveIcon "c:\windows\notepad.exe", Picture2.hDC, ricnLarge
Picture2.Picture = Picture2.Image
End Sub

Private Sub Command3_Click()
CurrentPath = "software\microsoft\windows\Currentversion"
clsWinApi.RegistryEnumerateKeys CurrentPath, List2.hWnd, LOCAL_MACHINE
End Sub

Private Sub Command4_Click()
clsWinApi.TrayIcon shelliconAdd, Picture1.hWnd, "Akhil", Picture1.Picture  ' To add an Icon to the Tray
End Sub

Private Sub Command5_Click()
clsWinApi.TrayIcon shelliconDelete, Picture1.hWnd, "Akhil", Picture1.Picture
End Sub

Private Sub Form_Load()
Set clsWinApi = New CWinAPI 'This makes a new instance of the class
MsgBox "This program just demonstrates a few functions that can be done with the CWinApi class and the other functions are present in the Code in commented form. You will be able to use other functions with no difficulties. This program is made to Demostrate the CWINApi class. So the UI is not that good.", vbApplicationModal + vbInformation + vbOKOnly, "WINAPI class Demonstrator"

End Sub
 Private Sub Form_Click()
'On Error GoTo er:

'These comments show how to use each functions

'clsWinApi.SetSysWorkArea 0, 0, 1024, 768  ' To set the system work area

'clswinapi.StartScreenSaver  'To start the screen saver

'clsWinApi.PlaySound sndWindowsStart  'To play system sounds such as the sound which is played when Windows starts

'clsWinApi.PlaySound sndWavFile, "c:\windows\media\chimes.wav"  'To Play a wave File

'clsWinApi.MakeTransparentForm frmDemo.hWnd  ' To make a Window Transparent

'MsgBox clsWinApi.GetSystemFolders(dirTemp)  ' To obtain system Folders Information

'clsWinApi.MenuBitmaps Form1.hWnd, 0, 0, Image1.Picture  ' To put a bitmap in a menu

'clswinapi.TrayIcon shelliconAdd, Picture1.hWnd, "Akhil", Picture1.Picture  ' To add an Icon to the Tray

'clswinapi.MoveFormWithoutBorders Form1.hWnd  'To move a form when the Title Bar of the form is disabled

'clswinapi.Launch strtHelp ': To launch various programs of the start menu

'clsWinApi.DocumentList "ADD", "e:\akhil.txt"   ' To add a file to the Document List

'clsWinApi.DocumentList "CLEAR", "d:\tiger.jpg"   ' To clear a file from the Document List

'clsWinApi.SetDeskWallPaper 1, "c:\windows\clouds.bmp" ' To set the wallpaper

'clsWinApi.SetDeskWallPaper 0 ' To set the wallpaper to None

'clswinapi.ExitWindows 1 ' To reboot Computer

'clswinapi.TaskBar 0 ' To hide the taskBar

'clswinapi.TaskBar 1 ' To show the taskBar

'Print clswinapi.WorkAreaLeft, clswinapi.TtoP(clswinapi.WorkAreaBottom) ' To retrieve the System Work Area Values

'Print clswinapi.TtoP(clswinapi.WorkAreaBottom) 'TtoP : Twips to Pixels Converter

'Print clswinapi.PtoT(clswinapi.TtoP(clswinapi.WorkAreaBottom)) 'PtoT : Pixels to Twips Converter

'clswinapi.Delay 2 ' Delay your program by 2 seconds

'clsWinApi.SetHotKey frmDemo.hWnd, keyControlD 'Set the hotkey of your program to CONTROL+D

'clswinapi.ShowMouse chtHide 'Hide the mouse

'MsgBox clswinapi.SoundCardDetect 'Detect sound Card

'clswinapi.PlayAvi "e:\animator\sprite.avi" 'Play Avi File

'clswinapi.ShowRecycleBin frmDemo.hWnd 'Show recycle Bin (requires Internet Explorer)

'clswinapi.OpenWebSite "http://www.microsoft.com" 'Open a website in IE Window

'clswinapi.SendEmail "somebody@somewhere.com" 'Send an E-Mail to a specified address

'clswinapi.ExecuteFile "notepad.exe" 'Open any File or Document

'clswinapi.ExecuteFile "c:\windows\notepad.exe", "e:\akhil.txt" Open the specified text File with the Specified program

'clsWinApi.SetMyPosTopMost Me.hWnd, False 'Set a Window in the TopMost Position

'clswinapi.SetMyPosOnDesktop hWnd 'Set a Window in the BottomMost Position

'MsgBox clswinapi.RetrieveFileTypeName("e:\akhil.txt")  'Retrieve a FileType Name

'clswinapi.RetrieveIcon "calc.exe", Picture2.hDC, ricnSmall: Picture2.Picture = Picture2.Image

'clswinapi.ListFunc 'Lists all the Functions of the WinAPI Class
'clswinapi.SetMousePos 100, 100 'Set the mouse Cursor to the Specified Pixel

'MsgBox clswinapi.MouseX 'Show the current Mouse X Position

'clswinapi.ChangeToolBarStyle Toolbar1.hWnd: Toolbar1.Refresh 'Changes toolbar style to that of IE

'clswinapi.RegistryValueS "software\microsoft\windows\currentversion\run", "Cheetah", "DeleteThis", LOCAL_MACHINE 'Make a string Value in the Registry

'clswinapi.RegistryValueD "software\microsoft\windows\currentversion\policies\explorer", "NoDesktop", 1, CURRENT_USER 'Make a DWORD Value

'clswinapi.RegistryRun "Notepad", "c:\windows\notepad.exe" 'You can use this function to load your application when windows starts.
'To get your App's EXE name, Use the function App.Exename

'clswinapi.RegistryDeleteValue "software\microsoft\windows\currentversion\run", "asdf", LOCAL_MACHINE 'Delete a registry value

'clswinapi.RegistryNewKey "Akhil", CURRENT_USER 'Create a new Key

'clswinapi.RegistryValueB "software\microsoft\windows\currentversion\policies\explorer", "NoDriveTypeAutoRun", 243, CURRENT_USER 'A new Binary Value

'Dim aa As Variant, bb As Long
'clswinapi.RegistryGetValue "software\microsoft\windows\currentversion\policies\explorer", "NoDriveTypeAutoRun", aa, regBinary, CURRENT_USER 'Retrieve a Value from registry

'Dim aa As String
'aa = Space$(11)
'clsWinApi.RegistryGetValue "software\microsoft\windows\currentversion\policies\explorer", "Cheetah", aa, regString, CURRENT_USER 'Retrieve a String Value from Registry
'MsgBox aa 'Please note that the Values given these examples are just for examples and may not Exist!

'clswinapi.RegistryDeletekey "Akhil", "", CURRENT_USER 'Delete a key from Registry

'clsWinApi.DisableCtrlAltDelete True ' Disable Ctrl+Alt+Del

'clsWinApi.DisableCtrlAltDelete False ' Enable Ctrl+Alt+Del

'The below lines of Code are for the FindFiles Function

'    Dim SearchPath As String, FindStr As String
'    Dim FileSize As Long
'    Dim NumFiles As Integer, NumDirs As Integer
'    List1.Clear
'    Screen.MousePointer = vbHourglass
'    SearchPath = "c:\"
'    FindStr = "*.bmp"
'    FileSize = clswinapi.FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs, List1.hWnd, True, False)
'    List1.AddItem NumFiles & " Files found in " & NumDirs + 1 & " Directories"
'    List1.AddItem "Size of " & NumFiles & " files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
'    Screen.MousePointer = vbDefault
        
        
'The below lines of code are to find the DiskSpace of your drive

'Dim aa As Double, bb As Double, cc As Double
't = clsWinApi.DiskSpace("c", aa, bb, cc) '"c" for C drive
' MsgBox aa & " " & bb & " " & cc

'clsWinApi.PlayAudioCd 1, False 'Play Audio CD track 1

'clsWinApi.PlayAudioCd 1, True 'Stop Playing Audio CD

'clswinapi.AutoRun runAllDrives
'---------------------------------------------------------------****Important Data*****-----------------------------------------------------------------
'Autorun Registry Values : These values are the ones to be passed to the Autorun function
'
'Drive -----------------------Value
'Floppy Drive - --------------251 'To autorun Floppies it must Contain Autorun.inf in the root directory and in it should be mentioned the program to be run
'Hard Drive - ----------------247
'Floppy & HardDrive --------- 243
'Data CD 's ------------------ 223
'Audio CD 's ----------------- 255
'Audio & Data CD's ---------- 223
'Floppy & Hard Drive & Data CD & Audio CD & Network Drive & RAM Drive - 131
'RAM Drive - -----------------191
'Network Drive - -------------239
'RAM & Network -------------- 175

'--------------------------------------------------------------------------------------------------------------------------------------------------------


'MsgBox clswinapi.ComputerName(cmpGetComputerName) 'Retrieve the Computer name

'-------------------------------------------------------
'Miscallenous Info About a Drive
'--------------------------------------------------------
'            Dim N As String, a As Long, m As Long, ad As String
'            N = Space$(255)
'            ad = Space$(255)
'            clswinapi.DiskVolumeInfo "x:\", N, 255, a, m, 0, ad, 255
'            If InStr(N, Chr$(0)) Then N = Left(N, Len(N) - 1)
'            MsgBox N ' & " " & a & " " & m & " " & ad
'            MsgBox a
'            MsgBox m
'            MsgBox ad



Exit Sub
er:
   MsgBox Err.Description & " " & Err.Number
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu Akhil
End Sub

'Private Sub List1_Click()
'On Error Resume Next
'Picture2.Picture = LoadPicture()
'clswinapi.RetrieveIcon List1.List(List1.ListIndex), Picture2.hDC, 32: Picture2.Picture = Picture2.Image
'End Sub
'
'Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'PopupMenu Akhil
'End Sub
'
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If clsWinApi.TryIcnRtnMsg(X) = xyzLeftButtonDoubleClick Then MsgBox "Hi, This is an Example of an Icon in Tray"
End Sub

Private Sub List1_Click()
'clswinapi.RegistryEnumerateKeys "software", False, List1.hWnd, CURRENT_USER
End Sub

Private Sub List1_DblClick()
'List1.Clear

'CurrentPath = CurrentPath & List1.List(List1.ListIndex) & "\"
End Sub

Private Sub Picture3_Click()
'clswinapi.RetrieveIcon "d:\calc.exe", Picture3.hDC, ricnSmall
'Picture3.Picture = Picture3.Image
'Picture = Picture3.Picture
'Refresh
'Picture3.Refresh
'clswinapi.MenuBitmaps Form1.hWnd, 0, 0, Image1.Picture

End Sub

