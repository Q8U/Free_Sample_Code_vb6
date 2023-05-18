VERSION 5.00
Begin VB.Form frmTrayIcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BJ's How to Get... Tray Icon."
   ClientHeight    =   15
   ClientLeft      =   2610
   ClientTop       =   2160
   ClientWidth     =   3105
   Icon            =   "BJ's Tray Icon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   15
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp_Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu EMail 
         Caption         =   "E-Mail me"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub EMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Tray Icon.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus

End Sub

Private Sub Form_Load()

    
    'centers form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    'sets cbSize to the Length of TrayIcon
    TrayIcon.cbSize = Len(TrayIcon)
    ' Handle of the window used to handle messages - which is the this form
    TrayIcon.hwnd = Me.hwnd
    ' ID code of the icon
    TrayIcon.uId = vbNull
    ' Flags
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    ' ID of the call back message
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    ' The icon - sets the icon that should be used
    TrayIcon.hIcon = frmTrayIcon.Icon
    ' The Tooltip for the icon - sets the Tooltip that will be displayed
    TrayIcon.szTip = "Right Click for Menu, Double Click to Exit" & Chr$(0)
    
    ' Add icon to the tray by calling the Shell_NotifyIcon API
    'NIM_ADD is a Constant - add icon to tray
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
    
    ' Don't let application appear in the Windows task list
    App.TaskVisible = False

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static Message As Long
Static RR As Boolean
    
    'x is the current mouse location along the x-axis
    Message = X / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case Message
            ' Left double click (This should bring up a dialog box)
            Case WM_LBUTTONDBLCLK
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
                End
            ' Right button up (This should bring up a menu)
            Case WM_RBUTTONUP
                Me.PopupMenu mnuPopup
        End Select
        RR = False
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = Me.hwnd
    TrayIcon.uId = vbNull
    'Remove icon for Tray
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    
End Sub


Private Sub mnuAbout_Click()

Dim Msg, Style, Title, Response, MyString
Msg = "This is an Example of how to create a Tray Icon"   ' Define message.
Style = vbOKOnly + vbQuestion ' Define buttons.
Title = "Tray Icon Message"   ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbOKOnly Then   ' User chose Yes.
   MyString = "Yes"   ' Perform some action.
End If

End Sub

Private Sub mnuExit_Click()
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)

Unload frmTrayIcon
End Sub


