VERSION 5.00
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.5#0"; "HACKPROG.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWin_Run 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   7440
   ClientTop       =   45
   ClientWidth     =   4530
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "BJ's Win Run.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5355
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog cdg 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3000
      Picture         =   "BJ's Win Run.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   7640
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "BJ's Win Run.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   7640
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "BJ's Win Run.frx":091E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   7640
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "BJ's Win Run.frx":0C28
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   7640
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1800
      Picture         =   "BJ's Win Run.frx":0F32
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   7640
      Width           =   480
   End
   Begin VB.Timer popupTmr1 
      Interval        =   1000
      Left            =   1440
      Top             =   7080
   End
   Begin VB.Timer Icontmr 
      Interval        =   1000
      Left            =   1920
      Top             =   7080
   End
   Begin VB.Timer popupTmr 
      Interval        =   1000
      Left            =   2520
      Top             =   7080
   End
   Begin VB.Timer tmrCheckTime 
      Interval        =   1000
      Left            =   840
      Top             =   7080
   End
   Begin CCRProgressBar.ccrpProgressBar Bar 
      Height          =   735
      Index           =   0
      Left            =   120
      Top             =   5400
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   1296
      AutoCaption     =   3
      AutoCaptionSuffix=   "Days"
      BackColor       =   0
      Caption         =   "0 Days"
      FillColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      IncrementSize   =   1
      Max             =   60
      MousePointer    =   11
      Smooth          =   -1  'True
      Vertical        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Windows has been running for..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "right click for menu"
      Top             =   80
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   120
      X2              =   4400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   8
      X1              =   120
      X2              =   4440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   9
      X1              =   120
      X2              =   4440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   7
      X1              =   120
      X2              =   4400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   6
      X1              =   120
      X2              =   4400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   5
      X1              =   120
      X2              =   4400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   4
      X1              =   120
      X2              =   4400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   120
      X2              =   4400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   4400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label DateTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Date && Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "right click for menu"
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label lblSecs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "right click for menu"
      Top             =   2800
      Width           =   4335
   End
   Begin VB.Label lblMins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "right click for menu"
      Top             =   2100
      Width           =   4335
   End
   Begin VB.Label lblHrs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "right click for menu"
      Top             =   1400
      Width           =   4335
   End
   Begin VB.Label lblDays 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "right click for menu"
      Top             =   680
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   4400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Begin VB.Menu mnuShow 
            Caption         =   "&Show"
         End
         Begin VB.Menu mnuHide 
            Caption         =   "&Hide"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuView 
            Caption         =   "&View Type"
            Begin VB.Menu mnuLblLook 
               Caption         =   "&Normal Look"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuProgLook 
               Caption         =   "&Progress Bar Look"
            End
            Begin VB.Menu mnu12 
               Caption         =   "-"
            End
            Begin VB.Menu mnuProgView 
               Caption         =   "Progress Bar &Views"
               Enabled         =   0   'False
               Begin VB.Menu mnuProgViewbar1 
                  Caption         =   "&Days"
                  Begin VB.Menu mnuProgViewDaysbar 
                     Caption         =   "&Bar View"
                     Checked         =   -1  'True
                  End
                  Begin VB.Menu mnuProgViewDaysClock 
                     Caption         =   "&Clock View"
                  End
                  Begin VB.Menu mnuProgViewDaysPie 
                     Caption         =   "&Pie View"
                  End
                  Begin VB.Menu mnu13 
                     Caption         =   "-"
                  End
                  Begin VB.Menu mnuProgViewDaysBackColor 
                     Caption         =   "&Back Color"
                  End
                  Begin VB.Menu mnuProgViewDaysFillColor 
                     Caption         =   "&Fill Color"
                  End
                  Begin VB.Menu mnuProgViewDaysForeColor 
                     Caption         =   "F&ore Color"
                  End
               End
               Begin VB.Menu mnu14 
                  Caption         =   "-"
               End
               Begin VB.Menu mnuProgViewbar2 
                  Caption         =   "&Hours"
                  Begin VB.Menu mnuProgViewHrsbar 
                     Caption         =   "&Bar View"
                     Checked         =   -1  'True
                  End
                  Begin VB.Menu mnuProgViewHrsClock 
                     Caption         =   "&Clock View"
                  End
                  Begin VB.Menu mnuProgViewHrsPie 
                     Caption         =   "&Pie View"
                  End
                  Begin VB.Menu mnu15 
                     Caption         =   "-"
                  End
                  Begin VB.Menu mnuProgViewHrsBackColor 
                     Caption         =   "&Back Color"
                  End
                  Begin VB.Menu mnuProgViewHrsFillColor 
                     Caption         =   "&Fill Color"
                  End
                  Begin VB.Menu mnuProgViewHrsForeColor 
                     Caption         =   "F&ore Color"
                  End
               End
               Begin VB.Menu mnu16 
                  Caption         =   "-"
               End
               Begin VB.Menu mnuProgViewbar3 
                  Caption         =   "&Minutes"
                  Begin VB.Menu mnuProgViewMinsbar 
                     Caption         =   "&Bar View"
                     Checked         =   -1  'True
                  End
                  Begin VB.Menu mnuProgViewMinsClock 
                     Caption         =   "&Clock View"
                  End
                  Begin VB.Menu mnuProgViewMinsPie 
                     Caption         =   "&Pie View"
                  End
                  Begin VB.Menu mnu17 
                     Caption         =   "-"
                  End
                  Begin VB.Menu mnuProgViewMinsBackColor 
                     Caption         =   "&Back Color"
                  End
                  Begin VB.Menu mnuProgViewMinsFillColor 
                     Caption         =   "&Fill Color"
                  End
                  Begin VB.Menu mnuProgViewMinsForeColor 
                     Caption         =   "F&ore Color"
                  End
               End
               Begin VB.Menu mnu18 
                  Caption         =   "-"
               End
               Begin VB.Menu mnuProgViewbar4 
                  Caption         =   "&Seconds"
                  Begin VB.Menu mnuProgViewSecsbar 
                     Caption         =   "&Bar View"
                     Checked         =   -1  'True
                  End
                  Begin VB.Menu mnuProgViewSecsClock 
                     Caption         =   "&Clock View"
                  End
                  Begin VB.Menu mnuProgViewSecsPie 
                     Caption         =   "&Pie View"
                  End
                  Begin VB.Menu mnu19 
                     Caption         =   "-"
                  End
                  Begin VB.Menu mnuProgViewSecsBackColor 
                     Caption         =   "&Back Color"
                  End
                  Begin VB.Menu mnuProgViewSecsFillColor 
                     Caption         =   "&Fill Color"
                  End
                  Begin VB.Menu mnuProgViewSecsForeColor 
                     Caption         =   "F&ore Color"
                  End
               End
            End
         End
      End
      Begin VB.Menu mnu20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMail 
         Caption         =   "&E-Mail me"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnu21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmWin_Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Const MS_PER_SEC As Long = 1000
Const MS_PER_MIN = MS_PER_SEC * 60
Const MS_PER_HR = MS_PER_MIN * 60
Const MS_PER_DAY = MS_PER_HR * 24

Dim ms As Long
Dim secs As Long
Dim mins As Long
Dim hrs As Long
Dim days As Long

Private Sub Bar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If

End Sub

Private Sub Form_Click()
Me.PopupMenu mnuPopup
End Sub

Private Sub Form_Load()
mnuHide.Visible = False
' The line below is linked to Icontmr. This is to have an animated icon
frmWin_Run.Icon = Picture2.Picture
Icontmr.Enabled = True
' Uncomment the next line only if you want a control box. ie.. shows caption at top
' Me.Caption = "Windows has been running for..."
popupTmr.Enabled = True
tmrCheckTime.Enabled = True
mnuLblLook.Checked = True
mnuProgLook.Checked = False
'--------------------------------------------------------------------------------------------
    ' used to set Icon in Tray
    
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
    TrayIcon.hIcon = frmWin_Run.Icon
    ' The Tooltip for the icon - sets the Tooltip that will be displayed
    TrayIcon.szTip = "Double Click for to Open, Right Click for Menu." & Chr$(0)
    
    ' Add icon to the tray by calling the Shell_NotifyIcon API
    'NIM_ADD is a Constant - add icon to tray
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
    
    ' Don't let application appear in the Windows task list
    App.TaskVisible = False
'--------------------------------------------------------------------------------------------

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' This is used for your mouse

Static Message As Long
Static RR As Boolean
    
    'x is the current mouse location along the x-axis
    Message = X / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case Message
            ' Left double click (This should bring up a dialog box)
            Case WM_LBUTTONDBLCLK
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
    popupTmr.Enabled = False
                Me.Show
                mnuHide.Visible = True
                mnuShow.Visible = False

            ' Right button up (This should bring up the popup menu)
            Case WM_RBUTTONUP
                Me.PopupMenu mnuPopup
        End Select
        RR = False
    End If
    
End Sub
'-------------------------------------------------------------------------------------------
'uncomment below to have a progressbar type view. Expend the forms height from 1455 to 7830
'-------------------------------------------------------------------------------------------
' The following is Used to Close the App
'You can use and Index to all. I couldn't be bothered doing it
'
Private Sub Form_Paint()
   Static AddBar As Integer, i As Integer
   If mnuProgLook.Checked = True Then
   ' When form is painting for first time,
         Bar(0).Top = 480
   If AddBar <> True Then
      For i = 1 To 3
         Load Bar(i)   ' Add five progressbars to array.
         Bar(i).Top = Bar(i - 1).Top + 840
         Bar(i).Visible = True
      Next i
      Bar(0).AutoCaptionSuffix = "Days"   ' Put caption on each progressbars.
      Bar(0).ForeColor = lblDays.ForeColor
      Bar(1).AutoCaptionSuffix = "Hours"
      Bar(1).ForeColor = lblHrs.ForeColor
      Bar(2).AutoCaptionSuffix = "Minutes"
      Bar(2).ForeColor = lblMins.ForeColor
      Bar(3).AutoCaptionSuffix = "Seconds"
      Bar(3).ForeColor = lblSecs.ForeColor
      AddBar = True   ' Form is done painting.
   End If
'Uncomment below in tmrCheckTime_Timer as well
        Bar(0).Value = days
        Bar(1).Value = hrs
        Bar(2).Value = mins
        Bar(3).Value = secs
        Else
Exit Sub
End If

End Sub

Private Sub mnuAbout_Click()
AboutBox Me.hwnd
End Sub

Private Sub mnuExit_Click()
Dim Msg, Style, Title, Response, MyString
Msg = lblDays.Caption & vbNewLine & _
    lblHrs.Caption & vbNewLine & _
    lblMins.Caption & vbNewLine & _
    lblSecs.Caption & vbNewLine & vbNewLine & _
    DateTime.Caption & vbNewLine & vbNewLine & _
    "Version: " & App.Major & "." & App.Minor & "." & App.Revision & "." & "b"
Style = vbOKCancel + vbQuestion
Title = Label1.Caption
Response = MsgBox(Msg, Style, Title)
If Response = vbOK Then
   MyString = "OK"
   
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = Me.hwnd
    TrayIcon.uId = vbNull
    'Remove icon from Tray
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    Unload frmWin_Run

Else
   MyString = "Cancel"

Exit Sub
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
mnuExit_Click
End Sub

Private Sub Form_Terminate()
mnuExit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
mnuExit_Click
End Sub

' Used to Hide the App
Private Sub mnuHide_Click()
mnuShow.Visible = True
mnuHide.Visible = False
popupTmr.Enabled = True
Me.Hide
End Sub

'When you right click anywhere on the app you get a popup menu

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If

End Sub

Private Sub DateTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If

End Sub

Private Sub lblDays_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If
End Sub

Private Sub lblHrs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If
End Sub

Private Sub lblMins_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If
End Sub

Private Sub lblSecs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu mnuPopup
Else
If Button = vbLeftButton Then
Exit Sub
End If
End If
End Sub

Private Sub mnuLblLook_Click()
Dim i As Integer
mnuLblLook.Checked = True
mnuProgLook.Checked = False
mnuProgView.Enabled = False
         Bar(0).Visible = False
         Bar(1).Visible = False
         Bar(2).Visible = False
         Bar(3).Visible = False
Call Form_Load
lblDays.Visible = True
lblHrs.Visible = True
lblMins.Visible = True
lblSecs.Visible = True
      For i = 1 To 9
Line1(i).Visible = True
Next i
End Sub

Private Sub mnuProgLook_Click()
Dim i As Integer
mnuLblLook.Checked = False
mnuProgLook.Checked = True
mnuProgView.Enabled = True
If mnuProgLook.Checked = True Then
lblDays.Visible = False
lblHrs.Visible = False
lblMins.Visible = False
lblSecs.Visible = False
Call Form_Paint
      For i = 1 To 9
Line1(i).Visible = False
Next i
Else
Exit Sub
End If
         Bar(0).Visible = True
         Bar(1).Visible = True
         Bar(2).Visible = True
         Bar(3).Visible = True
End Sub
'-------------------------------------------------------------------------------------------
'Days
Private Sub mnuProgViewDaysBar_Click()
mnuProgViewDaysbar.Checked = True
If mnuProgViewDaysbar.Checked = True Then
mnuProgViewDaysClock.Checked = False
mnuProgViewDaysPie.Checked = False
Bar(0).FillStyle = prgBar
'Bar(1).FillStyle = prgBar
'Bar(2).FillStyle = prgBar
'Bar(3).FillStyle = prgBar
End If
End Sub
'Days

Private Sub mnuProgViewDaysClock_Click()
mnuProgViewDaysClock.Checked = True
If mnuProgViewDaysClock.Checked = True Then
mnuProgViewDaysbar.Checked = False
mnuProgViewDaysPie.Checked = False
Bar(0).FillStyle = prgClock
'Bar(1).FillStyle = prgClock
'Bar(2).FillStyle = prgClock
'Bar(3).FillStyle = prgClock
End If
End Sub
'Days

Private Sub mnuProgViewDaysPie_Click()
mnuProgViewDaysPie.Checked = True
If mnuProgViewDaysPie.Checked = True Then
mnuProgViewDaysClock.Checked = False
mnuProgViewDaysbar.Checked = False
Bar(0).FillStyle = prgPie
'Bar(1).FillStyle = prgPie
'Bar(2).FillStyle = prgPie
'Bar(3).FillStyle = prgPie
End If
End Sub
'Days

Private Sub mnuProgViewDaysBackColor_Click()
cdg.ShowColor
Bar(0).BackColor = cdg.Color
'Bar(1).BackColor = cdg.Color
'Bar(2).BackColor = cdg.Color
'Bar(3).BackColor = cdg.Color
End Sub
'Days

Private Sub mnuProgViewDaysFillColor_Click()
cdg.ShowColor
Bar(0).FillColor = cdg.Color
'Bar(1).FillColor = cdg.Color
'Bar(2).FillColor = cdg.Color
'Bar(3).FillColor = cdg.Color

End Sub
'Days

Private Sub mnuProgViewDaysForeColor_Click()
cdg.ShowColor
Bar(0).ForeColor = cdg.Color
'Bar(1).ForeColor = cdg.Color
'Bar(2).ForeColor = cdg.Color
'Bar(3).ForeColor = cdg.Color

End Sub
'Days

'-------------------------------------------------------------------------------------------
'Hours
Private Sub mnuProgViewHrsBar_Click()
mnuProgViewHrsbar.Checked = True
If mnuProgViewHrsbar.Checked = True Then
mnuProgViewHrsClock.Checked = False
mnuProgViewHrsPie.Checked = False
'Bar(0).FillStyle = prgBar
Bar(1).FillStyle = prgBar
'Bar(2).FillStyle = prgBar
'Bar(3).FillStyle = prgBar
End If
End Sub
'Hours

Private Sub mnuProgViewHrsClock_Click()
mnuProgViewHrsClock.Checked = True
If mnuProgViewHrsClock.Checked = True Then
mnuProgViewHrsbar.Checked = False
mnuProgViewHrsPie.Checked = False
'Bar(0).FillStyle = prgClock
Bar(1).FillStyle = prgClock
'Bar(2).FillStyle = prgClock
'Bar(3).FillStyle = prgClock
End If
End Sub
'Hours

Private Sub mnuProgViewHrsPie_Click()
mnuProgViewHrsPie.Checked = True
If mnuProgViewHrsPie.Checked = True Then
mnuProgViewHrsClock.Checked = False
mnuProgViewHrsbar.Checked = False
'Bar(0).FillStyle = prgPie
Bar(1).FillStyle = prgPie
'Bar(2).FillStyle = prgPie
'Bar(3).FillStyle = prgPie
End If
End Sub
'Hours

Private Sub mnuProgViewHrsBackColor_Click()
cdg.ShowColor
'Bar(0).BackColor = cdg.Color
Bar(1).BackColor = cdg.Color
'Bar(2).BackColor = cdg.Color
'Bar(3).BackColor = cdg.Color
End Sub
'Hours

Private Sub mnuProgViewHrsFillColor_Click()
cdg.ShowColor
'Bar(0).FillColor = cdg.Color
Bar(1).FillColor = cdg.Color
'Bar(2).FillColor = cdg.Color
'Bar(3).FillColor = cdg.Color

End Sub
'Hours

Private Sub mnuProgViewHrsForeColor_Click()
cdg.ShowColor
'Bar(0).ForeColor = cdg.Color
Bar(1).ForeColor = cdg.Color
'Bar(2).ForeColor = cdg.Color
'Bar(3).ForeColor = cdg.Color

End Sub
'Hours

'-------------------------------------------------------------------------------------------
'Minutes
Private Sub mnuProgViewMinsBar_Click()
mnuProgViewMinsbar.Checked = True
If mnuProgViewMinsbar.Checked = True Then
mnuProgViewMinsClock.Checked = False
mnuProgViewMinsPie.Checked = False
'Bar(0).FillStyle = prgBar
'Bar(1).FillStyle = prgBar
Bar(2).FillStyle = prgBar
'Bar(3).FillStyle = prgBar
End If
End Sub
'Minutes

Private Sub mnuProgViewMinsClock_Click()
mnuProgViewMinsClock.Checked = True
If mnuProgViewMinsClock.Checked = True Then
mnuProgViewMinsbar.Checked = False
mnuProgViewMinsPie.Checked = False
'Bar(0).FillStyle = prgClock
'Bar(1).FillStyle = prgClock
Bar(2).FillStyle = prgClock
'Bar(3).FillStyle = prgClock
End If
End Sub
'Minutes

Private Sub mnuProgViewMinsPie_Click()
mnuProgViewMinsPie.Checked = True
If mnuProgViewMinsPie.Checked = True Then
mnuProgViewMinsClock.Checked = False
mnuProgViewMinsbar.Checked = False
'Bar(0).FillStyle = prgPie
'Bar(1).FillStyle = prgPie
Bar(2).FillStyle = prgPie
'Bar(3).FillStyle = prgPie
End If
End Sub
'Minutes

Private Sub mnuProgViewMinsBackColor_Click()
cdg.ShowColor
'Bar(0).BackColor = cdg.Color
'Bar(1).BackColor = cdg.Color
Bar(2).BackColor = cdg.Color
'Bar(3).BackColor = cdg.Color
End Sub
'Minutes

Private Sub mnuProgViewMinsFillColor_Click()
cdg.ShowColor
'Bar(0).FillColor = cdg.Color
'Bar(1).FillColor = cdg.Color
Bar(2).FillColor = cdg.Color
'Bar(3).FillColor = cdg.Color

End Sub
'Minutes

Private Sub mnuProgViewMinsForeColor_Click()
cdg.ShowColor
'Bar(0).ForeColor = cdg.Color
'Bar(1).ForeColor = cdg.Color
Bar(2).ForeColor = cdg.Color
'Bar(3).ForeColor = cdg.Color

End Sub
'Minutes

'-------------------------------------------------------------------------------------------
'Seconds
Private Sub mnuProgViewSecsBar_Click()
mnuProgViewSecsbar.Checked = True
If mnuProgViewSecsbar.Checked = True Then
mnuProgViewSecsClock.Checked = False
mnuProgViewSecsPie.Checked = False
'Bar(0).FillStyle = prgBar
'Bar(1).FillStyle = prgBar
'Bar(2).FillStyle = prgBar
Bar(3).FillStyle = prgBar
End If
End Sub
'Seconds

Private Sub mnuProgViewSecsClock_Click()
mnuProgViewSecsClock.Checked = True
If mnuProgViewSecsClock.Checked = True Then
mnuProgViewSecsbar.Checked = False
mnuProgViewSecsPie.Checked = False
'Bar(0).FillStyle = prgClock
'Bar(1).FillStyle = prgClock
'Bar(2).FillStyle = prgClock
Bar(3).FillStyle = prgClock
End If
End Sub
'Seconds

Private Sub mnuProgViewSecsPie_Click()
mnuProgViewSecsPie.Checked = True
If mnuProgViewSecsPie.Checked = True Then
mnuProgViewSecsClock.Checked = False
mnuProgViewSecsbar.Checked = False
'Bar(0).FillStyle = prgPie
'Bar(1).FillStyle = prgPie
'Bar(2).FillStyle = prgPie
Bar(3).FillStyle = prgPie
End If
End Sub
'Seconds

Private Sub mnuProgViewSecsBackColor_Click()
cdg.ShowColor
'Bar(0).BackColor = cdg.Color
'Bar(1).BackColor = cdg.Color
'Bar(2).BackColor = cdg.Color
Bar(3).BackColor = cdg.Color
End Sub
'Seconds

Private Sub mnuProgViewSecsFillColor_Click()
cdg.ShowColor
'Bar(0).FillColor = cdg.Color
'Bar(1).FillColor = cdg.Color
'Bar(2).FillColor = cdg.Color
Bar(3).FillColor = cdg.Color

End Sub
'Seconds

Private Sub mnuProgViewSecsForeColor_Click()
cdg.ShowColor
'Bar(0).ForeColor = cdg.Color
'Bar(1).ForeColor = cdg.Color
'Bar(2).ForeColor = cdg.Color
Bar(3).ForeColor = cdg.Color

End Sub
'Seconds


'-------------------------------------------------------------------------------------------
' Used to Show the App
Private Sub mnuShow_Click()
mnuHide.Visible = True
mnuShow.Visible = False
popupTmr.Enabled = False
Me.Show
End Sub

Private Sub popupTmr_Timer()

'The following can be changed to suite your needs

'Form will Popup at 0 Minutes
    If lblMins.Caption = "0 Minutes" Then
        Me.Show
'Form will hide at 1 Minute
    ElseIf lblMins.Caption = "1 Minute" Then
        Me.Hide
'Form will Popup at 15 Minutes
    ElseIf lblMins.Caption = "15 Minutes" Then
        Me.Show
'Form will hide at 16 Minutes
    ElseIf lblMins.Caption = "16 Minutes" Then
        Me.Hide
'Form will Popup at 30 Minutes
    ElseIf lblMins.Caption = "30 Minutes" Then
        Me.Show
'Form will hide at 31 Minutes
    ElseIf lblMins.Caption = "31 Minutes" Then
        Me.Hide
'Form will Popup at 45 Minutes
    ElseIf lblMins.Caption = "45 Minutes" Then
        Me.Show
'Form will hide at 46 Minutes
    ElseIf lblMins.Caption = "46 Minutes" Then
        Me.Hide
    Else
'Form will be hidden
        Me.Hide
    End If
End Sub

Private Sub tmrCheckTime_Timer()
Dim Today As Variant
Dim MyDate, MyTime, MyDay, MyDay1, MyDay2, MyWeek, MyWeek1, MyWeek2
   Static AddBar As Integer, i As Integer
'Format to get the time windows has been open for
     
     ms = GetTickCount()
    days = ms \ MS_PER_DAY
    ms = ms - days * MS_PER_DAY
    hrs = ms \ MS_PER_HR
    ms = ms - hrs * MS_PER_HR
    mins = ms \ MS_PER_MIN
    ms = ms - mins * MS_PER_MIN
    secs = ms \ MS_PER_SEC
    ms = ms - secs * MS_PER_SEC
    
'The following will display for Eg... If you have 1 Second you will see
' 1 Seconds, with the format below it will show 1 Second, 2 Seconds etc...

If days = 1 Then
    lblDays.Caption = Format$(days) & " Day"
Else
lblDays.Caption = Format$(days) & " Days"
End If

If hrs = 1 Then
    lblHrs.Caption = Format$(hrs) & " Hour"
Else
    lblHrs.Caption = Format$(hrs) & " Hours"
End If

If mins = 1 Then
    lblMins.Caption = Format$(mins) & " Minute"
Else
lblMins.Caption = Format$(mins) & " Minutes"
End If
    
If secs = 1 Then
    lblSecs.Caption = Format$(secs) & " Second"
    Else
    lblSecs.Caption = Format$(secs) & " Seconds"
End If
    
'Form will be hidden every time the seconds gets to 59, including being opened manually
    If secs = 0 Then
        Me.Hide
End If
'-------------------------------------------------------------------------------------------
' This is to show the date and time
'you can remove MyDay, MyDay2, MyWeek, MyWeek2, But one thing is that it
'won't show (Day) 10 (of 52) or (Week) 239 (of 366). It will just show
'10 or 239

'The full format will be like this

'with MyDay, MyDay2, MyWeek, MyWeek2

'Sunday, 28 September 1999
'04:39:58 AM
'Day 222 of 366
'Week 30 of 52

'or without MyDay, MyDay2, MyWeek, MyWeek2

'Sunday, 28 September 1999
'04:39:58 AM
'222
'30

'Look at the one I have commented out
Today = Now


MyDate = Format(Date, "dddd, " & "dd " & "mmmm " & "yyyy ") 'Sunday, 28 September 1999
MyTime = Format(Time, "hh:mm:ss ampm ") '04:39:58 AM
MyDay = Format("Day") 'Day
MyDay1 = Format(Date, "y") '222
MyDay2 = Format("of 366") 'of 366
MyWeek = Format("Week") 'Week
MyWeek1 = Format(Date, "ww") '30
MyWeek2 = Format("of 52") 'of 52

DateTime.Caption = MyDate & vbNewLine & _
 MyTime & vbNewLine & _
" " & MyDay & " " & MyDay1 & " " & MyDay2 & vbNewLine & _
" " & MyWeek & " " & MyWeek1 & " " & MyWeek2

'DateTime.Caption = Format(Today, "dddd " & "dd " & "mmmm " & "yyyy " & "hh:mm:ss ampm " & vbCrLf & _
'"y" & " of 366 - " & "ww" & " of 52")
    
'if you know of a better way let me know
'E-Mail me bryce3@bigpond.com
End Sub
Private Sub Icontmr_Timer()
    Static i As Integer
    Picture2.Picture = Picture1(i).Picture
    'increment the picture counter
    i = i + 1
    If i = 4 Then i = 0
frmWin_Run.Icon = Picture2.Picture

If mnuProgLook.Checked = True Then
        Bar(0).Value = days
        Bar(1).Value = hrs
        Bar(2).Value = mins
        Bar(3).Value = secs
        Else
        Exit Sub
End If
End Sub

Private Sub mnuEMail_Click()
    ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's How to Get... Win Run.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Sub

