VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTrialVersion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BJ's How to Get... Trial Version."
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "BJ's Trial Version.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   5160
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   45
      TabIndex        =   15
      Top             =   3960
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
      Max             =   30
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "                Expire Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   40
      TabIndex        =   0
      Top             =   1080
      Width           =   3650
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   15
         X2              =   1080
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   3520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   3520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2960
         TabIndex        =   18
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BJ's Trial Version will shutdown in:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Times"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStart 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblA 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Started at:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   45
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTimes 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblB 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "App Used:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   40
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblC 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Begin Trial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   40
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblD 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Expired:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   40
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblF 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Days Left:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   40
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblTrial 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   2355
      End
      Begin VB.Label lblExpired 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   1680
         Width           =   2355
      End
      Begin VB.Label lblLeft 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label lblE 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   40
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   15
         X2              =   1080
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   0
      TabIndex        =   22
      Top             =   520
      Width           =   3735
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "RegisteredOwner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   170
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   1
      Left            =   50
      TabIndex        =   19
      Top             =   90
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   2
      Left            =   100
      TabIndex        =   21
      Top             =   10
      Width           =   2505
   End
   Begin VB.Label Bryce 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BJ's Trial Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   3
      Left            =   150
      TabIndex        =   28
      Top             =   -70
      Width           =   2505
   End
End
Attribute VB_Name = "frmTrialVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private DateToday As Date

Private Sub Bryce_Click(Index As Integer)
Select Case Index
Case 0 To 4
ShellExecute 0, "Open", "mailto:bryce3@bigpond.com?subject=Info about BJ's Howto Get... Trial Version.&body=Type here the message that you want to send me.", "", "", vbMaximizedFocus
End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

Label7.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")

If Label7.Caption = "Error" Then
Label7.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
End If

If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened") = "Error" Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened", "1"
End If
If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened") = "" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened", "1"
End If
lblTimes.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened") & "  times"
If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Expire Date") = "Error" Then
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Expire Date", Now + 30
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Trial Start Date", Now
End If
If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Expire Date") = "" Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Expire Date", Now + 30
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Trial Start Date", Now
End If
lblTrial.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Trial Start Date")
lblTimes.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened")
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Times Opened", lblTimes.Caption + 1
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Todays Date", Now
lblStart.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Todays Date")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Todays Date", lblStart.Caption
lblExpired.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Expire Date")
lblLeft = Format$(days360(lblStart.Caption, lblExpired.Caption), "###,###") & " Days"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Days", lblLeft.Caption
lblLeft.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Days")
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Copyright", App.LegalCopyright
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Trade Mark", App.LegalTrademarks
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Version", App.Major & "." & App.Minor & "." & App.Revision & "." & "BJ's Trial Version"

If lblLeft = 1 & " Days" Then
lblLeft = 1 & " Day"
End If

If Val(lblLeft.Caption) < 0 Then
MsgBox "Your trial version is expired!" & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
Unload frmTrialVersion

ElseIf Val(lblLeft.Caption) > 30 Then
MsgBox "Do not adjust Date/Time." & vbCrLf & _
"Your trial version is expired!" & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
Unload frmTrialVersion

ElseIf Val(lblTimes.Caption) > 10 Then
MsgBox "This has now been opened 10 times." & vbCrLf & _
"Please get a newer version to continue." & vbCrLf & vbCrLf & _
"Good Bye...", vbOKOnly + vbCritical, "This Trial Version has Expired"
Unload frmTrialVersion
Else
ProgressBar1.Value = Val(lblLeft)
End If
Exit Sub
ErrorHandler:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFF&
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\BJ's How to Get... Trial Version", "Todays Date", lblStart.Caption
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
Label2.Caption = 1
End Sub

Private Sub Timer1_Timer()
DateToday = Now
lblToday.Caption = Format$(DateToday, "ddd dd mmm yyyy - HH:mm:ss")
End Sub
Public Function days360(dt1 As Date, dt2 As Date) As Long
    
    Dim z1 As Long, z2 As Long
    Dim d1 As Long, d2 As Long
    Dim m1 As Long, m2 As Long
    Dim y1 As Long, y2 As Long
    
    d1 = Day(dt1)
    m1 = Month(dt1)
    y1 = Year(dt1)
    
    d2 = Day(dt2)
    m2 = Month(dt2)
    y2 = Year(dt2)
    
    If d1 = 31 Then
        z1 = 30
    Else
        z1 = d1
    End If
    
    If d2 = 31 And d1 >= 30 Then
        z2 = 30
    Else
        z2 = d2
    End If

    days360 = (y2 - y1) * 360 + (m2 - m1) * 30 + (z2 - z1)

End Function


Private Sub Timer2_Timer()
Label2.Caption = Label2.Caption - 1
If Label2.Caption < 2 Then Label3.Caption = "Second"
If Label2.Caption = 0 Then

Label8.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
If Label8.Caption = "Error" Then
Label8.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
End If

Label9.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
If Label9.Caption = "Error" Then
Label9.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
End If

Label10.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "SystemRoot")
If Label10.Caption = "Error" Then
Label10.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "SystemRoot")
End If

Label11.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
If Label11.Caption = "Error" Then
Label11.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
End If

MsgBox "Thank you for trying this Trial Version." & vbCrLf & vbCrLf & _
"This has been configured and works on:" & vbCrLf & _
"Windows 95, 98, 2000, ME, XP and maybe NT" & vbCrLf & vbCrLf & _
"I hope this is what you have been looking for." & vbCrLf & _
"This Example demonstrated how to access the registry." & vbCrLf & _
"-------------------------------------------------------------" & vbCrLf & _
"Eg... Hello to:" & vbCrLf & _
"Registered Owner: " & Label8.Caption & vbCrLf & _
"Registered Organization:" & Label9.Caption & vbCrLf & _
"Your Windows Directory is: " & Label10.Caption & vbCrLf & _
"You are currently running: " & Label11.Caption & vbCrLf & _
"-------------------------------------------------------------" & vbCrLf & _
"E-Mail me if you have any problems. Thanks. BJ" & vbCrLf & _
"(bryce3@bigpond.com)", vbInformation + vbOKOnly, frmTrialVersion.Caption & " Information."
If 1 Then Unload frmTrialVersion

End If
End Sub


