VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Alarm Clock v.1.2"
   ClientHeight    =   2055
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdStopPlayback 
      Caption         =   "S&top"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "2"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "Sn&ooze"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set Alarm..."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer tmrChangeTime 
      Interval        =   1000
      Left            =   4200
      Tag             =   "2"
      Top             =   0
   End
   Begin VB.Label lblTimeSet 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblSetTime 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFSet 
         Caption         =   "&Set Alarm"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAAbout 
         Caption         =   "&About Alarm Clock..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public alarmtime00 As String ' Global string variable representing the time that the alarm will go off
Public alarmtime01 As String ' Global string variable representing the time that the alarm will go off, simply the above variable plus one second
' The reason for the difference is because, depending on how much the timer sub has to do, the time is refreshed and skips over the :00 second

Private Sub cmdQuit_Click()

' Input: None
' Process: Quits the program
' Output: None

    If isPlaying = True Then ' If statement to check whether there is a music file playing
        Call cmdStopPlayback_Click ' Calls the cmdStopPlayback sub to stop the song and close the file
    End If
    End ' End the program
End Sub

Private Sub cmdSet_Click()

' Input: None
' Process: Loads the form that sets the alarm
' Output: None

    frmSetAlarm.Show ' Shows frmSetAlarm.frm
End Sub

Private Sub cmdSnooze_Click()

' Input: None
' Process: Sets the alarm time to 10 minutes ahead of the originally set time
' Output: None

    If alarmtime00 = "" Then Exit Sub ' Checks to see if there is a set time and exits if there is not
    If lblTimeSet.Caption <> "No Alarm Set" Then Exit Sub ' Checks to see if the alarm has gone off before
    
    Dim SnoozeTimeMin As String ' String representing the number of minutes
    Dim SnoozeTimeHour As String ' String representing the number of hours
    Dim SnoozeTimeAMPM As String ' String representing the time of day - AM or PM
     
    Call cmdStopPlayback_Click ' Calling the cmdStopPlayback sub
    
    SnoozeTimeMin = Mid(alarmtime00, Len(alarmtime00) - 7, 2) ' Isolating the number of minutes
    
    If Len(alarmtime00) = 10 Then ' Checking to see if there are less than 10 hours
        SnoozeTimeHour = Mid(alarmtime00, Len(alarmtime00) - 9, 1) ' Isolating the number of hours
    ElseIf Len(alarmtime00) = 11 Then ' Checking to see if there are more than 9 hours
        SnoozeTimeHour = Mid(alarmtime00, Len(alarmtime00) - 10, 2) ' Isolating the number if hours
    End If
    
    SnoozeTimeAMPM = Mid(alarmtime00, Len(alarmtime00) - 1, 2) ' Isolating the time of day
    
    ' This block of code will only execute if the snooze timeframe
    ' is within an hour change
    
    If SnoozeTimeMin + 10 > 59 Then ' Checking to see if there will be a clock roll-over during this process
        SnoozeTimeMin = "0" & CStr(0 + 10 - (60 - SnoozeTimeMin)) ' Rolling over the number of minutes
        If SnoozeTimeMin = "010" Then SnoozeTimeMin = CStr(10) ' Something useless
        SnoozeTimeHour = SnoozeTimeHour + 1 ' Increasing the number of hours by one to accomodate the minute roll-over
        If SnoozeTimeHour > 12 Then ' Checking to see if there will be a clock roll-over during this process
            SnoozeTimeHour = CStr(1) ' Setting the number of hours to 1, the next hour after 12:00
            If SnoozeTimeAMPM = "AM" Then ' Changing the time of day
                SnoozeTimeAMPM = "PM"
            ElseIf SnoozeTimeAMPM = "PM" Then ' Changing the time of day
                SnoozeTimeAMPM = "AM"
            End If
        End If
        alarmtime00 = SnoozeTimeHour + ":" + SnoozeTimeMin + ":00" + " " + SnoozeTimeAMPM ' If the IF statement was tripped, this is the new alarm time
        alarmtime01 = SnoozeTimeHour + ":" + SnoozeTimeMin + ":01" + " " + SnoozeTimeAMPM ' This is the backup alarm time
        lblTimeSet.FontSize = 12 ' Cosmetic adjustments
        lblSetTime.Caption = "Alarm is set for:" ' Changing the caption of the main form to reflect the set time
        lblTimeSet.Caption = alarmtime00 ' Changing the caption on the main form to reflect the set time
        Exit Sub
    End If
    
    ' This block of code will only execute if the snooze timeframe
    ' is NOT within an hour change
    
    SnoozeTimeMin = CStr(SnoozeTimeMin + 10) ' If the IF statement was NOT tripped, this is the new minute setting
    alarmtime00 = SnoozeTimeHour + ":" + SnoozeTimeMin + ":00" + " " + SnoozeTimeAMPM ' If the IF statement was NOT tripped, this is the new set time
    alarmtime01 = SnoozeTimeHour + ":" + SnoozeTimeMin + ":01" + " " + SnoozeTimeAMPM ' This is the backup time
    lblTimeSet.FontSize = 12 ' Cosmetic adjustments
    lblTimeSet.Caption = alarmtime00 ' Changing the caption on the main form to reflect the set time
    lblSetTime.Caption = "Alarm is set for:" ' Changing the caption on the main form to reflect the set time
End Sub

Private Sub cmdStopPlayback_Click()

' Input: None
' Process: Stops the playback of any music file that is currently playing
' Output: None
    
    Dim Stopresult As String ' String representing the return value of the StopMPEG function
    Dim Closeresult As String ' String representing the return value of the CloseMPEG function
    
    Stopresult = StopMPEG() ' Calling the function that stops the multimedia file
    Closeresult = CloseMPEG() ' Calling the function that closes the multimedia file
    isPlaying = False ' Flags the file as NOT playing
    
End Sub

Private Sub Form_Load()

' Input: None
' Process: Loads the form and executes the code below
' Output: None

    lblTime.Caption = Time ' Setting the caption of the main window on the form to the current system time
End Sub

Private Sub mnuAAbout_Click()

' Input: None
' Process: Shows the about box
' Output: None

    frmAbout.Show ' Shows the about box
End Sub

Private Sub mnuFQuit_Click()

' Input: None
' Process: Menu item that quits the program
' Output: None

    Call cmdQuit_Click ' Quitting the program
End Sub

Private Sub mnuFSet_Click()

' Input: None
' Process: Shows the SetAlarm form, similar to clicking the 'Set' button
' Output: None

    Call cmdSet_Click ' Shows the SetAlarm form
End Sub

Private Sub tmrChangeTime_Timer()

' Input: None
' Process: This is what happens when the timer clickes over every 1000 ms., or 1 sec.
' Output: None

    lblTime.Caption = Time ' Setting the caption of the main window on the form to the current system time
    If isPlaying = True Then Exit Sub ' If this is not there then the song will start twice, as per the if statement below
    If lblTime.Caption = alarmtime00 Or lblTime.Caption = alarmtime01 Then ' Checking to see if the alarm time matches up with the system time
        If tmrChangeTime.Tag = 1 Then ' Checking to see if the alarm is set for 'Silent' or 'Music'
            Call PlayMusic ' Calling the function that plays the sounds
        ElseIf tmrChangeTime.Tag = 0 Then ' Silent alarm
            MsgBox "The time is:" & vbCrLf & vbCrLf & alarmtime00
        End If
        lblTimeSet.FontSize = 10 ' Cosmetic adjustments
        lblTimeSet.Caption = "No Alarm Set" ' Changing the caption to reflect the alarm state
        lblSetTime.Caption = "" ' More of the same
    End If
End Sub
