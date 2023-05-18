VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form iMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internal Media Player"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      Height          =   300
      Left            =   4275
      TabIndex        =   2
      Top             =   15
      Width           =   285
   End
   Begin VB.ListBox PlayList 
      Height          =   450
      Left            =   915
      TabIndex        =   1
      Top             =   615
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4530
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -180
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "iMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Resize()
    MediaPlayer1.Height = Me.ScaleHeight
    MediaPlayer1.Width = Me.ScaleWidth
    
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If NewState = mpStopped Then
        If PlayList.ListCount > 1 And Dir(PlayList.List(PlayList.ListIndex)) <> "" Then
            If PlayList.ListIndex = PlayList.ListCount - 1 Then
                PlayList.ListIndex = 0
            Else
                PlayList.ListIndex = PlayList.ListIndex + 1
            End If
            MediaPlayer1.FileName = PlayList.List(PlayList.ListIndex)
            MediaPlayer1.Play
        ElseIf PlayList.ListCount = 1 And Dir(PlayList.List(PlayList.ListIndex)) <> "" Then
            MediaPlayer1.FileName = PlayList.List(PlayList.ListIndex)
            MediaPlayer1.Play
        ElseIf Dir(MediaPlayer1.FileName) <> "" Then
            MediaPlayer1.Play
        End If
    End If
End Sub
