VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixer"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmMixer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Synthesizer"
      Height          =   5295
      Left            =   5090
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin ComctlLib.Slider sldSynFader 
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldSynBas 
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldSynTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldSynPan 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox chkSynMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldSyn 
         Height          =   1935
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   71
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   58
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   4
         Left            =   322
         TabIndex        =   53
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   4
         Left            =   405
         TabIndex        =   47
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblSyn 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblSyn"
         Height          =   195
         Left            =   330
         TabIndex        =   16
         Top             =   4560
         Width           =   420
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Wave"
      Height          =   5295
      Left            =   3660
      TabIndex        =   10
      Top             =   120
      Width           =   1455
      Begin Mixer.Progress Progress2 
         Height          =   4935
         Left            =   1080
         TabIndex        =   73
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   8705
         FillColor       =   -2147483633
      End
      Begin ComctlLib.Slider sldWaveFader 
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldWaveBas 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldWaveTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldWavePan 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox chkWavMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldWave 
         Height          =   1935
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   3
         Left            =   345
         TabIndex        =   70
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   57
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   3
         Left            =   322
         TabIndex        =   52
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   46
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblWave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblWave"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   4560
         Width           =   585
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ausgang"
      Height          =   5295
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      Begin ComctlLib.Slider sldMainFader 
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMainBas 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMainTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMainPan 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   960
         Top             =   3360
      End
      Begin VB.CheckBox chkVolMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldMain 
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin Mixer.Progress Progress1 
         Height          =   4935
         Left            =   1080
         TabIndex        =   60
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   8705
         FillColor       =   -2147483633
         Picture         =   "frmMixer.frx":0442
         Orientation     =   1
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   5
         Left            =   345
         TabIndex        =   72
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   5
         Left            =   375
         TabIndex        =   59
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   5
         Left            =   322
         TabIndex        =   54
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   48
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblMain"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   4560
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mikrofon"
      Height          =   5295
      Left            =   2480
      TabIndex        =   4
      Top             =   120
      Width           =   1215
      Begin ComctlLib.Slider sldMicFader 
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMicBas 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMicTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldMicPan 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox chkMicMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldMic 
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   69
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   56
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   2
         Left            =   322
         TabIndex        =   51
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   45
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblMic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblMic"
         Height          =   195
         Left            =   345
         TabIndex        =   14
         Top             =   4560
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CD"
      Height          =   5295
      Left            =   1300
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      Begin ComctlLib.Slider sldCDFader 
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldCDBas 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldCDTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldCDPan 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox chkCDMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldCD 
         Height          =   1935
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   68
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   55
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   1
         Left            =   322
         TabIndex        =   50
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   44
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblCD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblCD"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   4560
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin ComctlLib.Slider sldLineFader 
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldLineBas 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin ComctlLib.Slider sldLineTreb 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.CheckBox chkLinMute 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4920
         Width           =   735
      End
      Begin ComctlLib.Slider sldLine 
         Height          =   1935
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3413
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin ComctlLib.Slider sldLinePan 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fader"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   62
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pan"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   49
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bass"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   28
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treble"
         Height          =   195
         Index           =   0
         Left            =   322
         TabIndex        =   27
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblLine 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblLine"
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   4560
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------
' Sampleapplication that shows you how to use a part of the mixer API.
' Some functions either may not work with each soundcard, or this is very
' BUGGY. For example: Bass and Treble works only on the main channel with
' my computer. The Peakmeter isn´t doing anything since a view months
' and so on...
'
' Feel free to do with ist what you want.
'
' If you find a bug, and you´ve solved ist, let me know.
'
' Any suggestions, or anything else would be also fine!
'
' I hope I preventet some headaches with this project.
'
' Ascher Stefan (s.ascher@tirol.com)
'----------------------------------------------------------------------------

Option Explicit
Private vol As New cVolume

Private Sub chkCDMute_Click()
    vol.CompactDiscMute = IIf((chkCDMute.Value = 1), True, False)

End Sub

Private Sub chkLinMute_Click()
    vol.LineMute = IIf((chkLinMute.Value = 1), True, False)
End Sub

Private Sub chkMicMute_Click()
    vol.MikrofonMute = IIf((chkMicMute.Value = 1), True, False)
End Sub

Private Sub chkSynMute_Click()
    vol.SynthesizerMute = IIf((chkSynMute.Value = 1), True, False)
End Sub

Private Sub chkVolMute_Click()
     vol.VolumeMute = IIf((chkVolMute.Value = 1), True, False)
End Sub

Private Sub chkWavMute_Click()
    vol.WaveMute = IIf((chkWavMute.Value = 1), True, False)
End Sub

Private Sub Form_Load()
    Set vol = New cVolume
    
    With sldCD
        .Min = vol.CDMin
        .Max = vol.CDMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    
    With sldLine
        .Min = vol.LineMin
        .Max = vol.LineMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    
    With sldMain
        .Min = vol.VolumeMin
        .Max = vol.VolumeMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    
    With sldMic
        .Min = vol.MicMin
        .Max = vol.MicMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    
    With sldSyn
        .Min = vol.SynMin
        .Max = vol.SynMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    
    With sldWave
        .Min = vol.WaveMin
        .Max = vol.WaveMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With

' Additional Functions, they may not work with each Soundcard,
' but maybe ist´s only the wrong way ;-), or it´s impossible to do this???.
' I Don´t know!

    ' Line-------------------------------------------------------------------
    With sldLinePan
        If vol.VolPanMax = 0 Then
            .Visible = False
            Label4(0).Visible = False
        Else
            .Max = vol.VolPanMax
            .Min = vol.VolPanMin
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldLineFader
        If vol.LineFaderMax = 0 Then
            .Visible = False
            Label6(0).Visible = False
        Else
            .Min = vol.LineFaderMin
            .Max = vol.LineFaderMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldLineTreb
        If vol.LineTrebleMax = 0 Then
            .Visible = False
            Label2(0).Visible = False
        Else
            .Min = vol.LineTrebleMin
            .Max = vol.LineTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldLineBas
        If vol.LineBassMax = 0 Then
            .Visible = False
        Label3(0).Visible = False
        Else
            .Min = vol.LineBassMin
            .Max = vol.LineBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    ' CD---------------------------------------------------------------------
    With sldCDPan
        If vol.CDPanMax = 0 Then
            .Visible = False
            Label4(1).Visible = False
        Else
            .Min = vol.CDPanMin
            .Max = vol.CDPanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldCDBas
        If vol.CDBassMax = 0 Then
            .Visible = False
            Label3(1).Visible = False
        Else
            .Min = vol.CDBassMin
            .Max = vol.CDBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldCDFader
        If vol.CDFaderMax = 0 Then
            .Visible = False
            Label6(1).Visible = False
        Else
            .Min = vol.CDFaderMin
            .Max = vol.CDFaderMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldCDTreb
        If vol.CDTrebleMax = 0 Then
            .Visible = False
            Label2(1).Visible = False
        Else
            .Min = vol.CDTrebleMin
            .Max = vol.CDTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    ' Mic--------------------------------------------------------------------
    With sldMicPan
        If vol.MicPanMax = 0 Then
            .Visible = False
            Label4(2).Visible = False
        Else
            .Min = vol.MicPanMin
            .Max = vol.MicPanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMicBas
        If vol.MicBassMax = 0 Then
            .Visible = False
            Label3(2).Visible = False
        Else
            .Min = vol.MicBassMin
            .Max = vol.MicBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMicFader
        If vol.MicFaderMax = 0 Then
            .Visible = False
            Label6(2).Visible = False
        Else
            .Min = vol.MicFaderMin
            .Max = vol.MicFaderMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMicTreb
        If vol.MicTrebleMax = 0 Then
            .Visible = False
            Label2(2).Visible = False
        Else
            .Min = vol.MicTrebleMin
            .Max = vol.MicTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    ' Wave-------------------------------------------------------------------
    With sldWavePan
        If vol.WavePanMax = 0 Then
            .Visible = False
            Label4(3).Visible = False
        Else
            .Min = vol.WavePanMin
            .Max = vol.WavePanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldWaveBas
        If vol.WaveBassMax = 0 Then
            .Visible = False
            Label3(3).Visible = False
        Else
            .Min = vol.WaveBassMin
            .Max = vol.WaveBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldWaveFader
        If vol.WaveFaderMax = 0 Then
            .Visible = False
            Label6(3).Visible = False
        Else
            .Min = vol.WaveFaderMin
            .Max = vol.WaveFaderMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldWaveTreb
        If vol.WaveTrebleMax = 0 Then
            .Visible = False
            Label2(3).Visible = False
        Else
            .Min = vol.WaveTrebleMin
            .Max = vol.WaveTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    ' Synthesizer------------------------------------------------------------
    With sldSynPan
        If vol.SynPanMax = 0 Then
            .Visible = False
            Label4(4).Visible = False
        Else
            .Min = vol.SynPanMin
            .Max = vol.SynPanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldSynBas
        If vol.SynBassMax = 0 Then
            .Visible = False
            Label3(4).Visible = False
        Else
            .Min = vol.SynBassMin
            .Max = vol.SynBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldSynFader
        If vol.SynFaderMax = 0 Then
            .Visible = False
            Label6(4).Visible = False
        Else
            .Min = vol.SynFaderMin
            .Max = vol.SynFaderMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldSynTreb
        If vol.SynTrebleMax = 0 Then
            .Visible = False
            Label2(4).Visible = False
        Else
            .Min = vol.SynTrebleMin
            .Max = vol.SynTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    ' Main-------------------------------------------------------------------
    With sldMainPan
        If vol.VolPanMax = 0 Then
            .Visible = False
            Label4(5).Visible = False
        Else
            .Min = vol.VolPanMin
            .Max = vol.VolPanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMainFader
        If vol.VolPanMax = 0 Then
            .Visible = False
            Label6(5).Visible = False
        Else
            .Min = vol.VolFadMin
            .Max = vol.VolPanMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMainTreb
        If vol.VolTrebleMax = 0 Then
            .Visible = False
            Label2(5).Visible = False
        Else
            .Min = vol.VolTrebleMin
            .Max = vol.VolTrebleMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    With sldMainBas
        If vol.VolBassMax = 0 Then
            .Visible = False
            Label3(5).Visible = False
        Else
            .Min = vol.VolBassMin
            .Max = vol.VolBassMax
            .TickFrequency = (.Max - .Min) \ 10
            .LargeChange = .TickFrequency
        End If
    End With
    
    Progress1.Max = vol.MaxVolumeMeterInput
    Progress2.Max = vol.MaxVolumeMeterOutput
End Sub

Private Sub Form_Paint()
    sldCD.Value = sldCD.Max - vol.CDLevel
    lblCD.Caption = format$((vol.CDLevel / vol.CDMax), "##0 %")
    chkCDMute.Value = IIf(vol.CompactDiscMute, 1, 0)
    
    sldLine.Value = sldLine.Max - vol.LineLevel
    lblLine.Caption = format$((vol.LineLevel / vol.LineMax), "##0 %")
    chkLinMute.Value = IIf(vol.LineMute, 1, 0)
    
    sldMain.Value = sldMain.Max - vol.VolumeLevel
    lblMain.Caption = format$((vol.VolumeLevel / vol.VolumeMax), "##0 %")
    chkVolMute.Value = IIf(vol.VolumeMute, 1, 0)
    
    sldMic.Value = sldMic.Max - vol.MicLevel
    lblMic.Caption = format$((vol.MicLevel / vol.MicMax), "##0 %")
    chkMicMute.Value = IIf(vol.MikrofonMute, 1, 0)
    
    sldSyn.Value = sldSyn.Max - vol.SynLevel
    lblSyn.Caption = format$((vol.SynLevel / vol.SynMax), "##0 %")
    chkSynMute.Value = IIf(vol.SynthesizerMute, 1, 0)
    
    sldWave.Value = sldWave.Max - vol.WaveLevel
    lblWave.Caption = format$((vol.WaveLevel / vol.WaveMax), "##0 %")
    chkWavMute.Value = IIf(vol.WaveMute, 1, 0)
    
    If sldLinePan.Visible Then sldLinePan.Value = vol.LineLevelPan
    If sldLineFader.Visible Then sldLineFader.Value = vol.LineLevelFader
    If sldLineTreb.Visible Then sldLineTreb.Value = vol.LineLevelTreble
    If sldLineBas.Visible Then sldLineBas.Value = vol.LineLevelBass

    If sldCDPan.Visible Then sldCDPan.Value = vol.CDLevelPan
    If sldCDFader.Visible Then sldCDFader.Value = vol.CDLevelFader
    If sldCDTreb.Visible Then sldCDTreb.Value = vol.CDLevelTreble
    If sldCDBas.Visible Then sldCDBas.Value = vol.CDLevelBass

    If sldMicPan.Visible Then sldMicPan.Value = vol.MicLevelPan
    If sldMicFader.Visible Then sldMicFader.Value = vol.MicLevelFader
    If sldMicTreb.Visible Then sldMicTreb.Value = vol.MicLevelTreble
    If sldMicBas.Visible Then sldMicBas.Value = vol.MicLevelBass

    If sldWavePan.Visible Then sldWavePan.Value = vol.WaveLevelPan
    If sldWaveFader.Visible Then sldWaveFader.Value = vol.WaveLevelFader
    If sldWaveTreb.Visible Then sldWaveTreb.Value = vol.WaveLevelTreble
    If sldWaveBas.Visible Then sldWaveBas.Value = vol.WaveLevelBass

    If sldSynPan.Visible Then sldSynPan.Value = vol.SynLevelPan
    If sldSynFader.Visible Then sldSynFader.Value = vol.SynLevelFader
    If sldSynTreb.Visible Then sldSynTreb.Value = vol.SynLevelTreble
    If sldSynBas.Visible Then sldSynBas.Value = vol.SynLevelBass

    If sldMainPan.Visible Then sldMainPan.Value = vol.VolumeLevelPan
    If sldMainFader.Visible Then sldMainFader.Value = vol.VolumeLevelFader
    If sldMainTreb.Visible Then sldMainTreb.Value = vol.VolumeLevelTreble
    If sldMainBas.Visible Then sldMainBas.Value = vol.VolumeLevelBass
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set vol = Nothing
    
End Sub


Private Sub sldCD_Scroll()
    vol.CDLevel = sldCD.Max - sldCD.Value
    lblCD.Caption = format$((vol.CDLevel / vol.CDMax), "##0 %")

End Sub

Private Sub sldCDBas_Scroll()
    vol.CDLevelBass = sldCDBas.Value
End Sub

Private Sub sldCDFader_Scroll()
    vol.CDLevelFader = sldCDFader.Value
End Sub

Private Sub sldCDPan_Scroll()
    vol.CDLevelPan = sldCDPan.Value
End Sub

Private Sub sldCDTreb_Scroll()
    vol.CDLevelTreble = sldCDTreb.Value
End Sub

Private Sub sldLine_Scroll()
    vol.LineLevel = sldLine.Max - sldLine.Value
    lblLine.Caption = format$((vol.LineLevel / vol.LineMax), "##0 %")

End Sub

Private Sub sldLineBas_Scroll()
    vol.LineLevelBass = sldLineBas.Value
End Sub

Private Sub sldLineFader_Scroll()
    vol.LineLevelFader = sldLineFader.Value
End Sub

Private Sub sldLinePan_Scroll()
    vol.LineLevelPan = sldLinePan.Value

End Sub

Private Sub sldLineTreb_Scroll()
    vol.LineLevelTreble = sldLineTreb.Value
End Sub

Private Sub sldMain_Scroll()
    vol.VolumeLevel = sldMain.Max - sldMain.Value
    lblMain.Caption = format$((vol.VolumeLevel / vol.VolumeMax), "##0 %")

End Sub

Private Sub sldMainBas_Scroll()
    vol.VolumeLevelBass = sldMainBas.Value
End Sub

Private Sub sldMainFader_Scroll()
    vol.VolumeLevelFader = sldMainFader.Value
End Sub

Private Sub sldMainPan_Scroll()
    vol.VolumeLevelPan = sldMainPan.Value
End Sub

Private Sub sldMainTreb_Scroll()
    vol.VolumeLevelTreble = sldMainTreb.Value
End Sub

Private Sub sldMic_Scroll()
    vol.MicLevel = sldMic.Max - sldMic.Value
    lblMic.Caption = format$((vol.MicLevel / vol.MicMax), "##0 %")
End Sub

Private Sub sldMicBas_Scroll()
    vol.MicLevelBass = sldMicBas.Value
End Sub

Private Sub sldMicFader_Scroll()
    vol.MicLevelFader = sldMicFader.Value
End Sub

Private Sub sldMicPan_Scroll()
    vol.MicLevelPan = sldMicPan.Value
End Sub

Private Sub sldMicTreb_Scroll()
    vol.MicLevelTreble = sldMicTreb.Value
End Sub

Private Sub sldSyn_Scroll()
    vol.SynLevel = sldSyn.Max - sldSyn.Value
    lblSyn.Caption = format$((vol.SynLevel / vol.SynMax), "##0 %")

End Sub

Private Sub sldSynBas_Scroll()
    vol.SynLevelBass = sldSynBas.Value
End Sub

Private Sub sldSynFader_Scroll()
    vol.SynLevelFader = sldSynFader.Value
End Sub

Private Sub sldSynPan_Scroll()
    vol.SynLevelPan = sldSynPan.Value
End Sub

Private Sub sldSynTreb_Scroll()
    vol.SynLevelTreble = sldSynTreb.Value
End Sub

Private Sub sldWave_Scroll()
    vol.WaveLevel = sldWave.Max - sldWave.Value
    lblWave.Caption = format$((vol.WaveLevel / vol.WaveMax), "##0 %")

End Sub

Private Sub sldWaveBas_Scroll()
    vol.WaveLevelBass = sldWaveBas.Value
End Sub

Private Sub sldWaveFader_Scroll()
    vol.WaveLevelFader = sldWaveFader.Value

End Sub

Private Sub sldWavePan_Scroll()
    vol.WaveLevelPan = sldWavePan.Value
End Sub

Private Sub sldWaveTreb_Scroll()
    vol.WaveLevelTreble = sldWaveTreb.Value
End Sub

Private Sub Timer1_Timer()
    Progress1.Position = vol.CurrentVolumeMeterInput
    Progress2.Position = vol.CurrentVolumeMeterOutput
End Sub
