VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A450CC9B-4DC9-11D3-9DB5-444553540000}#1.0#0"; "EQProDemo.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQPro 1.5 Tester"
   ClientHeight    =   4350
   ClientLeft      =   3885
   ClientTop       =   4335
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5535
   Begin VB.CommandButton cmdQuery 
      Caption         =   "&Query Mixer"
      Height          =   420
      Left            =   60
      TabIndex        =   3
      Top             =   3855
      Width           =   1170
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4605
      TabIndex        =   9
      Top             =   3855
      Width           =   870
   End
   Begin VB.Timer tmrMonitor 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3120
      Top             =   3810
   End
   Begin VB.ComboBox cmbMixers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   135
      List            =   "frmMain.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   3615
   End
   Begin VB.ComboBox cmbControls 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":0446
      Left            =   135
      List            =   "frmMain.frx":0450
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1140
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mixer Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   0
      TabIndex        =   2
      Top             =   1725
      Width           =   3855
      Begin EQProDemo.ucEQPro EQProCtrl 
         Height          =   450
         Left            =   180
         TabIndex        =   14
         Top             =   855
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         LayoutAlign     =   0
      End
      Begin MSComctlLib.Slider sldPan 
         Height          =   315
         Left            =   2505
         TabIndex        =   13
         Top             =   1350
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.CommandButton cmdCenter 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2903
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Center"
         Top             =   1695
         Width           =   225
      End
      Begin VB.CheckBox chkInput 
         Caption         =   "Input"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Select as recording source"
         Top             =   1395
         Width           =   945
      End
      Begin VB.CheckBox chkMute 
         Caption         =   "Mute"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Mute"
         Top             =   1395
         Width           =   945
      End
      Begin VB.ComboBox cmbMixerLines 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":0475
         Left            =   120
         List            =   "frmMain.frx":0477
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   3615
      End
   End
   Begin VB.Frame frmAdvanced 
      Caption         =   "Advanced"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   3945
      TabIndex        =   0
      Top             =   105
      Width           =   1560
      Begin VB.CheckBox chkAdv 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3885
      Top             =   3810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available Mixers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   11
      Top             =   105
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   10
      Top             =   885
      Width           =   660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EQPro Tester application
'21-feb-1999
'by Xavier Flix (xFX JumpStart: Software Division)
'
'
'
'In order to understand how to use EQPro control and how this sample works
'start by the Form_Load event and keep folowing the comments, where I say
'that certain line of code triggers a certain event, follow it.
'
'
'
'
'
'
'Please excuse me for the gramatical errors, if any... :)
'
'
'

Option Explicit

Private Sub chkAdv_Click(Index As Integer)

    'The user clicked?
    'Well.. let's tell EQPro...
    
    Dim selID As Integer
    
    selID = CInt(chkAdv(Index).Tag)
    
    With EQProCtrl
        Select Case chkAdv(Index).Value
            Case Checked
                .SetAdvancedLinesValues selID, CLng(.AdvancedLinesRange(selID)(2))
            Case Unchecked
                .SetAdvancedLinesValues selID, CLng(.AdvancedLinesRange(selID)(1))
        End Select
    End With
        
End Sub

Private Sub chkInput_Click()

    'The user clicked?
    'Well.. let's tell EQPro...
    EQProCtrl.SelectForRecording = -chkInput.Value

End Sub

Private Sub chkMute_Click()

    'The user clicked?
    'Well.. let's tell EQPro...
    EQProCtrl.Mute = -chkMute.Value

End Sub

Private Sub cmbControls_Click()

    'Since VB does not support multi-threading we use a timer
    'to trigger the sub that will populate the mixer lines combo
    'and will create the necesary buttons for the advanced
    'mixer lines.
    tmrRefresh.Enabled = True
    
    'Since the buttons are dynamically loaded/unloaded the code
    'can not be included in this event... why? Ask Microsoft!

End Sub

Private Sub cmbMixerLines_Click()

    Dim SelLineID As Long
    
    'Let's store the line ID of the selected line in the combo
    SelLineID = cmbMixerLines.ItemData(cmbMixerLines.ListIndex)

    'if listindex=1 then we're viewing the Rec Controls
    'then lets prepare the buttons according to our view
    chkInput.Enabled = -cmbControls.ListIndex
    If Not chkInput.Enabled Then chkInput.Value = 0
    chkMute.Enabled = EQProCtrl.HasMute(SelLineID)
    If Not chkMute.Enabled Then chkMute.Value = 0
    
    'Set the line on which we're working
    EQProCtrl.ActiveLine = SelLineID

End Sub

Private Sub cmbMixers_Click()

    'Here we're telling the control which mixer we want
    'to work with
    EQProCtrl.ActiveMixer = cmbMixers.ListIndex
    'This selects the first item in the control and triggers
    'the Click event
    'This combo contains two items,
    ' 0 = Output Lines (dOutput)
    ' 1 = Rec/Input Lines (dinput)
    cmbControls.ListIndex = dOutput
    cmbControls_Click

End Sub

Private Sub cmdCenter_Click()

    EQProCtrl.Panning = 0

End Sub

Private Sub cmdClose_Click()

    '...duh?
    Unload Me

End Sub

Private Sub cmdQuery_Click()

    frmQuery.Show vbModal

End Sub

Private Sub Form_Load()

    Dim i As Integer

    'Let's center the dialog...
    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
    Caption = "EQPro " + EQProCtrl.Version + " Tester Application"
    
    With EQProCtrl
        'First of all we must START the control
        .IniEQ
        'Now let's see if the system has some mixer/sound card installed
        If .MixerCount Then
            'Let's walk through all the available mixers to
            'populate the Mixers combo.
            'Note that ListIndex property of the combo will be used
            'to address each mixer.
            For i = 0 To .MixerCount - 1
                cmbMixers.AddItem .MixerList(i)
            Next i
            .RefreshPriority = eqpHigh
            'This selects the first item in the combo and triggers
            'the Click event
            cmbMixers.ListIndex = 0
            'Now, we start a timer which will monitor the values
            'that the control does not refresh automaticly
            tmrMonitor.Enabled = True
        Else
            MsgBox "There're no mixers installed on this system... quiting!"
            End
        End If
    End With
        
End Sub

Private Sub sldPan_Click()

    sldPan_Scroll

End Sub

Private Sub sldPan_Scroll()

    EQProCtrl.Panning = sldPan.Value

End Sub

Private Sub tmrMonitor_Timer()

    Dim i As Integer

    'This timer is what keeps all the non-auto-refreshed controls
    'refreshed... don't you wish VB could support (easily) callbacks?!?!
    With EQProCtrl
        If .HasMute(.ActiveLine) Then
            chkMute.Value = Abs(EQProCtrl.Mute)
        End If
        chkInput.Value = Abs(EQProCtrl.SelectForRecording)
        
        'This For/Next cycle through all the advanced lines, gets
        'their values and assings it to each button
        'Please note that the index that corresponds to each line is
        'stored in the Tag property.
        For i = 0 To chkAdv.Count - 1
            chkAdv(i).Value = Abs(.AdvancedLinesValue(Val(chkAdv(i).Tag)) > 0)
        Next i
        
        sldPan.Value = .Panning
        
    End With
    
End Sub

Private Sub tmrRefresh_Timer()

    Dim i As Integer
    Dim cmdidx As Integer
    
    'Let's turn off the timer so this code doesn't call it self!
    tmrRefresh.Enabled = False
    
    'Here, we clean the controls, both the combo and
    'the buttons
    cmbMixerLines.Clear
    For i = chkAdv.Count - 1 To 1 Step -1
        Unload chkAdv(i)
    Next i
    'We can not unload the first control, so we hide it
    chkAdv(0).Visible = False

    With EQProCtrl
        Select Case cmbControls.ListIndex
            Case dOutput      'Show Output lines
                'This For/Next cycle will populate the combo
                'with the output lines detected
                For i = 0 To .OutputLineCount - 1
                    cmbMixerLines.AddItem .OutputLineList(i)
                    cmbMixerLines.ItemData(cmbMixerLines.NewIndex) = .OutputLineID(i)
                Next i
                'We will include in the same combo all the
                'detected advanced lines that are of type Fader
                'and we will control them using the EQPro's
                'integrated slider
                For i = 0 To .AdvancedLinesCount - 1
                    'Is the Advanced line for Output?
                    If .AdvancedLinesDirection(i) = dOutput Then
                        If .AdvancedLinesType(i) = ltFader Then
                            cmbMixerLines.AddItem .AdvancedLinesList(i)
                            cmbMixerLines.ItemData(cmbMixerLines.NewIndex) = .AdvancedLinesID(i)
                        Else
                            'If the advanced line #i is not of type fader
                            'we must create a button which will
                            'control the value of this line
                            If cmdidx > 0 Then
                                Load chkAdv(cmdidx)
                            End If
                            chkAdv(cmdidx).Visible = True
                            chkAdv(cmdidx).Container = frmAdvanced
                            chkAdv(cmdidx).Left = chkAdv(0).Left
                            chkAdv(cmdidx).Width = chkAdv(0).Width
                            chkAdv(cmdidx).Height = chkAdv(0).Height
                            chkAdv(cmdidx).Top = (.Height + 100) * cmdidx + 275
                            chkAdv(cmdidx).Caption = .AdvancedLinesList(i)
                            'This will become useful later...
                            chkAdv(cmdidx).Tag = i
                            cmdidx = cmdidx + 1
                            'Since we can not know how many controls will the mixer have
                            'because this will depend on the installed sound card and
                            'drivers version, we must dynamically create the appropiate
                            'controls.
                        End If
                    End If
                Next i
            Case dInput      'Show Input lines
                'This cycle populates the combo with the rec/input lines
                'and works in the same way as the above one
                For i = 0 To .InputLineCount - 1
                    cmbMixerLines.AddItem .InputLineList(i)
                    cmbMixerLines.ItemData(cmbMixerLines.NewIndex) = .InputLineID(i)
                Next i
                For i = 0 To .AdvancedLinesCount - 1
                    'Is the Advanced line for Input?
                    If .AdvancedLinesDirection(i) = dInput Then
                        If .AdvancedLinesType(i) = ltFader Then
                            cmbMixerLines.AddItem .AdvancedLinesList(i)
                            cmbMixerLines.ItemData(cmbMixerLines.NewIndex) = .AdvancedLinesID(i)
                        Else
                            'If the advanced line #i is not of type fader
                            'we must create a button which will
                            'control the value of this line
                            If cmdidx > 0 Then
                                Load chkAdv(cmdidx)
                            End If
                            chkAdv(cmdidx).Visible = True
                            chkAdv(cmdidx).Container = frmAdvanced
                            chkAdv(cmdidx).Left = chkAdv(0).Left
                            chkAdv(cmdidx).Width = chkAdv(0).Width
                            chkAdv(cmdidx).Height = chkAdv(0).Height
                            chkAdv(cmdidx).Top = (.Height + 100) * cmdidx + 275
                            chkAdv(cmdidx).Caption = .AdvancedLinesList(i)
                            'This will become useful later...
                            chkAdv(cmdidx).Tag = i
                            cmdidx = cmdidx + 1
                            'Since we can not know how many controls will the mixer have
                            'because this will depend on the installed sound card and
                            'drivers version, we must dynamically create the appropiate
                            'controls.
                        End If
                    End If
                Next i
        End Select
    End With
    
    'This triggers the Click event for this combo
    If cmbMixerLines.ListCount Then cmbMixerLines.ListIndex = 0
    
End Sub

