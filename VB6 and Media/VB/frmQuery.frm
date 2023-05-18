VERSION 5.00
Begin VB.Form frmQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Mixer"
   ClientHeight    =   2895
   ClientLeft      =   5175
   ClientTop       =   5100
   ClientWidth     =   5010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   45
      TabIndex        =   6
      Top             =   30
      Width           =   4920
      Begin VB.CheckBox chkHold 
         Caption         =   "Hold"
         Height          =   210
         Left            =   3195
         TabIndex        =   10
         Top             =   1290
         Width           =   840
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   330
         Left            =   3180
         TabIndex        =   9
         Top             =   855
         Width           =   885
      End
      Begin VB.ComboBox cmbReturn 
         Height          =   330
         ItemData        =   "frmQuery.frx":0000
         Left            =   135
         List            =   "frmQuery.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   1710
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "&Query"
         Height          =   330
         Left            =   3180
         TabIndex        =   2
         Top             =   435
         Width           =   885
      End
      Begin VB.TextBox txtQuery 
         Height          =   315
         Left            =   135
         TabIndex        =   0
         Top             =   435
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Return"
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   870
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Query String"
         Height          =   210
         Left            =   135
         TabIndex        =   7
         Top             =   165
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4050
      TabIndex        =   3
      Top             =   2445
      Width           =   900
   End
   Begin VB.TextBox txtResult 
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2940
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   1740
      Width           =   495
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This dialog is designed to illustrate the functionality of the
'FindLine method.
'
'The FindLine method takes two parameters,
'ReturnType: Which indicates what information we want from the line
'LineName: This is the name, or part of the name of the line we're looking for.
'          The first time we call the method it should be called with both parameters.
'          Subsequential calls may omit the LineName parameter, and the method will
'          return the next matching line.
'          FineLine will return an empty string when no more lines match and the ReturnType
'          argument is set rtName.

'If you're looking for the Microphone line, for example, you should do this:
'
' Dim FullName as String
' Dim MicLineID as Long
'
' FullName = EqProCtrl.FindLine(rtName, "mic")
' If FullName<>"" then
'   MicLineID = EqProCtrl.FindLine(rtID, FullName)
' else
'   msgbox "No MIC was found!!!"
' end if
'
' EqProCtrl.ActiveLine = MicLineID

Option Explicit

Private Sub chkHold_Click()

    cmdNext.Enabled = Not -chkHold.Value

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdNext_Click()

    AnalyzeResult frmMain.EQProCtrl.FindLines(cmbReturn.ListIndex)

End Sub

Private Sub cmdQuery_Click()

    AnalyzeResult frmMain.EQProCtrl.FindLines( _
                                    cmbReturn.ListIndex, _
                                    txtQuery.Text, _
                                    -chkHold.Value)

End Sub

Private Sub AnalyzeResult(Ret As Variant)

    'This simple function analyzes the result of the query
    'and makes the out readable
    Select Case cmbReturn.ListIndex
        Case rtName, rtID
            txtResult.Text = Ret
        Case rtDirection
            Select Case Ret
                Case dOutput
                    txtResult.Text = "Output"
                Case dInput
                    txtResult.Text = "Input"
            End Select
        Case rtType
            Select Case Ret
                Case ltFader
                    txtResult.Text = "Fader"
                Case ltSwitch
                    txtResult.Text = "Switch?"
            End Select
    End Select

End Sub

Private Sub Form_Load()

    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With
    
    cmbReturn.ListIndex = 0
        
End Sub


