VERSION 5.00
Object = "*\AvbpScrllngFrm.vbp"
Begin VB.Form frmExample 
   Caption         =   "Example"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   76
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton Command3 
         Height          =   300
         Left            =   1200
         Picture         =   "frmExample.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Next >>"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   840
         Picture         =   "frmExample.frx":0188
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Last Page >>|"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   480
         Picture         =   "frmExample.frx":01CE
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "|<< First Page"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   120
         Picture         =   "frmExample.frx":0213
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "<< Previous"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0 of 0"
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
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   120
         Width           =   1215
      End
   End
   Begin vbpScrllngFrm.ScrllngFrm ScrllngFrm1 
      Height          =   2295
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      BackPicture     =   "frmExample.frx":0251
      BackColor       =   14737632
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Form"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   3375
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   1080
               TabIndex        =   0
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   1080
               TabIndex        =   3
               Top             =   2160
               Width           =   1455
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1080
               TabIndex        =   2
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   1080
               TabIndex        =   1
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   75
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Address: "
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   74
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Phone:"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   73
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "E-mail:"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   72
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Personal Information:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   71
               Top             =   360
               Width           =   2535
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Information:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Website:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   60
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Address: "
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Business:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton Command7 
            Caption         =   "Poor"
            Height          =   375
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Good"
            Height          =   375
            Index           =   2
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Very good"
            Height          =   375
            Index           =   3
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Excellent"
            Height          =   375
            Index           =   1
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Select an option:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Poor:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Good:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   65
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Very good: "
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Excellent:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   3375
            Begin VB.TextBox Text10 
               Height          =   1125
               Left            =   240
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   600
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Comments:"
               Height          =   255
               Left            =   240
               TabIndex        =   48
               Top             =   360
               Width           =   1335
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   0
         ScaleHeight     =   6615
         ScaleWidth      =   3735
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton Command10 
            Caption         =   "Command3"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   5640
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option3"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   5880
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option4"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   34
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option5"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   35
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option6"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   36
            Top             =   5640
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option7"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   37
            Top             =   5880
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option8"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   38
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option9"
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   39
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option10"
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   40
            Top             =   5640
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option11"
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   41
            Top             =   5880
            Width           =   1035
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame1"
            Height          =   975
            Left            =   120
            TabIndex        =   51
            Top             =   4080
            Width           =   3015
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   120
               TabIndex        =   29
               Text            =   "Text10"
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            Text            =   "Text9"
            Top             =   3600
            Width           =   1815
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Text            =   "Text3"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   840
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   3120
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3120
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Text            =   "Text8"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Text            =   "Text7"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Text            =   "Text6"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Text            =   "Text5"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   1800
            TabIndex        =   20
            Text            =   "Text4"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command121 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Command2"
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command49 
            Caption         =   "Command4"
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1080
            Width           =   1575
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            Picture         =   "frmExample.frx":2289
            ScaleHeight     =   225
            ScaleWidth      =   240
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   3600
            Width           =   240
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Aditional Information:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   255
            Left            =   600
            TabIndex        =   52
            Top             =   3600
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Go To >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   2520
      Picture         =   "frmExample.frx":2638
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Delete Page"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   1920
      Picture         =   "frmExample.frx":26BA
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Add Page"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   2760
      TabIndex        =   56
      Text            =   "1"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vote for this code >>"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   4680
      Top             =   3000
   End
   Begin VB.Label Label9 
      Caption         =   $"frmExample.frx":2796
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   69
      Top             =   3480
      Width           =   4575
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Add first four pages...
Private Sub Form_Load()
    ScrllngFrm1.AddPage Picture1
    ScrllngFrm1.AddPage Picture2
    ScrllngFrm1.AddPage Picture3
    ScrllngFrm1.AddPage Picture4
    
End Sub

'================================
'==== Navigate through pages ====
'================================

'Go to Previous Page...
Private Sub Command1_Click()
    ScrllngFrm1.PreviousPage
    
End Sub

'Go to First Page...
Private Sub Command2_Click()
    ScrllngFrm1.FirstPage
    
End Sub

'Go to Last Page...
Private Sub Command5_Click()
    ScrllngFrm1.LastPage
    
End Sub

'Go to Next Page...
Private Sub Command3_Click()
    ScrllngFrm1.NextPage
    
End Sub

'Go to page number displayed on TextBox...
Private Sub Command6_Click()
    ScrllngFrm1.CurrentPage = Text9.Text
    
End Sub

'====================
'==== Edit pages ====
'====================

'Add page 5...
Private Sub Command8_Click()
    ScrllngFrm1.AddPage Picture5
    
End Sub

'Delete current page...
Private Sub Command9_Click()
    Call ScrllngFrm1.DeletePage(ScrllngFrm1.CurrentPage)
    
End Sub

'=====================================
'==== Open Default Browser on     ====
'==== Planet-Source-Code to vote. ====
'=====================================

Private Sub Command12_Click()
    'The API that allows me to open the browser
    'in shell mode is on the modShell Module.
    Call Shell("cmd /c start http://www.planet-source-code.com/vb/default.asp?lngCId=32374&lngWId=1")
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    'I had to call this API a second time
    'because, some times, PSC opens a default
    'page instead of the page with my submission...
    Call Shell("cmd /c start http://www.planet-source-code.com/vb/default.asp?lngCId=32374&lngWId=1")
    
End Sub

'Put all the objects on proper
'place when resizing.
Private Sub Form_Resize()
    Dim intTempTop
    Dim intTempSpace
    
    On Error Resume Next
    
    'Prevent user from resizing the
    'Form into a size too small.
    If (Me.Height < 1500) Then
        Me.Height = 1500
    End If
    
    intTempTop = 600
    intTempSpace = ScrllngFrm1.Top + ScrllngFrm1.Height + intTempTop
    
    ScrllngFrm1.Width = Me.Width - 300
    ScrllngFrm1.Height = Me.Height - (2000 + intTempTop)
    
    Frame4.Top = ScrllngFrm1.Top + ScrllngFrm1.Height
    
    Command9.Top = ScrllngFrm1.Top + ScrllngFrm1.Height + 220
    Command8.Top = ScrllngFrm1.Top + ScrllngFrm1.Height + 220
    
    Command6.Top = Command8.Top + Command8.Height
    Text9.Top = Command8.Top + Command8.Height
    
    Command12.Top = ScrllngFrm1.Top + ScrllngFrm1.Height + 230
    
    Label9.Top = ScrllngFrm1.Top + ScrllngFrm1.Height + 900
End Sub

'On PageChanged event, update current
'page number, total of pages and if
'Navigation buttons should be enabled
'and/or visible.
Private Sub ScrllngFrm1_PageChanged()
    Label1.Caption = ScrllngFrm1.CurrentPage & " of " & ScrllngFrm1.HowManyPages
    Text9.Text = ScrllngFrm1.CurrentPage
    
    Command1.Enabled = ScrllngFrm1.PreviousEnabled
    Command2.Enabled = ScrllngFrm1.PreviousEnabled
    Command5.Enabled = ScrllngFrm1.NextEnabled
    Command3.Enabled = ScrllngFrm1.NextEnabled
    
    If (ScrllngFrm1.HowManyPages < 2) Then
        Command1.Visible = False
        Command2.Visible = False
        Command5.Visible = False
        Command3.Visible = False
    Else
        Command1.Visible = True
        Command2.Visible = True
        Command5.Visible = True
        Command3.Visible = True
    End If
    
End Sub
