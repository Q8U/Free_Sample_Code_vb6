VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Details"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmPrint.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSCroll 
      Interval        =   50
      Left            =   2400
      Top             =   3240
   End
   Begin VB.TextBox txtTip 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   5520
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H80000018&
      Height          =   495
      Left            =   3240
      Picture         =   "frmPrint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      BackColor       =   &H80000018&
      Height          =   495
      Left            =   3240
      Picture         =   "frmPrint.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtCopies 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "1"
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5415
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4200
      Width           =   5415
      Begin VB.TextBox txtScroll 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdFIWandBack 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   360
      Picture         =   "frmPrint.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdFISandBack 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   360
      Picture         =   "frmPrint.frx":2CD0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdFrontandBack 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   360
      Picture         =   "frmPrint.frx":4D86
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   2760
      Picture         =   "frmPrint.frx":6078
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdFront 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   1800
      Picture         =   "frmPrint.frx":6382
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdFIS 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   3720
      Picture         =   "frmPrint.frx":668C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdFIW 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   4680
      Picture         =   "frmPrint.frx":6996
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   360
      Picture         =   "frmPrint.frx":6CA0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblCopies 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Number of Copies to Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4695
      Left            =   5400
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ScrollText As Integer

'//Number of copies command button code Starts Here
Private Sub cmdDown_Click()
Dim copies As Integer

'Adds 1 to copies to print
copies = txtCopies.Text

copies = copies - 1
If copies < 1 Then
    copies = 1
End If

txtCopies.Text = copies

End Sub

Private Sub cmdUp_Click()
Dim copies As Integer

'Subtracts 1 to copies to print
copies = txtCopies.Text

copies = copies + 1
If copies > 99 Then
    copies = 99
End If

txtCopies.Text = copies
End Sub
'//Number of copies to print command button code Ends Here



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim ctl As Control

OnTop = True

'//This disables all controls in frmMain to stop user
'//clicking anything when this form is actve
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = False
    End If
Next

Call SortPrint

'//This is part of the scrolling text code
ScrollText = FreeFile
Open App.Path & "\Scroll.txt" For Input As ScrollText
txtScroll = Input(LOF(ScrollText), ScrollText)
Close #ScrollText
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim ctl As Control

OnTop = True

'//This enabled all controls again
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = True
    End If
Next
End Sub


'//Tips code Starts Here
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = ""
End Sub

Private Sub cmdFront_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front Cover Only"
End Sub

Private Sub cmdFront_GotFocus()
    txtTip.Text = "Print Front Cover Only"
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Back Cover Only"
End Sub

Private Sub cmdBack_GotFocus()
    txtTip.Text = "Print Back Cover Only"
End Sub

Private Sub cmdFrontandBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front & Back Covers"
End Sub

Private Sub cmdFrontandBack_GotFocus()
    txtTip.Text = "Print Front & Back Covers"
End Sub

Private Sub cmdFIS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front Cover & Inside Jacket (Seperate) Only"
End Sub

Private Sub cmdFIS_GotFocus()
    txtTip.Text = "Print Front Cover & Inside Jacket (Seperate) Only"
End Sub

Private Sub cmdFIW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front Cover & Inside Jacket (Whole) Only"
End Sub

Private Sub cmdFIW_GotFocus()
    txtTip.Text = "Print Front Cover & Inside Jacket (Whole) Only"
End Sub

Private Sub cmdFISandBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front Cover, Inside Jacket (Seperate) & Back Cover"
End Sub

Private Sub cmdFISandBack_GotFocus()
    txtTip.Text = "Print Front Cover, Inside Jacket (Seperate) & Back Cover"
End Sub

Private Sub cmdFIWandBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Print Front Cover, Inside Jacket (Whole) & Back Cover"
End Sub

Private Sub cmdFIWandBack_GotFocus()
    txtTip.Text = "Print Front Cover, Inside Jacket (Whole) & Back Cover"
End Sub

Private Sub cmdUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Increment Copies to Print by 1. To increment Copies to Print by more than 1 enter the amount using the number keys on your keyboard"
End Sub

Private Sub cmdUp_GotFocus()
    txtTip.Text = "Increment Copies to Print by 1. To increment Copies to Print by more than 1 enter the amount using the numberical keys on your keyboard"
End Sub

Private Sub cmdDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Decrement Copies to Print by 1. To decrement Copies to Print by more than 1 enter the amount using the number keys on your keyboard"
End Sub

Private Sub cmdDown_GotFocus()
    txtTip.Text = "Decrement Copies to Print by 1. To decrement Copies to Print by more than 1 enter the amount using the number keys on your keyboard"
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Close the Print Details Screen"
End Sub

Private Sub cmdExit_GotFocus()
    txtTip.Text = "Close the Print Details Screen"
End Sub
'//Tip code Ends Here


'//Commandbutton Print Code Starts Here
Private Sub cmdFront_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s)?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintFront
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
'Do nothing
End If
End Sub

Private Sub cmdBack_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Back Cover(s)?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintBack
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub

Private Sub cmdFIS_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s) & Inside Jacket(s) ?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintFrontAndInside
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub

Private Sub cmdFIW_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s) & Inside Jacket(s) ?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintWhole
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub

Private Sub cmdFrontandBack_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s) & Back Cover(s) ?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintFrontAndBack
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub

Private Sub cmdFISandBack_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s), Inside Jacket(s) & Back Cover(s)?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintFrontsAndInsideSAndBack
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub

Private Sub cmdFIWandBack_Click()
Dim Response As Long

prnCopies = txtCopies.Text
OnTop = False 'without this you won#t see the msgbox

Response = MsgBox("Are you sure you want to print " & prnCopies & _
                " Front Cover(s), Inside Jacket(s) & Back Cover(s)?", vbYesNo + vbInformation, "Print Details")
If Response = vbYes Then
    Call PrintWholeAndBack
    OnTop = True 'we want to make this form topmost again
ElseIf Response = vbNo Then
    OnTop = True 'we want to make this form topmost again
End If
End Sub
'//Commandbutton Print Code Ends Here


'//+++++++++++++ Form OnTop Code
Private Property Let OnTop(Setting As Boolean)
    If Setting Then
        'make this form topmost
        SetWindowPos hwnd, HWND_TOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        'Make this form non-topmost
        SetWindowPos hwnd, HWND_NOTOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    mbOnTop = Setting
End Property

Private Property Get OnTop() As Boolean
    'Return the private variable set in property Let
    OnTop = mbOnTop
End Property


'//Sort print code Starts Here
Private Sub SortPrint()
Dim Front As Boolean
Dim Back As Boolean
Dim FIS As Boolean
Dim FIW As Boolean

With frmMain
    If .imgFront.Picture = LoadPicture("") Then
        cmdFront.Enabled = False
        Front = False
    Else
        cmdFront.Enabled = True
        Front = True
    End If
    
    If .imgBack.Picture = LoadPicture("") Then
        cmdBack.Enabled = False
        Back = False
    Else
        cmdBack.Enabled = True
        Back = True
    End If
    
    If .imgFrontS.Picture = LoadPicture("") Or .imgInsideS.Picture = LoadPicture("") Then
        cmdFIS.Enabled = False
        FIS = False
    Else
        cmdFIS.Enabled = True
        FIS = True
    End If
    
    If .imgFront_Inside_W.Picture = LoadPicture("") Then
        cmdFIW.Enabled = False
        FIW = False
    Else
        cmdFIW.Enabled = True
        FIW = True
    End If

    If Front = False Or Back = False Then
        cmdFrontandBack.Enabled = False
    Else
        cmdFrontandBack.Enabled = True
    End If
    
    If FIS = False Or Back = False Then
        cmdFISandBack.Enabled = False
    Else
        cmdFISandBack.Enabled = True
    End If
    
    If FIW = False Or Back = False Then
        cmdFIWandBack.Enabled = False
    Else
        cmdFIWandBack.Enabled = True
    End If

End With

End Sub
'//SortPrint code Ends Here


Private Sub tmrScroll_Timer()
Dim Temp As Long

'This resizez text box to fit the text
Temp = Len(txtScroll.Text)
Temp = Temp * 105
txtScroll.Width = Temp
txtScroll.Left = txtScroll.Left - 20

If (txtScroll.Left + txtScroll.Width) < picScroll.Left Then
    txtScroll.Left = picScroll.ScaleWidth
End If

End Sub

'//Text Box Code Starts Here
'The scrolling textboxes can't get focus, not for a major
'reason it just looks sloppy
Private Sub txtTips_GotFocus()
    cmdUp.SetFocus
End Sub

Private Sub txtTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTip.Text = "Tooltips Tips Appear Here"
End Sub

Private Sub txtTip_GotFocus()
    cmdUp.SetFocus
End Sub

Private Sub txtScroll_GotFocus()
    cmdUp.SetFocus
End Sub
'//Text Box Code Starts Here
