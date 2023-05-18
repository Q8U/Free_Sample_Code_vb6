VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cover Designer Pro 3.00"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -120
      ScaleHeight     =   495
      ScaleWidth      =   6975
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6240
      Width           =   6975
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
         Left            =   6960
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Top             =   120
         Width           =   200
      End
   End
   Begin VB.Timer tmrScroll 
      Interval        =   100
      Left            =   2280
      Top             =   5760
   End
   Begin VB.TextBox txtTips 
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
      Left            =   6960
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame fraFront_Inside_W 
      BackColor       =   &H80000018&
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblFrontW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Front Cover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   12
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label lblInsideW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inside Jacket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   1620
      End
      Begin VB.Image imgFront_Inside_W 
         BorderStyle     =   1  'Fixed Single
         Height          =   2880
         Left            =   480
         Stretch         =   -1  'True
         Top             =   720
         Width           =   5760
      End
   End
   Begin VB.Frame fraFront 
      BackColor       =   &H80000018&
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblFront 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Front Cover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1425
      End
      Begin VB.Image imgFront 
         BorderStyle     =   1  'Fixed Single
         Height          =   3720
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3720
      End
   End
   Begin VB.Frame fraFront_Inside_S 
      BackColor       =   &H80000018&
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lblInsideS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inside Jacket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblFrontS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Front Cover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   1425
      End
      Begin VB.Image imgFrontS 
         BorderStyle     =   1  'Fixed Single
         Height          =   2880
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2880
      End
      Begin VB.Image imgInsideS 
         BorderStyle     =   1  'Fixed Single
         Height          =   2880
         Left            =   480
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2880
      End
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   6000
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   4920
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   240
      Picture         =   "frmMain.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdFIW 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   3840
      Picture         =   "frmMain.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdFIS 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   3000
      Picture         =   "frmMain.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdFront 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   1320
      Picture         =   "frmMain.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000018&
      Height          =   555
      Left            =   2160
      Picture         =   "frmMain.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraBack 
      BackColor       =   &H80000018&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
      Begin VB.Label lblBack 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Cover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   1380
      End
      Begin VB.Image imgBack 
         BorderStyle     =   1  'Fixed Single
         Height          =   3600
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   720
         Width           =   4200
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6975
      Left            =   6840
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DefaultDir As String
Dim ScrollText As Integer

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

'//Command Button Code Starts Here
Private Sub cmdExit_Click()
Dim Response As Long

Response = MsgBox("Are you sure you want to exit" & vbCr & _
                    "Cover Designer Pro 3.00 ?", vbYesNo + vbInformation, "Cover Deisigner Pro")
If Response = vbYes Then
    SaveString HKEY_CURRENT_USER, _
    "Software\Cover Designer Pro 3\Options", "Default Path", DefaultDir
    
    Unload Me
    End
ElseIf Response = vbNo Then
'do nothing
End If
End Sub

Private Sub cmdFront_Click()
    HideALL
    fraFront.Visible = True
End Sub

Private Sub cmdBack_Click()
    HideALL
    fraBack.Visible = True
End Sub

Private Sub cmdFIS_Click()
    HideALL
    fraFront_Inside_S.Visible = True
End Sub

Private Sub cmdFIW_Click()
    HideALL
    fraFront_Inside_W.Visible = True
End Sub

Private Sub cmdPrint_Click()
    frmPrint.Show
End Sub
'//Command Button Code Ends Here


'//Form Code Ends Here
Private Sub Form_Activate()
Dim DefDir As String

'look for default path
DefDir = GetString(HKEY_CURRENT_USER, _
"Software\Cover Designer Pro 3\Options", "Default Path")

'if default path doesn't exist create it as C:\
If DefDir = "" Then
    SaveString HKEY_CURRENT_USER, _
    "Software\Cover Designer Pro 3\Options", "Default Path", "C:\"
    DefaultDir = "C:\"
End If

End Sub

Private Sub Form_Load()
    
'//This is part of the scrolling text code
ScrollText = FreeFile
Open App.Path & "\Scroll.txt" For Input As ScrollText
txtScroll = Input(LOF(ScrollText), ScrollText)
Close #ScrollText
    
    HideALL 'call sub HideAll
    PositionAll 'call sub PositionAll
    fraFront.Visible = True 'Make fraFront visible
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'when form unloads save currrent default dir
    'to registry
    SaveString HKEY_CURRENT_USER, _
    "Software\Cover Designer Pro 3\Options", "Default Path", DefaultDir
End Sub
'//Form Code Ends Here


'//Load Cover Code Starts Here
Private Sub imgFront_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OpenFileName As String
        
On Error GoTo ErrHandler 'if an error occurs goto ErrHandler

If Button = 1 Then
    With CD
        .CancelError = True 'if cancel is pressed this causes an error that can then be handled by ErrHandler
        .DialogTitle = "Select an Image"
        .Flags = cdlOFNHideReadOnly 'hides Open as read only checkbox
        .InitDir = DefaultDir 'sets default directory
        'Sets filters
        .Filter = "JPEG or JIFF Compliant (*.jpg *.jif *.jpeg)|*.JPG|Windows or OS/2 Bitmap (*.bmp)|*.BMP|CompuServe Graphics Interface (*.gif)|*.GIF|All files (*.*)|*.*"
        .FileName = "" 'Clears any text out of dialog box
        .ShowOpen 'loads open dialog box
    End With
        
        'takes the first 3 characters from OpenFileName
        'which is the drive letter i.e. C:\ or D:\ etc
        DefaultDir = Mid(CD.FileName, 1, 3)
        OpenFileName = CD.FileName 'sets variable as chosen filename
        MousePointer = 11 'arrow with egg timer mousepointer
        imgFront.Picture = LoadPicture(OpenFileName) 'loads selected file
        MousePointer = 0 'default mousepointer
       
        'prnFront = True 'sets boolean as true

ElseIf Button = 2 Then
    imgFront.Picture = LoadPicture("")
End If

Exit Sub
ErrHandler:
    imgFront.Picture = LoadPicture("")

End Sub


Private Sub imgBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OpenFileName As String
        
On Error GoTo ErrHandler 'if an error occurs goto ErrHandler

If Button = 1 Then
    With CD
        .CancelError = True 'if cancel is pressed this causes an error that can then be handled by ErrHandler
        .DialogTitle = "Select an Image"
        .Flags = cdlOFNHideReadOnly 'hides Open as read only checkbox
        .InitDir = DefaultDir 'sets default directory
        'Sets filters
        .Filter = "JPEG or JIFF Compliant (*.jpg *.jif *.jpeg)|*.JPG|Windows or OS/2 Bitmap (*.bmp)|*.BMP|CompuServe Graphics Interface (*.gif)|*.GIF|All files (*.*)|*.*"
        .FileName = "" 'Clears any text out of dialog box
        .ShowOpen 'loads open dialog box
    End With
        
        'takes the first 3 characters from OpenFileName
        'which is the drive letter i.e. C:\ or D:\ etc
        DefaultDir = Mid(OpenFileName, 1, 3)
        OpenFileName = CD.FileName 'sets variable as chosen filename
        MousePointer = 11 'arrow with egg timer mousepointer
        imgBack.Picture = LoadPicture(OpenFileName) 'loads selected file
        MousePointer = 0 'default mousepointer
       
        'prnFront = True 'sets boolean as true

ElseIf Button = 2 Then
    imgBack.Picture = LoadPicture("")
End If

Exit Sub
ErrHandler:
    imgBack.Picture = LoadPicture("")

End Sub


Private Sub imgFrontS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OpenFileName As String
        
On Error GoTo ErrHandler 'if an error occurs goto ErrHandler

If Button = 1 Then
    With CD
        .CancelError = True 'if cancel is pressed this causes an error that can then be handled by ErrHandler
        .DialogTitle = "Select an Image"
        .Flags = cdlOFNHideReadOnly 'hides Open as read only checkbox
        .InitDir = DefaultDir 'sets default directory
        'Sets filters
        .Filter = "JPEG or JIFF Compliant (*.jpg *.jif *.jpeg)|*.JPG|Windows or OS/2 Bitmap (*.bmp)|*.BMP|CompuServe Graphics Interface (*.gif)|*.GIF|All files (*.*)|*.*"
        .FileName = "" 'Clears any text out of dialog box
        .ShowOpen 'loads open dialog box
    End With
        
        'takes the first 3 characters from OpenFileName
        'which is the drive letter i.e. C:\ or D:\ etc
        DefaultDir = Mid(OpenFileName, 1, 3)
        OpenFileName = CD.FileName 'sets variable as chosen filename
        MousePointer = 11 'arrow with egg timer mousepointer
        imgFrontS.Picture = LoadPicture(OpenFileName) 'loads selected file
        MousePointer = 0 'default mousepointer
       
        'prnFront = True 'sets boolean as true

ElseIf Button = 2 Then
    imgFrontS.Picture = LoadPicture("")
End If

Exit Sub
ErrHandler:
    imgFrontS.Picture = LoadPicture("")

End Sub


Private Sub imgInsideS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OpenFileName As String
        
On Error GoTo ErrHandler 'if an error occurs goto ErrHandler

If Button = 1 Then
    With CD
        .CancelError = True 'if cancel is pressed this causes an error that can then be handled by ErrHandler
        .DialogTitle = "Select an Image"
        .Flags = cdlOFNHideReadOnly 'hides Open as read only checkbox
        .InitDir = DefaultDir 'sets default directory
        'Sets filters
        .Filter = "JPEG or JIFF Compliant (*.jpg *.jif *.jpeg)|*.JPG|Windows or OS/2 Bitmap (*.bmp)|*.BMP|CompuServe Graphics Interface (*.gif)|*.GIF|All files (*.*)|*.*"
        .FileName = "" 'Clears any text out of dialog box
        .ShowOpen 'loads open dialog box
    End With
        
        'takes the first 3 characters from OpenFileName
        'which is the drive letter i.e. C:\ or D:\ etc
        DefaultDir = Mid(OpenFileName, 1, 3)
        OpenFileName = CD.FileName 'sets variable as chosen filename
        MousePointer = 11 'arrow with egg timer mousepointer
        imgInsideS.Picture = LoadPicture(OpenFileName) 'loads selected file
        MousePointer = 0 'default mousepointer
       
        'prnFront = True 'sets boolean as true

ElseIf Button = 2 Then
    imgInsideS.Picture = LoadPicture("")
End If

Exit Sub
ErrHandler:
    imgInsideS.Picture = LoadPicture("")

End Sub


Private Sub imgFront_Inside_W_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OpenFileName As String
        
On Error GoTo ErrHandler 'if an error occurs goto ErrHandler

If Button = 1 Then
    With CD
        .CancelError = True 'if cancel is pressed this causes an error that can then be handled by ErrHandler
        .DialogTitle = "Select an Image"
        .Flags = cdlOFNHideReadOnly 'hides Open as read only checkbox
        .InitDir = DefaultDir 'sets default directory
        'Sets filters
        .Filter = "JPEG or JIFF Compliant (*.jpg *.jif *.jpeg)|*.JPG|Windows or OS/2 Bitmap (*.bmp)|*.BMP|CompuServe Graphics Interface (*.gif)|*.GIF|All files (*.*)|*.*"
        .FileName = "" 'Clears any text out of dialog box
        .ShowOpen 'loads open dialog box
    End With
        
        'takes the first 3 characters from OpenFileName
        'which is the drive letter i.e. C:\ or D:\ etc
        DefaultDir = Mid(OpenFileName, 1, 3)
        OpenFileName = CD.FileName 'sets variable as chosen filename
        MousePointer = 11 'arrow with egg timer mousepointer
        imgFront_Inside_W.Picture = LoadPicture(OpenFileName) 'loads selected file
        MousePointer = 0 'default mousepointer
       
        'prnFront = True 'sets boolean as true

ElseIf Button = 2 Then
    imgFront_Inside_W.Picture = LoadPicture("")
End If

Exit Sub
ErrHandler:
    imgFront_Inside_W.Picture = LoadPicture("")

End Sub
'//Load Cover code Ends Here


'//Tooltip code Starts here
'//Command Buttons
Private Sub cmdFIS_GotFocus()
    txtTips.Text = "Display Front Cover & Inside Jacket (Separate)"
End Sub

Private Sub cmdFIS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Display Front Cover & Inside Jacket (Separate)"
End Sub

Private Sub cmdFIW_GotFocus()
    txtTips.Text = "Display Front Cover & Inside Jacket (Whole)"
End Sub

Private Sub cmdFIW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Display Front Cover & Inside Jacket (Whole)"
End Sub

Private Sub cmdFront_GotFocus()
    txtTips.Text = "Display Front Cover"
End Sub

Private Sub cmdFront_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Display Front Cover"
End Sub

Private Sub cmdExit_GotFocus()
    txtTips.Text = "Close this fantastic program"
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Close this fantastic program"
End Sub

Private Sub cmdBack_GotFocus()
    txtTips.Text = "Display Back Cover"
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Display Back Cover"
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "View Print Details"
End Sub

Private Sub cmdPrint_GotFocus()
    txtTips.Text = "View Print Details"
End Sub

Private Sub cmdAbout_GotFocus()
    txtTips.Text = "About Cover Designer Pro 3.00"
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "About Cover Designer Pro 3.00"
End Sub
'//End Command Buttons

'//Frames
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = ""
End Sub

Private Sub fraBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = ""
End Sub

Private Sub fraFront_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = ""
End Sub

Private Sub fraFront_InsideS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = ""
End Sub

Private Sub fraFront_InsideW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = ""
End Sub
'//End frames

'//Images
Private Sub imgFront_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Use the Left Mouse Button to click here to load Front Cover, Right click to clear the box"
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Use the Left Mouse Button to click here to load Back Cover, Right click to clear the box"
End Sub

Private Sub imgFront_Inside_W_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Use the Left Mouse Button to click here to load Front Cover && Inside Jacket, Right click to clear the box"
End Sub

Private Sub imgFrontS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Use the Left Mouse Button to click here to load Front Cover, Right click to clear the box"
End Sub

Private Sub imgInsideS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Use the Left Mouse Button to click here to load Inside Jacket, Right click to clear the box"
End Sub
'//End Images
'//Tooltip code ends here


'//Other Sub Code Starts Here
Private Sub HideALL()
Dim ctl As Control

For Each ctl In Me.Controls
    If (TypeOf ctl Is Frame) Then
        ctl.Visible = False
    End If
Next
End Sub

Private Sub PositionAll()
Dim ctl As Control

For Each ctl In Me.Controls
    If (TypeOf ctl Is Frame) Then
        ctl.Top = 1080 '840
        ctl.Left = 120
        ctl.BorderStyle = 0
        'ctl.Height = 4215
        ctl.Width = 6375
    End If
Next
End Sub
'//Other Sub Code Ends Here


'//Text Box Code Starts Here
'The scrolling textboxes can't get focus, not for a major
'reason it just looks sloppy
Private Sub txtTips_GotFocus()
    cmdFront.SetFocus
End Sub

Private Sub txtTips_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTips.Text = "Tooltips Tips Appear Here"
End Sub

Private Sub txtScroll_GotFocus()
    cmdFront.SetFocus
End Sub
'//Text Box Code Starts Here


'//Scrolling Text Code Starts Here
Private Sub tmrScroll_Timer()
Dim Temp As Long

'This resizez text box to fit the text
Temp = Len(txtScroll.Text)
Temp = Temp * 105
txtScroll.Width = Temp
txtScroll.Left = txtScroll.Left - 40

If (txtScroll.Left + txtScroll.Width) < Picture1.Left Then
    txtScroll.Left = Picture1.ScaleWidth
End If
End Sub
'//Scrolling Text Code Ends Here
