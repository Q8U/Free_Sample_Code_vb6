VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  jakthebug..........CD JACKET    "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   8160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":1CFA
   MousePointer    =   99  'Custom
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "PRINT"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      ItemData        =   "Form1.frx":2004
      Left            =   6240
      List            =   "Form1.frx":2011
      TabIndex        =   4
      Text            =   " Choose Cover"
      ToolTipText     =   "  Pick cover Template"
      Top             =   4080
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CMN_Picture 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      Picture         =   "Form1.frx":2038
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      Picture         =   "Form1.frx":8103A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "               P         R          I          N           T        "
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
      Left            =   0
      MouseIcon       =   "Form1.frx":10003C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Covers Make Sure You Have An Image In Both Boxs"
      Top             =   3960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "PRINT COVERS    "
      Top             =   3960
      Width           =   6135
      ForeColor       =   65280
      BackColor       =   -2147483630
      Caption         =   "P          R              I                   N                     T"
      Size            =   "10821;873"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   120
      MouseIcon       =   "Form1.frx":100346
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":100650
      Stretch         =   -1  'True
      ToolTipText     =   "Click to Load Image"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   3600
      MouseIcon       =   "Form1.frx":10BEBC
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":10C1C6
      Stretch         =   -1  'True
      ToolTipText     =   "Click to Load Image"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   32
      X2              =   32
      Y1              =   96
      Y2              =   98
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CheckBox1_Click()
End Sub

Private Sub Combo1_Click()
Dim a
Dim b
Dim c
j = "Two Fronts"
a = "Full Cover"
k = "One Front"
If Combo1 = j Then
Image2.Left = 312
Image2.Visible = True
Image2.Width = 225
Command4.Visible = True
CommandButton1.Visible = False
Command1.Visible = False

End If
If Combo1 = a Then

CommandButton1.Visible = True
Command4.Visible = False
Command1.Visible = False
Image2.Visible = True
Image2.Left = 240

Image2.Width = 297
End If
If Combo1 = k Then
Image2.Visible = False
Command1.Visible = True
Command4.Visible = False
CommandButton1.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim ImageLeft, ImageTop, imagewidth, imageheight As Single
    
    On Error Resume Next
    
    Answer = MsgBox("              Please Confirm Printing on " & Printer.DeviceName, vbYesNo)
    If Answer = vbNo Then Exit Sub
    
    Printer.ScaleMode = vbCentimeters
    
        
    Printer.Print "";
    
        
    Picture2.ScaleMode = vbCentimeters
    Picture2.AutoSize = True
    Picture2.Refresh
    Picture2.AutoSize = False
    
    ImageLeft = 4.1
    
    ImageTop = 2.1
    imageheight = 12.1
    imagewidth = 12.1
    
    Printer.ScaleMode = vbCentimeters
    Printer.PaintPicture Picture2.Picture, ImageLeft, ImageTop, imageheight, imagewidth
     Printer.EndDoc
     
End Sub

Private Sub Command4_Click()
    
    Dim ImageLeft, ImageTop, imagewidth, imageheight As Single
    
    On Error Resume Next
    
    Answer = MsgBox("              Please Confirm Printing on " & Printer.DeviceName, vbYesNo)
    If Answer = vbNo Then Exit Sub
    
    Printer.ScaleMode = vbCentimeters
    
        
    Printer.Print "";
    
        
    Picture2.ScaleMode = vbCentimeters
    Picture2.AutoSize = True
    Picture2.Refresh
    Picture2.AutoSize = False
    
    ImageLeft = 4.1
    
    ImageTop = 2.1
    imageheight = 12.1
    imagewidth = 12.1
    
    Printer.ScaleMode = vbCentimeters
    Printer.PaintPicture Picture2.Picture, ImageLeft, ImageTop, imageheight, imagewidth
    
     
    
    Picture1.ScaleMode = vbCentimeters
    Picture1.AutoSize = True
    Picture1.Refresh
    Picture1.AutoSize = False
    
    ImageLeft = 4.1
    
    ImageTop = 15.2
    
    imageheight = 12.1
    
    imagewidth = 12.1
    
    Printer.ScaleMode = vbCentimeters
    
    Printer.PaintPicture Picture1.Picture, ImageLeft, ImageTop, imageheight, imagewidth
    
    Printer.EndDoc
End Sub
    
    
Private Sub CommandButton1_Click()
 Dim ImageLeft, ImageTop, imagewidth, imageheight As Single
    
    On Error Resume Next
    
    Answer = MsgBox("              Please Confirm Printing on " & Printer.DeviceName, vbYesNo)
    If Answer = vbNo Then Exit Sub
    
    Printer.ScaleMode = vbCentimeters
    
        
    Printer.Print "";
    
        
    Picture2.ScaleMode = vbCentimeters
    Picture2.AutoSize = True
    Picture2.Refresh
    Picture2.AutoSize = False
    
    ImageLeft = 4.1
    
    ImageTop = 2.1
    imageheight = 12.1
    imagewidth = 12.1
    
    Printer.ScaleMode = vbCentimeters
    Printer.PaintPicture Picture2.Picture, ImageLeft, ImageTop, imageheight, imagewidth
    
     
    
    Picture1.ScaleMode = vbCentimeters
    Picture1.AutoSize = True
    Picture1.Refresh
    Picture1.AutoSize = False
    
    ImageLeft = 2.6
    
    ImageTop = 15.2
    
    imageheight = 15
    
    imagewidth = 12
    
    Printer.ScaleMode = vbCentimeters
    
    Printer.PaintPicture Picture1.Picture, ImageLeft, ImageTop, imageheight, imagewidth
    
    Printer.EndDoc
End Sub

Private Sub Image1_Click()
    On Error GoTo cancelled
    CMN_Picture.CancelError = True
    CMN_Picture.DialogTitle = "CD JACKET        link2me80@hotmail.com"
    CMN_Picture.Filter = "Cd Jacket  bmp,jpg |*.bmp;*.jpg|(*.*)|*.*|"
    CMN_Picture.ShowOpen
    Image1.Picture = LoadPicture(CMN_Picture.FileName)
    Picture2.Picture = Image1.Picture
cancelled:
End Sub
    
Private Sub Image2_Click()
    On Error GoTo cancelled
    Cd1.DialogTitle = "CD JACKET        link2me80@hotmail.com"
    Cd1.CancelError = True
    Cd1.Filter = " Cd Jacket  bmp, jpg|*.bmp;*.jpg|(*.*)|*.*|"
    Cd1.ShowOpen
    
    Image2.Picture = LoadPicture(Cd1.FileName)
    Picture1.Picture = Image2.Picture
cancelled:
End Sub
    
    
