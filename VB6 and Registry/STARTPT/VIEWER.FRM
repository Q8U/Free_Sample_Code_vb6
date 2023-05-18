VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Viewer 
   Caption         =   "Viewer"
   ClientHeight    =   7320
   ClientLeft      =   255
   ClientTop       =   1050
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8700
   Begin VB.VScrollBar VScroll1 
      Height          =   6855
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.PictureBox Container 
      AutoRedraw      =   -1  'True
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.PictureBox Document 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   6855
         Left            =   0
         ScaleHeight     =   6795
         ScaleWidth      =   7635
         TabIndex        =   1
         Top             =   0
         Width           =   7695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Graphics"
      Filter          =   "*.bmp;*.gif;*.jpg;*.wmf"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Document Layout"
         Index           =   0
         Begin VB.Menu mnuLayout 
            Caption         =   "&Portrait"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLayout 
            Caption         =   "&Landscape"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Print"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "P&rint Setup"
         Index           =   2
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Print PreVie&w"
         Index           =   3
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Save"
         Index           =   4
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DocLayOut As Integer
Public Sub SizeDoc(DLayout As Integer)
    Select Case DLayout
        Case 1
            Document.Width = 8.5 * 1440
            Document.Height = 11 * 1440
        Case 2
            Document.Width = 11 * 1440
            Document.Height = 8.5 * 1440
    End Select
    SizeScrolls
End Sub
Public Sub PView()
    Dim SHeight As Single, SWidth As Single
    Container.Cls
    'Computes the size ratio of the document to display it proportionly within
    'the container
    If Document.Height > Document.Width Then
        SWidth = (Document.Width / Document.Height) * Container.ScaleHeight
        SHeight = Container.ScaleHeight
    Else
        SHeight = (Document.Height / Document.Width) * Container.ScaleHeight
        SWidth = Container.ScaleWidth
    End If
    'The document is not visible. Print on the container box
    Container.PaintPicture Document.Image, 0, 0, SWidth, SHeight
End Sub
Public Sub SizeScrolls()
    With VScroll1
        .Left = Container.Left + Container.Width
        .Top = Container.Top
        .Max = Document.Height - Container.ScaleHeight '32,767
        .Min = MTOP
        .Value = .Min
        .Height = Container.Height
        .SmallChange = Container.Height / 10    '1/10 of the container height
        .LargeChange = Container.Height
    End With
    If Document.ScaleHeight > Container.ScaleHeight Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    With HScroll1
        .Left = Container.Left
        .Top = Container.Top + Container.Height
        .Min = MLEFT
        .Width = Container.Width
        .Value = .Min
        .Max = Document.Width - Container.ScaleWidth
        .SmallChange = Container.ScaleWidth / 10
        .LargeChange = Container.Width
    End With
    If Document.ScaleWidth > Container.ScaleWidth Then
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Document.ScaleMode = vbTwips
    SizeDoc Printer.Orientation
    Me.Show
End Sub

Private Sub Form_Resize()
    'This code will insure the viewer and scroll bars look the same regardless
    'the users screen resolution.
    'Resize if the form left is off the screen
    'If Me.Left + Me.Width > Screen.Width Then Me.Width = Screen.Width - Me.Left
    'Size the container to the screen size(always in TWIPS).
    Container.Width = Me.Width - ((Container.Left * 2) + VScroll1.Width)
    Container.Height = Me.ScaleHeight - ((Container.Top * 2) + HScroll1.Height)
    SizeScrolls
End Sub

Private Sub HScroll1_Change()
    Document.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Select Case mnuFileSub(Index).Caption
        Case "&Save"
            CommonDialog1.ShowSave
            If Not CommonDialog1.CancelError And CommonDialog1.FileName = "" Then SavePicture Document.Image, CommonDialog1.FileName
        Case "&Print"
            Document.Visible = True
            Printer.Orientation = DocLayOut
            Printer.PaintPicture Document.Image, 0, 0
            Printer.EndDoc
        Case "Print PreVie&w"
            Document.Visible = False
            HScroll1.Visible = False
            VScroll1.Visible = False
            PView
            mnuFileSub(Index).Caption = "Close Preview"
        Case "Close Preview"
            Container.Cls
            Document.Visible = True
            mnuFileSub(Index).Caption = "Print PreVie&w"
            SizeScrolls
        Case "P&rint Setup"
            CommonDialog1.Flags = cdlPDPrintSetup
            CommonDialog1.ShowPrinter
            DoEvents
        Case "E&xit"
            End
    End Select
End Sub

Private Sub mnuLayout_Click(Index As Integer)
    For I = 0 To mnuLayout.Count - 1
            mnuLayout(I).Checked = False
    Next
    mnuLayout(Index).Checked = True
    DocLayOut = Index + 1
    SizeDoc DocLayOut
    If Document.Visible = False Then
        HScroll1.Visible = False
        VScroll1.Visible = False
        PView
    End If
End Sub

Private Sub VScroll1_Change()
    Document.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
