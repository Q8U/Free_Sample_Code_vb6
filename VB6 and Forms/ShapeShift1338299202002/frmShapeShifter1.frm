VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ShapeShifter"
   ClientHeight    =   3195
   ClientLeft      =   8235
   ClientTop       =   5790
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Show the Form"
      Height          =   2415
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Your Own"
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Arrow-Head"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Star"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hand-Drawn Square"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "visit: http://www.geocities.com/campbell_donalde/vb"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "programmed by: Donald Evan Campbell"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' declare the array to store points.
Dim p(20) As POINTAPI

Private Sub Command1_Click()
' declare counter variables.
Dim counter As Integer

' ask the user for x,y for 13 pts.
For counter = 1 To 13 Step 1
    p(counter).x = InputBox("Please enter x value for point " & counter & ".")
    p(counter).y = InputBox("Please enter y value for point " & counter & ".")
Next counter

'tell user all values are finished.
MsgBox "All values are finished."
End Sub

Private Sub Command2_Click()
' create the shape.
' create the polygon region
' and set the window to abide by it.

SetWindowRgn hWnd, CreatePolygonRgn(p(1), 13, 1), True

' open the form from which users can exit.

Form2.Show 1
End Sub

Private Sub Option1_Click()
' assign points to create a square.
    p(1).x = 100
    p(1).y = 0
    
    p(2).x = 120
    p(2).y = 2.7
    
    p(3).x = 180
    p(3).y = 2.7
    
    p(4).x = 178.7
    p(4).y = 80
    
    p(5).x = 180
    p(5).y = 164.5
    
    p(6).x = 120
    p(6).y = 165.7
        
    p(7).x = 83.6
    p(7).y = 160
    
    p(8).x = 80
    p(8).y = 159.1
    
    p(9).x = 20
    p(9).y = 159.1
    
    p(10).x = 20.8
    p(10).y = 80
    
    p(11).x = 20
    p(11).y = 2.7
    
    p(12).x = 80
    p(12).y = 2.7
    
    p(13).x = 100
    p(13).y = 0
End Sub

Private Sub Option2_Click()
' assign points to create a star.
    p(1).x = 100
    p(1).y = 0
    
    p(2).x = 120
    p(2).y = 40
    
    p(3).x = 180
    p(3).y = 40
    
    p(4).x = 140
    p(4).y = 80
    
    p(5).x = 180
    p(5).y = 120
    
    p(6).x = 120
    p(6).y = 120
        
    p(7).x = 100
    p(7).y = 160
    
    p(8).x = 80
    p(8).y = 120
    
    p(9).x = 20
    p(9).y = 120
    
    p(10).x = 60
    p(10).y = 80
    
    p(11).x = 20
    p(11).y = 40
    
    p(12).x = 80
    p(12).y = 40
    
    p(13).x = 100
    p(13).y = 0
End Sub

Private Sub Option3_Click()
' assign points to create a triangle.
    p(1).x = 100
    p(1).y = 0
    
    p(2).x = 120
    p(2).y = 0.6
    
    p(3).x = 180
    p(3).y = -0.3
    
    p(4).x = 178.7
    p(4).y = -0.3
    
    p(5).x = 180
    p(5).y = -0.9
    
    p(6).x = 127.9
    p(6).y = 95.2
        
    p(7).x = 102.7
    p(7).y = 145.5
    
    p(8).x = 80
    p(8).y = 159.1
    
    p(9).x = 20
    p(9).y = 34.5
    
    p(10).x = 2.2
    p(10).y = 1
    
    p(11).x = 20
    p(11).y = -0.9
    
    p(12).x = 80
    p(12).y = -0.9
    
    p(13).x = 100
    p(13).y = 0
End Sub
