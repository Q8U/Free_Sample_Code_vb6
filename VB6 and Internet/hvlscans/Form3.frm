VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   795
   ClientLeft      =   9405
   ClientTop       =   2190
   ClientWidth     =   2370
   LinkTopic       =   "Form3"
   ScaleHeight     =   795
   ScaleWidth      =   2370
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Timer1.Enabled = False
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
Winsock1.Close
Form2.List1.AddItem Text1.Text & ":" & Text2.Text
Form2.List2.AddItem "Connection Found " & Text1.Text
If Form2.Check1.Value = 1 Then SendChat "  · • • ¤ {HVL]•¤•  Found Connection @ " & Text1.Text & " Port " & Text2.Text
Form2.opens.Text = Val(Form2.opens.Text) - 1
Unload Me
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.Close
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
Form2.opens.Text = Val(Form2.opens.Text) - 1
Unload Me
End Sub
