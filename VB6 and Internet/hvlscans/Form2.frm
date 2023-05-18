VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HVL High-Speed IP Scanner - By LoST DaTa"
   ClientHeight    =   2025
   ClientLeft      =   3870
   ClientTop       =   2160
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "Attack"
      Height          =   195
      Left            =   3240
      TabIndex        =   25
      Top             =   1680
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Scan"
      Height          =   195
      Left            =   2520
      TabIndex        =   24
      Top             =   1680
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7680
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Interval        =   300
      Left            =   6480
      Top             =   240
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   23
      Text            =   "0"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   7680
      Top             =   120
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   21
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "200"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox opens 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   18
      Text            =   "0"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox stop 
      Height          =   285
      Left            =   7200
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unload"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   480
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Report Finds in Chat Room?"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox done1 
      Height          =   285
      Left            =   7200
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Text            =   "c:\windows\desktop\Serverscan.txt"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   2040
      TabIndex        =   10
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "12345"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Scanned"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Delay:"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Open Sockets"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Log:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public t2
Public t1

Private Sub Command1_Click()
On Error Resume Next
If Timer2.Enabled = True Then
Timer2.Enabled = False
Command1.Caption = "Scan"
t2 = Timer
ti = t2 - t1
SendChat "  · • • ¤ {HVL}•¤•  Scanned " & Text8.Text & " at " & Format(Val(Text8.Text / ti), "##.#") & " Svrs/s"
SendChat "  · • • ¤ {HVL}•¤•  Found " & List1.ListCount & " Active Servers"
List2.AddItem "Scan Stopped"
List2.AddItem "Scan Speed " & Format(Val(Text8.Text / ti), "##.#") & " Servers/Second"
Else
Command1.Caption = "Stop"
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text7.Text = "" Then
MsgBox "Please fill out the form Compleatly": Exit Sub
End If
If Text7.Text = "0" Then Text7.Text = "1"
If Text1.Text = "0" Then MsgBox "Please Enter a Valid IP Address": Exit Sub
t1 = Timer
Text8.Text = "0"
List2.AddItem "Scan Started"
Timer2.Enabled = True
SendChat "  · • • ¤ {HVL}•¤•  Scanning From " & Form2.Text1.Text & "." & Form2.Text2.Text & "." & Form2.Text3.Text & "." & Form2.Text4.Text & " Port: " & Text5.Text
End If
End Sub

Private Sub Command2_Click()
MsgBox "This is fairly Simple to use, Basically just put in the address that you want to start scanning at  (i.e. 152.11.202.41), and the port to look for connections on those servers then hit scan. So far this has had 103 open connections at one time. This will also automaticly traverse Up C Class, B Class and A Class IPs as well as standard D Class IPs. It is also suggested that you run the least amount of TCP/IP connections as you can, this eats system resources like hell when scanning. I sugest you use the Delay wisely or keep it it is currnt setting if you don't know much about winsock."
End Sub

Private Sub Command3_Click()
Timer2.Enabled = False
Unload Me
End Sub

Private Sub Form_Activate()
If done1.Text = "1" Then Exit Sub
On Error GoTo errorz:
List2.AddItem "Int Sockets"
For del = 1 To 255
List1.AddItem del & "." & del & "." & del & "." & del & ":88888"
DoEvents
Next del
For del = 1 To 255
List1.RemoveItem 0
DoEvents
Next
List2.AddItem "Control INT Complete"
Socket1.AddressFamily = 2
Socket1.SocketType = 1
Socket1.Protocol = 6
Socket1.RemotePort = 7
Socket1.HostAddress = "127.0.0.1"
opens.Text = Val(opens.Text) + 1
Socket1.Action = 2
done1.Text = "1"
Exit Sub
errorz:
List2.AddItem "Error In Socket INT"
End Sub

Private Sub Form_Load()
SendChat "  · • • ¤ {HVL]•¤•  HVL High Speed IP Scanner Loaded"
DoEvents
SendChat "  · • • ¤ {HVL]•¤•  Coded By LoST DaTa"
DoEvents
SendChatlink "  · • • ¤ {HVL]•¤•  "
StayOnTop Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SendChat "  · • • ¤ {HVL]•¤•  HVL High Speed IP Scanner Unloaded"
End
End Sub

Private Sub Socket1_Connect()
List2.AddItem "Socket INT Complete"
Socket1.Action = 7
Socket1.Action = 6
opens.Text = Val(opens.Text) - 1
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
List2.AddItem ErrorString
List2.AddItem "Socket INT Complete"
Socket1.Action = 7
Socket1.Action = 6
opens.Text = Val(opens.Text) - 1
End Sub

Private Sub Timer1_Timer()
If List2.ListCount > 4 Then List2.RemoveItem 0
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
opens.Text = Val(opens.Text) + 1
Text8.Text = Val(Text8.Text) + 1
Dim foo As New Form3
IP = Form2.Text1.Text & "." & Form2.Text2.Text & "." & Form2.Text3.Text & "." & Form2.Text4.Text
Port = Form2.Text5.Text
foo.Text1.Text = IP
foo.Text2.Text = Port
'foo.Winsock1.RemoteHostIP IP
'foo.Winsock1.RemotePort Port
foo.Winsock1.Connect IP, Port
'foo.Command1.Value = True
If Option2.Value = True Then GoTo skip:
If Text4.Text = "255" Then
    If Text3.Text = 255 Then
        If Text2.Text = "255" Then
            If Text1.Text = 255 Then
                List2.AddItem "Maximum IP reached!"
                Timer2.Enabled = False
             End If
        Text2.Text = "0"
        Text1.Text = Val(Text1.Text) + 1
        Else
      '  Text2.Text = Val(Text2.Text) + 1
        End If
    Text3.Text = "0"
        Text2.Text = Val(Text2.Text) + 1
        Else
     '   Text3.Text = Val(Text3.Text) + 1
    End If
Text4.Text = "0"
        Text3.Text = Val(Text3.Text) + 1
        Else
'        Text5.Text = Val(Text5.Text) + 1
Text4.Text = Val(Text4.Text) + 1
End If
skip:
End Sub

Private Sub Timer3_Timer()
If Text7.Text = "" Then Text7.Text = "1"
If Text7.Text = "0" Then Text7.Text = "1"
Timer2.Interval = Text7.Text
End Sub

Private Sub Timer4_Timer()
Text9.Text = opens.Text
End Sub
