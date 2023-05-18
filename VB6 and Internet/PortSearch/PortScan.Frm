VERSION 4.00
Begin VB.Form frmPortSearch 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Searcher by ^Matt"
   ClientHeight    =   4185
   ClientLeft      =   1785
   ClientTop       =   1725
   ClientWidth     =   6315
   Height          =   4875
   Icon            =   "PortScan.frx":0000
   Left            =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6315
   Top             =   1095
   Width           =   6435
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtOutput 
      Height          =   3075
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1080
      Width           =   6285
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stop"
      Height          =   465
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   945
   End
   Begin VB.TextBox txtRemotePort 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Width           =   945
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Start"
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the IP number that you want to find the port of. Then click start."
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   3615
   End
   Begin WINSOCKLib.TCP TCP 
      Index           =   0
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      Blocking        =   0   'False
      SleepTime       =   10
      RemoteHost      =   ""
      RemotePort      =   0
      LocalPort       =   0
   End
   Begin VB.Menu mnu1 
      Caption         =   "File"
      Begin VB.Menu mnu2 
         Caption         =   "Start"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnu4 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPortSearch"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' Copyright  (c) 1996 NetManage Inc. -  All rights reserved
'
' File:     TCPPORTS.FRM
'
' Date:     21 May 1996
'
' Description:
'
'           This file is part of the TCP control sample app-
'   lication of the NetManage Inc. Internet Control Pack.  It
'   contains the TCP Port Search form of the sample application
'   and the form's associated code.
'
'------------------------------------------------------------

Option Explicit

Dim StillSearching As Boolean                           ' Halt search notification variable
Const MAXPORT = 80                                      ' Maximum # of tcp controls instantances that can be created at run time

Private Sub cmdCancel_Click()
    StillSearching = False                              ' Cancel search
End Sub

Private Sub cmdConnect_Click()
    Dim socket As Variant
    Dim port As Long
    On Error Resume Next
    port = Val(txtRemotePort.Text)
    StillSearching = True
    Do While StillSearching
        For Each socket In TCP
          If (socket.State = sckClosed Or socket.State = sckError) Then
                txtRemotePort.Text = CStr(port)
                socket.Close
                DoEvents
                socket.Connect txtName.Text, port
                port = port + 1
                If (port = 19) Then port = port + 1
            End If
            DoEvents
        Next
    Loop
    For Each socket In TCP
        socket.Close
    Next
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To MAXPORT
        Load TCP(i)
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click                                     ' Call cancel to stop searching socket ports...
      Dim socket As Variant
    For Each socket In TCP                              ' Loop through each socket instance...
        socket.Close                                    ' Close all connections and clear any errors...
    Next                                                ' Get next socket instance...
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim socket As Variant
    For Each socket In TCP
        socket.Close
        If (socket.Index <> 0) Then Unload socket
    Next
    End
End Sub

Private Sub mnu2_Click()
    Dim socket As Variant
    Dim port As Long
    On Error Resume Next
    
    port = Val(txtRemotePort.Text)
    StillSearching = True
    
    Do While StillSearching
        For Each socket In TCP
            If (socket.State = sckClosed Or socket.State = sckError) Then
                txtRemotePort.Text = CStr(port)
                socket.Close
                DoEvents
                socket.Connect txtName.Text, port
                port = port + 1
                If (port = 19) Then port = port + 1
            End If
            DoEvents
        Next
    Loop

    For Each socket In TCP
        socket.Close
    Next
End Sub

Private Sub mnu3_Click()
  StillSearching = False
End Sub

Private Sub mnu4_Click()
  Unload Me
End Sub

Private Sub mnu6_Click()
  frmAbout.Show
End Sub

Private Sub mnu7_Click()
  frmUsage.Show
End Sub

Private Sub TCP_Connect(Index As Integer)
    txtOutput.Text = txtOutput.Text & "[Connect]" & _
                    "[" & TCP(Index).RemoteHost & "] " & _
                    "[" & TCP(Index).RemoteHostIP & "] " & _
                    "[" & CStr(TCP(Index).RemotePort) & "]" & vbCrLf
End Sub

Private Sub TCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String
    If (bytesTotal > 0) Then
        TCP(Index).GetData Data
        txtOutput.Text = txtOutput.Text & _
                         "[" & TCP(Index).RemoteHost & "] [" & _
                         TCP(Index).RemotePort & "]" & vbCrLf & _
                         Data & "  " & vbCrLf
    End If
End Sub

Private Sub txtName_Change()

End Sub


Private Sub txtOutput_Change()

End Sub


