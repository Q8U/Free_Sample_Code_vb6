VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Command1_Click()
 Dim ret As Integer
 
 ' The following will open the sequencer with the C:\WIN31\CANYON.MID
 ' file. Canyon is the device_id.
 ret = mciSendString("open c:\windows\CANYON.MID type sequencer alias canyon", 0&, 0, 0)

 ' The wait tells the MCI command to complete before returning control
 ' to the application.
 ret = mciSendString("play canyon wait", 0&, 0, 0)
 
 ' Close CANYON.MID file and sequencer device
 ret = mciSendString("close canyon", 0&, 0, 0)
End Sub

