VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Package setup"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Out 
      Left            =   1800
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog SelectFiles 
      Left            =   3000
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Package software"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "%extractdir%\setup.exe"
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Enter the path of the output package."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Select the files you would like to be included."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "When the files have been decompressed and extracted, enter what program your would like to be executed."
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SelectFiles.DialogTitle = "Select files to package"
SelectFiles.ShowOpen
If SelectFiles.FileName <> "" Then
    List1.AddItem SelectFiles.FileName
End If
End Sub

Private Sub Command2_Click()
Add App.Path & "\Data\Template.xyz", Text3.Text
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Out.DialogTitle = "Select output package file"
Out.ShowSave
If Out.FileName <> "" Then
    Text3.Text = Out.FileName
End If

End Sub

Private Sub Form_Load()
'Add "C:\myEXE.exe", "C:\Windows\Desktop\Output.exe"
'End
End Sub


Function Add(Template As String, OutputEXE As String)
Dim Temp As String
Dim AVI As String

Open Template For Binary As #1
    Temp = String(LOF(1), "a")
    Get #1, , Temp
Close #1

Open App.Path & "\Data\anim.avi" For Binary As #1
    AVI = String(LOF(1), "a")
    Get #1, , AVI
Close #1
Destroy OutputEXE
Open OutputEXE For Binary As #1
    Put #1, , Temp
    Put #1, , "<start AVI>"
    Put #1, , AVI
    Put #1, , "<end AVI>"
'Close #1

DoCompress OutputEXE

Put #1, , "<start execute command>"
Put #1, , Text2.Text '"%extractdir%\setup.exe"
Put #1, , "<end execute command>"
Close #1
MsgBox "Your package has been created."
End
End Function

Function Destroy(xpath As String)
On Error Resume Next
Kill xpath
End Function

Function DoCompress(OutputEXE As String)
Dim MyZip As New Zip
Dim FileData As String

For i = 0 To List1.ListCount
    MyZip.AddFile List1.List(i)
Next i
Destroy "C:\windows\temp\tmpzip"
MyZip.ZipFileName = "C:\windows\temp\tmpzip"
MyZip.MakeZipFile

Open "C:\windows\temp\tmpzip.zip" For Binary Access Read As #3
    FileData = String(LOF(3), "a")
    Get #3, , FileData
Close #3

'Open OutputEXE For Binary As #4
    Put #1, , "<start compressed file>"
    'MsgBox FileData
    Put #1, , FileData
    Put #1, , "<end compressed file>"

Destroy "C:\windows\temp\tmpzip.zip"
End Function
