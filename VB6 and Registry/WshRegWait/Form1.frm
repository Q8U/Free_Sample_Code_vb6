VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Scripting Host"
   ClientHeight    =   2505
   ClientLeft      =   5310
   ClientTop       =   3165
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdRegistry 
      Caption         =   "Registry Operations"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "System Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Start an Application"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton optWait 
         Caption         =   "Start the application and wait for it to terminate"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton optWait 
         Caption         =   "Start the application and return immediately"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start It"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oShl As New IWshShell_Class
Dim oEnv As New IWshEnvironment_Class
Dim oNet As New IWshNetwork_Class

Private Sub cmdBrowse_Click()
'
' Let user select a file.
'
With CommonDialog1
    .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist
    .Filter = "Applications (*.exe)|*.exe"
    .ShowOpen
    If .FileName <> "" Then
        txtTarget = .FileName
    End If
End With
End Sub
Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdRegistry_Click()
Dim vValue As Variant

On Error Resume Next
'
' RegWrite can write a string (REG_SZ or REG_EXPAND_SZ)
' an integer (REG_DWORD) or binary value which also must
' be an integer(REG_BINARY) to a new or existing registry
' value.  The top level registry key must be HKCU,
' HKLM, HKCR, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE,
' HKEY_CLASSES_ROOT, HKEY_USERS, or HKEY_CURRENT_CONFIG.
'
' These commands create the key HKCU\Test and write
' three values under that key.
'
Call oShl.RegWrite("HKCU\Test\StringValue", "HELLO")
Call oShl.RegWrite("HKCU\Test\Dword", 1, "REG_DWORD")
Call oShl.RegWrite("HKCU\Test\Binary", 1, "REG_BINARY")
'
' RegRead can read a string (REG_SZ or REG_MULTI_SZ,
' REG_EXPAND_SZ) an integer (REG_DWORD) or binary
' The top level registry key must be HKCU, HKLM, HKCR,
' HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_CLASSES_ROOT,
' HKEY_USERS, or HKEY_CURRENT_CONFIG.
'
' These commands read the above values.
'
' If the path specifies a value, that values is
' read. If it ends with a "\" the key's default
' value is read.
'
vValue = oShl.RegRead("HKCU\Test\StringValue")
vValue = oShl.RegRead("HKCU\Test\Dword")
'
' RegDelete will delete a value or a key.  If the
' path ends with "\" the key is deleted.
'
Call oShl.RegDelete("HKCU\Test\StringValue")
Call oShl.RegDelete("HKCU\Test\")
End Sub
Private Sub cmdStart_Click()
Dim l As Long

Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNOACTIVE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9

If Trim$(txtTarget) = "" Then Exit Sub

If optWait(0) Then
    '
    ' The RUN method allows you to start a program
    ' asynchronously or wait until it terminates
    ' depending on the last parameter. If true,
    ' the error code, if any, from the shelled app
    ' is returned.
    '
    Call oShl.Run(txtTarget, SW_SHOWNORMAL, False)
    '
    ' The POPUP method displays a message box that will
    ' disappear after the number of seconds specified by
    ' the second parameter.
    '
    ' The syntax is:
    '    object.popup(Message, [seconds], [title], [buttons])
    '
    ' the buttons are the same values as used for a message box.
    '
    Call oShl.Popup("Application was run asynchronously", 4, "WSH Example", vbInformation)
Else
    l = oShl.Run(txtTarget, SW_SHOWNORMAL, True)
    Call oShl.Popup("Application terminated", 4, "WSH Example", vbInformation)
End If
End Sub
Private Sub cmdSettings_Click()
Dim sValue As String

On Error Resume Next
'
' Environmental values.
'
' Step through this code with the Immediate window open.
'
Set oEnv = oShl.Environment("PROCESS")
sValue = oEnv("WINDIR")
Debug.Print sValue
sValue = oEnv("PATH")
Debug.Print sValue
sValue = oEnv("PROMPT")
Debug.Print sValue
sValue = oEnv("TEMP")
Debug.Print sValue
sValue = oEnv("TMP")
Debug.Print sValue
sValue = oEnv("COMSPEC")
Debug.Print sValue
' NT only.
'sValue = oEnv("NUMBER_OF_PROCESSORS")
'sValue = oEnv("OS")
'sValue = oEnv("HOMEDRIVE")
'sValue = oEnv("HOMEPATH")
'sValue = oEnv("PATHEXT")
'sValue = oEnv("SYSTEMDRIVE")
'sValue = oEnv("SYSTEMROOT")
'
' Network values.
'
sValue = oNet.UserName
Debug.Print sValue
sValue = oNet.ComputerName
Debug.Print sValue
sValue = oNet.UserDomain
Debug.Print sValue
'
' May or may not be supported.
'
sValue = oNet.Organization
Debug.Print sValue
sValue = oNet.Site
Debug.Print sValue
sValue = oNet.UserProfile
Debug.Print sValue

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oShl = Nothing
Set oEnv = Nothing
Set oNet = Nothing
End Sub


