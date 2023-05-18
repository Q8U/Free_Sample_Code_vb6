VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hide Drives"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "UnHide"
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton command1 
         Caption         =   "Hide"
         Height          =   255
         Left            =   6720
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check3 
         Caption         =   "C"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check4 
         Caption         =   "D"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check5 
         Caption         =   "E"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check6 
         Caption         =   "F"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check7 
         Caption         =   "G"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check8 
         Caption         =   "H"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check9 
         Caption         =   "I"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check10 
         Caption         =   "J"
         Height          =   255
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check14 
         Caption         =   "N"
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check15 
         Caption         =   "O"
         Height          =   240
         Left            =   720
         TabIndex        =   12
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check16 
         Caption         =   "P"
         Height          =   240
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Q"
         Height          =   240
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check18 
         Caption         =   "R"
         Height          =   240
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check19 
         Caption         =   "S"
         Height          =   240
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check20 
         Caption         =   "T"
         Height          =   240
         Left            =   3120
         TabIndex        =   7
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check21 
         Caption         =   "U"
         Height          =   240
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check22 
         Caption         =   "V"
         Height          =   240
         Left            =   4080
         TabIndex        =   5
         Top             =   600
         Width           =   400
      End
      Begin VB.CheckBox Check23 
         Caption         =   "W"
         Height          =   240
         Left            =   4560
         TabIndex        =   4
         Top             =   600
         Width           =   525
      End
      Begin VB.CheckBox Check11 
         Caption         =   "K"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check12 
         Caption         =   "L"
         Height          =   240
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check13 
         Caption         =   "M"
         Height          =   240
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Width           =   400
      End
      Begin VB.CheckBox Check24 
         Caption         =   "X"
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Y"
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox Check26 
         Caption         =   "Z"
         Height          =   255
         Left            =   6000
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "You Must Restart Your Windows For The Changes To Take Effect"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Width           =   4710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "javaidshpk@yahoo.com"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6240
      TabIndex        =   32
      Top             =   3240
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Email:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5280
      TabIndex        =   31
      Top             =   3240
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Muhammad Javaid"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6240
      TabIndex        =   30
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5280
      TabIndex        =   29
      Top             =   2880
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hide Drives In My Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   28
      Top             =   480
      Width           =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Private Sub Command1_Click()
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
If Check1.Value = 1 Then
    Check1.Tag = 1
Else
    Check1.Tag = 0
End If

If Check2.Value = 1 Then
    Check2.Tag = 2
Else
    Check2.Tag = 0
End If
If Check3.Value = 1 Then
    Check3.Tag = 4
Else
    Check3.Tag = 0
End If

If Check4.Value = 1 Then
    Check4.Tag = 8
Else
    Check4.Tag = 0
End If
If Check5.Value = 1 Then
    Check5.Tag = 16
Else
    Check5.Tag = 0
End If

If Check6.Value = 1 Then
    Check6.Tag = 32
Else
    Check6.Tag = 0
End If
If Check7.Value = 1 Then
    Check7.Tag = 64
Else
    Check7.Tag = 0
End If

If Check8.Value = 1 Then
    Check8.Tag = 128
Else
    Check8.Tag = 0
End If
If Check9.Value = 1 Then
    Check9.Tag = 256
Else
    Check9.Tag = 0
End If

If Check10.Value = 1 Then
    Check10.Tag = 512
Else
    Check10.Tag = 0
End If
If Check11.Value = 1 Then
    Check11.Tag = 1024
Else
    Check11.Tag = 0
End If

If Check12.Value = 1 Then
    Check12.Tag = 2048
Else
    Check12.Tag = 0
End If
If Check13.Value = 1 Then
    Check13.Tag = 4096
Else
    Check13.Tag = 0
End If

If Check14.Value = 1 Then
    Check14.Tag = 8192
Else
    Check14.Tag = 0
End If
If Check15.Value = 1 Then
    Check15.Tag = 16384
Else
    Check15.Tag = 0
End If

If Check16.Value = 1 Then
    Check16.Tag = 32768
Else
    Check16.Tag = 0
End If
If Check17.Value = 1 Then
    Check17.Tag = 65536
Else
    Check17.Tag = 0
End If

If Check18.Value = 1 Then
    Check18.Tag = 131072
Else
    Check18.Tag = 0
End If
If Check19.Value = 1 Then
    Check19.Tag = 262144
Else
    Check19.Tag = 0
End If '

If Check20.Value = 1 Then
    Check20.Tag = 524288
Else
    Check20.Tag = 0
End If
If Check21.Value = 1 Then
    Check21.Tag = 1048576
Else
    Check21.Tag = 0
End If

If Check22.Value = 1 Then
    Check22.Tag = 2097152
Else
    Check22.Tag = 0
End If
If Check23.Value = 1 Then
    Check23.Tag = 4194304
Else
    Check23.Tag = 0
End If

If Check24.Value = 1 Then
    Check24.Tag = 8388608
Else
    Check24.Tag = 0
End If
If Check25.Value = 1 Then
    Check25.Tag = 16777216
Else
    Check25.Tag = 0
End If

If Check26.Value = 1 Then
    Check26.Tag = 33554432
Else
    Check26.Tag = 0
End If

a1 = CLng(Check1.Tag) + CLng(Check2.Tag) + CLng(Check3.Tag) _
+ CLng(Check4.Tag) + CLng(Check5.Tag) + CLng(Check6.Tag) + _
CLng(Check7.Tag) + CLng(Check8.Tag) + CLng(Check9.Tag) + _
CLng(Check10.Tag) + CLng(Check11.Tag) + CLng(Check12.Tag) _
+ CLng(Check13.Tag) + CLng(Check14.Tag) + CLng(Check15.Tag) + _
CLng(Check16.Tag) + CLng(Check17.Tag) + CLng(Check18.Tag) _
+ CLng(Check19.Tag) + CLng(Check20.Tag) + CLng(Check21.Tag) _
+ CLng(Check22.Tag) + CLng(Check23.Tag) + CLng(Check24.Tag) _
+ CLng(Check25.Tag) + CLng(Check26.Tag)
If a1 = 0 Then
    s = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    a = "NoDrives"
    lRegResult = RegOpenKey(HKEY_CURRENT_USER, s, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, a)
    lRegResult = RegCloseKey(hCurKey)
Else
    If a1 <> 0 Then
    regkey = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    retvalue = RegCreateKey(HKEY_CURRENT_USER, regkey, KeyID)
    subKey = "NoDrives"
    keyvalue = a1
    retvalue = RegSetValueEx(KeyID, subKey, 0&, 4, keyvalue, 4)
End If
End If
End Sub




Private Sub Command2_Click()
Dim retvalue As Long, result As Long
Dim KeyID As Long, keyvalue As Long
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
Dim abc As Long
Dim a1 As Long
Dim hCurKey As Long
Dim lRegResult As Long
Dim s As String
Dim a As String
 s = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
 a = "NoDrives"
 lRegResult = RegOpenKey(HKEY_CURRENT_USER, s, hCurKey)
 lRegResult = RegDeleteValue(hCurKey, a)
 lRegResult = RegCloseKey(hCurKey)
End Sub
