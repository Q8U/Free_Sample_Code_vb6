VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Windows Info..."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Get Windows Product Key"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Windows Product ID"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_READ = &H20019

Private Sub Command1_Click()
    Dim RetVal As Long, BufferLen As Long, KeyHandle As Long, DataType As Long
    Dim SubKey As String, Buffer As String
    
    SubKey = "Software\Microsoft\Windows\CurrentVersion" 'For Win95/98
    RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_READ, KeyHandle) 'Try to open Win95/98 key
    If RetVal <> 0 Then 'if Win95/98 key open failed
        SubKey = "Software\Microsoft\Windows NT\CurrentVersion" 'For WinNT
        RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_READ, KeyHandle) 'Try to open WinNT key
        If RetVal <> 0 Then 'if WinNT key open failed
            MsgBox "Could not open registry key!"
            Exit Sub
        End If
    End If
    Buffer = Space(255)
    BufferLen = Len(Buffer)
    RetVal = RegQueryValueEx(KeyHandle, "ProductID", 0, DataType, ByVal Buffer, BufferLen) 'Get Windows Product ID value
    Buffer = Left(Buffer, BufferLen) 'Remove empty space from the end of the 'Buffer' string
    RetVal = RegCloseKey(KeyHandle) 'Close the open key
    MsgBox "Windows Product ID: " & Buffer 'Prompt user with Windows Product ID
End Sub

Private Sub Command2_Click()
    Dim RetVal As Long, BufferLen As Long, KeyHandle As Long, DataType As Long
    Dim SubKey As String, Buffer As String
    
    SubKey = "Software\Microsoft\Windows\CurrentVersion" 'For Win95/98
    RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_READ, KeyHandle) 'Try to open Win95/98 key
    If RetVal <> 0 Then 'if Win95/98 key open failed
        SubKey = "Software\Microsoft\Windows NT\CurrentVersion" 'For WinNT
        RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_READ, KeyHandle) 'Try to open WinNT key
        If RetVal <> 0 Then 'if WinNT key open failed
            MsgBox "Could not open registry key!"
            Exit Sub
        End If
    End If
    Buffer = Space(255)
    BufferLen = Len(Buffer)
    RetVal = RegQueryValueEx(KeyHandle, "ProductKey", 0, DataType, ByVal Buffer, BufferLen) 'Get Windows Product Key value
    Buffer = Left(Buffer, BufferLen) 'Remove empty space from the end of the 'Buffer' string
    RetVal = RegCloseKey(KeyHandle) 'Close the open key
    MsgBox "Windows Product Key: " & Buffer 'Prompt user with Windows Product Key
End Sub
