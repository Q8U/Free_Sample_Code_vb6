Attribute VB_Name = "moddragdrop"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Easy CD Cover v1.0 Source Release
'
' by: G@Te^k3eP3R (johnge2@yahoo.com)
'
' Drag 'n' Drop function by  MaD^DoGg

Option Explicit

Type POINTAPI
  x As Long
  y As Long
End Type
Type MSG
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Const PM_NOREMOVE = &H0
Public Const PM_NOYIELD = &H2
Public Const PM_REMOVE = &H1
Public Const WM_DROPFILES = &H233
