Attribute VB_Name = "modhttplinks"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Easy CD Cover v1.0 Source Release
'
' by: G@Te^k3eP3R (johnge2@yahoo.com)
'
Option Explicit


Public Const email = "johnge2@yahoo.com"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1


Public Sub sendemail()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub


