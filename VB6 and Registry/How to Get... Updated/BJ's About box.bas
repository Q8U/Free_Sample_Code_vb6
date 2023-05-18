Attribute VB_Name = "modAboutBox"
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long

Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
    ShellAbout hwnd, "BJ's How to Get... " & App.EXEName, _
               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", 0
'-----------------------------------------------------------------
End Sub


