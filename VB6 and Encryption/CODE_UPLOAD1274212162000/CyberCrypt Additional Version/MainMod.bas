Attribute VB_Name = "MainMod"
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public ChkWarningMsg As Boolean
Public ExtractPath As String
Public sString As String
Public lLength As Long
Public dFileName As String
Public LastDrive As String
Public ChkIfLoad As Boolean
Public LoadArchive As Boolean
Public ArchiveName As String
Public MoveMe As Boolean
Public ChkFastLoad As Boolean
Public CyTFile As String
Public FileListStart As Long
Public Header As String
Public TmpFile As String
Public FileList As String

'shlwapi.dll is used to get the format of converting bytes
'into KB and MB
Public Function FormatKB(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then FormatKB = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function

'This sub is designed to centre a picture to a fixed set ratio
'of the oraginal size (In other words sets the ratio of the picture
'so it fits perfectly into the picture box returned by the target value
'without off setting the ratio size.
Public Sub CentrePic(Target As PictureBox, Source As StdPicture)
    On Error Resume Next
    Dim PicWidth As Integer
    Dim PicHeight As Integer
    Dim NewWidth As Integer
    Dim NewHeight As Integer
    Dim CenterX As Integer
    Dim CenterY As Integer
    PicWidth = Source.Width / 16.763
    PicHeight = Source.Height / 16.763
    Aspect = PicWidth / PicHeight
    If PicWidth > PicHeight Then
        NewWidth = Target.Width - 240
        NewHeight = Target.Width / Aspect
    Else
        NewWidth = Target.Height * Aspect
        NewHeight = Target.Height - 240
    End If
    CenterX = Target.Width / 2 - NewWidth / 2
    CenterY = Target.Height / 2 - NewHeight / 2
    Target.PaintPicture Source, CenterX, CenterY, NewWidth, NewHeight
End Sub
