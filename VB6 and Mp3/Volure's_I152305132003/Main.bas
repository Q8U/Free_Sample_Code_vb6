Attribute VB_Name = "xMain"
Option Explicit
Public ErrString As String
Global MP3Tag As cIDV3

Public Function ValidateDir(DirPath As String) As Boolean
    On Error Resume Next
    If Dir(DirPath, vbDirectory) = "" Then
        Err.Clear
        MkDir DirPath
    End If
    ValidateDir = (Err = 0)
End Function

Public Function FileExists(ByVal sSpec As String) As Boolean
On Error Resume Next
  Err.Clear
  Call FileLen(sSpec)
  FileExists = (Err = 0)

End Function

Public Function Bin2Dec(Binary As String) As Long
Dim I As Integer
  Bin2Dec = 0
  For I = 1 To Len(Binary)
      Bin2Dec = Bin2Dec + Mid(Binary, I, 1) * 2 ^ ((Len(Binary) - I))
  Next I
End Function

Public Function SumChr(ChrStr As String) As Long
    Dim I As Integer
    For I = 1 To Len(ChrStr)
        SumChr = SumChr + Val(Asc(Mid(ChrStr, I, 1)) * (256 ^ (Len(ChrStr) - I)))
    Next I
End Function
