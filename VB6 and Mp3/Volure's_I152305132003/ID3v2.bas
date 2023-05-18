Attribute VB_Name = "ID3v2"
'ID3v2.4.0



Public Sub SetID3Data(Filename As String, Artist As String, Title As String)
    Dim FF As Integer
    Dim FileData As String
    FF = FreeFile
    If Dir(Filename) = "" Then Exit Sub
    If Dir(Filename) = "." Then Exit Sub
    If LCase(Right(Filename, 4)) <> ".mp3" Then Exit Sub
    If LCase(Right(Dir(Filename), 4)) = ".mp3" Then
        Open Filename For Binary As FF
            If LOF(FF) <> 0 Then 'if the file is larger then 0
                FileData = String(LOF(FF), Chr(0))
                
                Input #FF, FileData ' Grab the File
            End If
            If Left(FileData, 5) = "ID3" & Chr(4) & Chr(0) Then
                'Version 4.0
                
            End If
        Close FF
    End If
End Sub
