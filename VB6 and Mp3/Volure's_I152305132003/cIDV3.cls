VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIDV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "MP3 Tag Information/Editor class."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'----------------------------------------------------
' Filename: cIDV3.cls
' Module:   cIDV3
' By:       Shannon Harmon
' Date:     11/22/1999
' Purpose:  MP3 Tag/Info - Viewer/Editor
'
' If you update this please email me a copy via
' sharmon@vpcusa.com
'
' I realize there are lot's of code snippets to
' do something similar to this, but I would like
' to get a fully functional version, hopefully
' some of you will improve on this, or point me
' to one that is already better than this.  Most
' everything I have seen has been limited to just
' the tag info and not the file info.
'
' Feel free to use this code in whatever, I take
' no repsonibility for any damages caused by using
' this freeware code.  Use at your own risk!
'
' NOTICE: Made with VB6, not tested with anything else.
'
'----------------------------------------------------


Option Explicit
'----------------------------------------------------
Public Enum eGenre
  [] = -1
  [Blues] = 0
  [Classic Rock] = 1
  [Country] = 2
  [Dance] = 3
  [Disco] = 4
  [Funk] = 5
  [Grunge] = 6
  [Hip-Hop] = 7
  [Jazz] = 8
  [Metal] = 9
  [New Age] = 10
  [Oldies] = 11
  [Other] = 12
  [Pop] = 13
  [R&B] = 14
  [Rap] = 15
  [Reggae] = 16
  [Rock] = 17
  [Techno] = 18
  [Industrial] = 19
  [Alternative] = 20
  [Ska] = 21
  [Death Metal] = 22
  [Pranks] = 23
  [Soundtrack] = 24
  [Euro-Techno] = 25
  [Ambient] = 26
  [Trip-Hop] = 27
  [Vocal] = 28
  [Jazz+Funk] = 29
  [Fusion] = 30
  [Trance] = 31
  [Classical] = 32
  [Instrumental] = 33
  [Acid] = 34
  [House] = 35
  [Game] = 36
  [Sound Clip] = 37
  [Gospel] = 38
  [Noise] = 39
  [Alt. Rock] = 40
  [Bass] = 41
  [Soul] = 42
  [Punk] = 43
  [Space] = 44
  [Meditative] = 45
  [Instrumental Pop] = 46
  [Instrumental Rock] = 47
  [Ethnic] = 48
  [Gothic] = 49
  [Darkwave] = 50
  [Techno-Industrial] = 51
  [Electronic] = 52
  [Pop-Folk] = 53
  [Eurodance] = 54
  [Dream] = 55
  [Southern Rock] = 56
  [Comedy] = 57
  [Cult] = 58
  [Gangsta Rap] = 59
  [Top 40] = 60
  [Christian Rap] = 61
  [Pop/Funk] = 62
  [Jungle] = 63
  [Native American] = 64
  [Cabaret] = 65
  [New Wave] = 66
  [Phychedelic] = 67
  [Rave] = 68
  [Showtunes] = 69
  [Trailer] = 70
  [Lo-Fi] = 71
  [Tribal] = 72
  [Acid Punk] = 73
  [Acid Jazz] = 74
  [Polka] = 75
  [Retro] = 76
  [Musical] = 77
  [Rock & Roll] = 78
  [Hard Rock] = 79
  [Folk] = 80
  [Folk/Rock] = 81
  [National Folk] = 82
  [Swing] = 83
  [Fast-Fusion] = 84
  [Bebob] = 85
  [Latin] = 86
  [Revival] = 87
  [Celtic] = 88
  [Bluegrass] = 89
  [Avantegarde] = 90
  [Gothic Rock] = 91
  [Progressive Rock] = 92
  [Psychedelic Rock] = 93
  [Symphonic Rock] = 94
  [Slow Rock] = 95
  [Big Band] = 96
  [Chorus] = 97
  [Easy Listening] = 98
  [Acoustic] = 99
  [Humour] = 100
  [Speech] = 101
  [Chanson] = 102
  [Opera] = 103
  [Chamber Music] = 104
  [Sonata] = 105
  [Symphony] = 106
  [Booty Bass] = 107
  [Primus] = 108
  [Porn Groove] = 109
  [Satire] = 110
  [Slow Jam] = 111
  [Club] = 112
  [Tango] = 113
  [Samba] = 114
  [Folklore] = 115
  [Ballad] = 116
  [Power Ballad] = 117
  [Rhythmic Soul] = 118
  [Freestyle] = 119
  [Duet] = 120
  [Punk Rock] = 121
  [Drum Solo] = 122
  [A Capella] = 123
  [Euro-House] = 124
  [Dance Hall] = 125
  [Goa] = 126
  [Drum & Bass] = 127
  [Club-House] = 128
  [Hardcore] = 129
  [Terror] = 130
  [Indie] = 131
  [BritPop] = 132
  [Negerpunk] = 133
  [Polsk Punk] = 134
  [Beat] = 135
  [Christian Gangsta Rap] = 136
  [Heavy Metal] = 137
  [Black Metal] = 138
  [Crossover] = 139
  [Contemporary Christian] = 140
  [Christian Rock] = 141
  [Merengue] = 142
  [Salsa] = 143
  [Trash Metal] = 144
  [Anime] = 145
  [JPop] = 146
  [Synthpop] = 147
End Enum
'----------------------------------------------------
Dim Tags() As Tagframe

Private Type Tagframe
    Name As String
    Flags As String * 2
    Data As String
End Type
'----------------------------------------------------
Private Type tId3Info
  strTitle As String * 30
  strArtist As String * 30
  strAlbum As String * 30
  strYear  As String * 4
  strComments As String * 30
  byteGenre As Byte
End Type

Private Type tId3v2Info
  strTitle As String
  strArtist As String
  strAlbum As String
  strYear  As String
  strComments As String
  Genre As String
End Type

'----------------------------------------------------

'----------------------------------------------------
Private Type tMPEGInfo
  lFileBytes As Long
  iSeconds As Integer
  iMpegVersion As Integer
  iLayerVersion As Integer
  iKbits As Integer
  lFrames As Long
  lHz As Long
  strMode As String
  strPrivate As String      'Yes/No
  strCRCs As String         'Yes/No
  strCopyrighted As String  'Yes/No
  strOriginal As String     'Yes/No
  strEmphasis As String     'Yes/No
  iChannels As Integer
  strInfoString As String
End Type
'----------------------------------------------------
  
'----------------------------------------------------
'Local variable(s) to hold property value(s)
Private mvarFilename As String
Private id3Info As tId3Info
Private ID3v2Tag As String
Private id3v2Info As tId3v2Info
Private MPEGInfo As tMPEGInfo
Public ID3v2 As Boolean
'----------------------------------------------------
'
'

'----------------------------------------------------
Private Sub Class_Initialize()
  ClearTagVariables
  ReDim Tags(0 To 0)
End Sub
'----------------------------------------------------

Public Property Get InfoString() As String
  InfoString = MPEGInfo.strInfoString
End Property

'----------------------------------------------------
Public Property Get FileBytes() As Long
  FileBytes = MPEGInfo.lFileBytes
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Seconds() As Integer
  Seconds = MPEGInfo.iSeconds
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get MpegVersion() As Integer
  MpegVersion = MPEGInfo.iMpegVersion
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get LayerVersion() As Integer
  LayerVersion = MPEGInfo.iLayerVersion
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Kbits() As Integer
  Kbits = MPEGInfo.iKbits
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Frames() As Long
  Frames = MPEGInfo.lFrames
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Hz() As Long
  Hz = MPEGInfo.lHz
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Mode() As String
  Mode = MPEGInfo.strMode
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get PrivateBit() As String
  PrivateBit = MPEGInfo.strPrivate
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get CRCs() As String
  CRCs = MPEGInfo.strCRCs
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Copyrighted() As String
  Copyrighted = MPEGInfo.strCopyrighted
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Original() As String
  Original = MPEGInfo.strOriginal
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Emphasis() As String
  Emphasis = MPEGInfo.strEmphasis
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Get Channels() As Integer
  Channels = MPEGInfo.iChannels
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Filename(ByVal vData As String)
  ClearTagVariables
  If ExistFile(vData) Then
    mvarFilename = vData
  Else
    mvarFilename = ""
  End If
  
End Property

Public Property Get Filename() As String
    Filename = mvarFilename
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Comments(ByVal vData As String)
    If ID3v2 Then
        id3v2Info.strComments = vData
    Else
        id3Info.strComments = vData
    End If
End Property

Public Property Get Comments() As String
    If ID3v2 Then
        Comments = Replace(id3v2Info.strComments, Chr(0), "")
        If Left(Comments, 3) = "eng" Then Comments = Mid(Comments, 4, Len(Comments))
    Else
        Comments = RTrim(id3Info.strComments)
    End If
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Year(ByVal vData As String)
    If ID3v2 Then
        id3v2Info.strYear = vData
    Else
        id3Info.strYear = vData
    End If
End Property

Public Property Get Year() As String
    If ID3v2 Then
        Year = Replace(id3v2Info.strYear, Chr(0), "")
    Else
        Year = RTrim(id3Info.strYear)
    End If
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Album(ByVal vData As String)
    If ID3v2 Then
        id3v2Info.strAlbum = vData
    Else
        id3Info.strAlbum = vData
    End If
End Property

Public Property Get Album() As String
    If ID3v2 Then
        Album = Replace(id3v2Info.strAlbum, Chr(0), "")
    Else
        Album = RTrim(id3Info.strAlbum)
    End If
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Artist(ByVal vData As String)
    If ID3v2 Then
        id3v2Info.strArtist = vData
    Else
        id3Info.strArtist = vData
    End If
End Property

Public Property Get Artist() As String
    If ID3v2 Then
        Artist = Replace(id3v2Info.strArtist, Chr(0), "")
    Else
        Artist = RTrim(id3Info.strArtist)
    End If
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let Title(ByVal vData As String)
    If ID3v2 Then
        id3v2Info.strTitle = vData
    Else
        id3Info.strTitle = vData
    End If
End Property

Public Property Get Title() As String
    If ID3v2 Then
        Title = Replace(id3v2Info.strTitle, Chr(0), "")
    Else
        Title = RTrim(id3Info.strTitle)
    End If
End Property
'----------------------------------------------------

'----------------------------------------------------
Public Property Let GenreV1(ByVal vData As eGenre)
    If vData = -1 Then
        vData = 0
    End If
    id3Info.byteGenre = CByte(vData)
End Property

Public Property Get GenreV1() As eGenre
Attribute GenreV1.VB_Description = "Returns/Sets current genre."
    GenreV1 = CInt(id3Info.byteGenre)
    If GenreV1 > 148 Then GenreV1 = 0
End Property

Public Property Let GenreV2(ByVal vData As String)
    id3v2Info.Genre = vData
End Property

Public Property Get GenreV2() As String
    GenreV2 = Replace(id3v2Info.Genre, Chr(0), "")
End Property

'----------------------------------------------------

'----------------------------------------------------
Public Function ClearTag() As Boolean
Dim Blank As String * 127
Dim FF As Integer
FF = FreeFile
On Error GoTo clearTagError
    If ID3v2 Then
        If ExistFile(mvarFilename) Then
            Dim Header As String
            Dim Tag As String
            Dim FileDat As String
            Open mvarFilename For Binary As #FF
                
                Header = String(10, Chr(0))
                Get #FF, , Header
                If Left(Header, 3) = "ID3" Then
                    Tag = String(v2TagLen(Mid(Header, 7, 4)), Chr(0))
                    Get #FF, , Tag
                    'Exists
                    FileDat = String(LOF(FF) - Len(Header & Tag), Chr(0))
                    Get #FF, , FileDat
                Else
                    
                    FileDat = String(LOF(FF), Chr(0))
                    Get #FF, 1, FileDat
                End If
            Close #FF
            Open mvarFilename For Output As #FF
            Close #FF
            Open mvarFilename For Binary As #FF
                Put #FF, , FileDat
            Close #FF
        Else
            ClearTag = False
        End If
    Else
        If ExistFile(mvarFilename) Then
            Open mvarFilename For Binary Access Write As #FF
            Seek #FF, LOF(1) - 127
            Put #FF, , Blank
            Close #FF
            ClearTag = True
        Else
            ClearTag = False
        End If
    End If
Exit Function
  
clearTagError:
  ClearTag = False

End Function
'----------------------------------------------------

'----------------------------------------------------
Public Sub ReadTag()
    On Error Resume Next
    Dim I As Integer
    ReDim Tags(0 To 0)
    If Not ExistFile(mvarFilename) Then Exit Sub
        If HasTag Then
        Dim FF As Long
        FF = FreeFile
        If ID3v2 Then
            Dim Header As String
            Dim Tag As String
            Dim FileDat As String
            Open mvarFilename For Binary As #FF
                
                Header = String(10, Chr(0))
                Get #FF, , Header
                If Left(Header, 3) = "ID3" Then
                    Tag = String(v2TagLen(Mid(Header, 7, 4)), Chr(0))
                    Get #FF, , Tag
                    'Exists
                    Do Until Tag = String(Len(Tag), Chr(0))
                        Header = Mid(Tag, 1, 10)
                        Tag = Mid(Tag, 11, Len(Tag))
                        ReDim Preserve Tags(0 To UBound(Tags) + 1)
                        Tags(UBound(Tags)).Name = Mid(Header, 1, 4)
                        Tags(UBound(Tags)).Flags = Mid(Header, 9, 2)
                        Tags(UBound(Tags)).Data = Mid(Tag, 1, v2TagLen(Mid(Header, 5, 4)))
                        Tag = Mid(Tag, Len(Tags(UBound(Tags)).Data) + 1, Len(Tag))
                    Loop
                    Dim sComment() As String
                    Dim I2 As Integer
                    For I = 1 To UBound(Tags)
                        Select Case Tags(I).Name
                            Case "COMM" 'Comment
                                sComment = Split(Tags(I).Data, Chr(0))
                                For I2 = LBound(sComment) To UBound(sComment)
                                    If sComment(I2) = "eng" Then
                                        id3v2Info.strComments = sComment(I2 + 1)
                                        Exit For
                                    ElseIf I2 = UBound(sComment) Then
                                        id3v2Info.strComments = sComment(I2)
                                    End If
                                Next I2
                            Case "TALB" 'Album
                                id3v2Info.strAlbum = Tags(I).Data
                            Case "TIT2" 'Title
                                id3v2Info.strTitle = Tags(I).Data
                            Case "TPE1" 'Artist
                                id3v2Info.strArtist = Tags(I).Data
                            Case "TRCK" 'Track
                                
                            Case "TYER" 'Year
                                id3v2Info.strYear = Tags(I).Data
                            Case "TOPE"
                                'id3v2Info.strArtist = Tags(I).Data
                            Case "TCON" 'Genre
                                id3v2Info.Genre = Tags(I).Data
                            Case ""
                            Case Else
                        End Select
                    Next I
                    
                End If
            Close #FF
        Else
            Open mvarFilename For Binary As #1
            Get #1, FileLen(mvarFilename) - 124, id3Info
            Close #1
        End If
    End If

  ReadMPEGInfo

End Sub
'----------------------------------------------------

'----------------------------------------------------
' Portions of this routine were found on the www.
' Part that finds frames is not working....
'
Private Sub ReadMPEGInfo()
Dim ByteArray(4) As Byte, X As Byte, VBR As Boolean
Dim XingH As String * 4, LayerVersion As String, bin As String
Dim Headstart As Integer, I As Integer, z As Integer, FIO As Integer
Dim sMode, sVersion, sLayer, sEmph, sFreq, sTemp, bRate, NoYes, YesNo
On Error Resume Next

  '----------------------------------------------------
  FIO = FreeFile
  Open mvarFilename For Binary Access Read As FIO
  MPEGInfo.lFileBytes = FileLen(mvarFilename)
  '----------------------------------------------------
  
  '----------------------------------------------------
  'If start position <>1 then id3v2 tag exists
  For I = 1 To 5000 'Check up to 5000 bytes for the header
    Get FIO, I, X
    If X = 255 Then 'Header always start with 255 followed by 250 or 251
      Get FIO, I + 1, X
      If X > 249 And X < 252 Then
        Headstart = I 'Set header start position
        Exit For
      End If
    End If
  Next I
  'End check start position for header'''''''''''''
  '----------------------------------------------------

  '----------------------------------------------------
  'Start check for XingHeader
  Get FIO, Headstart + 36, XingH
  If XingH = "Xing" Then
    VBR = True
    For z = 1 To 4
      Get #1, Headstart + 43 + z, ByteArray(z)  'Get framelength to array
    Next z
    MPEGInfo.lFrames = BinToDec(ByteToBit(ByteArray))   'Calculate # of frames
  Else
    VBR = False
  End If
   'End check for XingHeader
  '----------------------------------------------------

  '----------------------------------------------------
  'Start extract the first 4 bytes (32 bits) to an array
   For z = 1 To 4 '
     Get #1, Headstart + z - 1, ByteArray(z)
   Next z
  'Stop extract the first 4 bytes (32 bits) to an array
  '----------------------------------------------------

  Close FIO
  bin = ByteToBit(ByteArray)

  '----------------------------------------------------
  sVersion = Array(25, 0, 2, 1) 'Mpegversion table
  MPEGInfo.iMpegVersion = sVersion(BinToDec(Mid(bin, 12, 2))) 'Get mpegversion from table
  '----------------------------------------------------
  
  '----------------------------------------------------
  sLayer = Array(0, 3, 2, 1) 'Layer table
  MPEGInfo.iLayerVersion = sLayer(BinToDec(Mid(bin, 14, 2))) 'Get layer from table
  '----------------------------------------------------
  
  '----------------------------------------------------
  sMode = Array("Stereo", "Joint Stereo", "Dual Channel", "Single Channel") 'Mode table
  MPEGInfo.strMode = sMode(BinToDec(Mid(bin, 25, 2))) 'Get mode from table
  
  If MPEGInfo.strMode = "Single Channel" Then
    MPEGInfo.iChannels = 1
  Else
    MPEGInfo.iChannels = 2
  End If
  '----------------------------------------------------
  
  '----------------------------------------------------
  sEmph = Array("None", "50/15", "Reserved", "CCITT J 17") 'Emphasis table
  MPEGInfo.strEmphasis = sEmph(BinToDec(Mid(bin, 31, 2))) 'Get emphasis from table

  Select Case MPEGInfo.iMpegVersion 'Look for version to create right table
    Case 1 'For version 1
      sFreq = Array(44100, 48000, 32000)
    Case 2 Or 25 'For version 2 or 2.5
      sFreq = Array(22050, 24000, 16000)
    Case Else
      MPEGInfo.lHz = 0
      Exit Sub
  End Select
  '----------------------------------------------------
  
  MPEGInfo.lHz = sFreq(BinToDec(Mid(bin, 21, 2))) 'Look for frequency in table
  
  '----------------------------------------------------
  LayerVersion = MPEGInfo.iMpegVersion & MPEGInfo.iLayerVersion 'Combine version and layer to string
  Select Case Val(LayerVersion) 'Look for the right Bits table
    Case 11 'Version 1, Layer 1
      bRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
    Case 12 'V1 L1
      bRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
    Case 13 'V1 L3
      bRate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
    Case 21 Or 251 'V2 L1 and 'V2.5 L1
      bRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
    Case 22 Or 252 Or 23 Or 253 'V2 L2 and 'V2.5 L2 etc...
      bRate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
    Case Else 'If variable Bits
      MPEGInfo.iKbits = 1 'E.G. for variable Bits
      Exit Sub
  End Select
  '----------------------------------------------------
    
  MPEGInfo.iKbits = bRate(BinToDec(Mid(bin, 17, 4)))
  
  '----------------------------------------------------
  NoYes = Array("No", "Yes")
  MPEGInfo.strOriginal = NoYes(Mid(bin, 30, 1))
  MPEGInfo.strCopyrighted = NoYes(Mid(bin, 29, 1))
  MPEGInfo.strPrivate = NoYes(Mid(bin, 24, 1))
  '----------------------------------------------------
  
  '----------------------------------------------------
  YesNo = Array("Yes", "No")
  MPEGInfo.strCRCs = YesNo(Mid(bin, 16, 1))
  '----------------------------------------------------
  
  MPEGInfo.iSeconds = Int(((MPEGInfo.lFileBytes * 8) / MPEGInfo.iKbits) / 1000)
  
  With MPEGInfo 'Formatted info string like Winamp's(tm) MPEG Info box
    .strInfoString = "Size: " & .lFileBytes & " bytes" & vbCrLf _
                    & "Length: " & .iSeconds & " seconds" & vbCrLf _
                    & "MPEG " & Format(.iMpegVersion, "#.0") & " layer " & .iLayerVersion & vbCrLf _
                    & .iKbits & "kbit, " & .lFrames & " frames" & vbCrLf _
                    & .lHz & "hz " & .strMode & vbCrLf _
                    & "Private: " & .strPrivate & vbCrLf _
                    & "CRCs: " & .strCRCs & vbCrLf _
                    & "Copyrighted: " & .strCopyrighted & vbCrLf _
                    & "Original: " & .strOriginal & vbCrLf _
                    & "Emphasis: " & .strEmphasis
  End With

End Sub
'----------------------------------------------------

'----------------------------------------------------
Private Sub ClearTagVariables()
Dim Inbuf As String * 124
  
  With id3Info
    .byteGenre = 12
    .strAlbum = ""
    .strArtist = ""
    .strComments = ""
    .strTitle = ""
    .strYear = ""
  End With

  With id3v2Info
    .Genre = ""
    .strAlbum = ""
    .strArtist = ""
    .strComments = ""
    .strTitle = ""
    .strYear = ""
  End With

  With MPEGInfo
    .lFileBytes = 0
    .iSeconds = 0
    .iMpegVersion = 0
    .iLayerVersion = 0
    .iKbits = 0
    .lFrames = 0
    .lHz = 0
    .strMode = ""
    .strPrivate = ""
    .strCRCs = ""
    .strCopyrighted = ""
    .strOriginal = ""
    .strEmphasis = ""
    .iChannels = 0
  End With

End Sub
'----------------------------------------------------

'----------------------------------------------------
Public Function HasTag() As Boolean
    On Error GoTo hasTagError
    Dim Tag As String * 3
    Dim FF As Integer
    FF = FreeFile
    If Not ExistFile(mvarFilename) Then HasTag = False: Exit Function
    If ID3v2 Then
        Open mvarFilename For Binary As #FF
            Get #1, , Tag
        Close #FF
        If Tag = "ID3" Then
            HasTag = True
        Else
            HasTag = False
        End If
        Exit Function
    Else
        Open mvarFilename For Binary As #FF
        Get #FF, FileLen(mvarFilename) - 127, Tag
        Close #FF
        
        If UCase(Tag) = "TAG" Then
            HasTag = True
        Else
            HasTag = False
        End If
        
        Exit Function
    End If
hasTagError:
    HasTag = False

End Function
'----------------------------------------------------

'----------------------------------------------------
Public Function WriteTag() As Boolean
    Close
    Close
    On Error GoTo writeError
    Dim Tag As String * 3
    Dim I As Integer
    Dim FF As Integer
    FF = FreeFile
    If ID3v2 Then
        Dim v2Tag As String
        Dim FileDat As String
        CheckTag
        For I = 1 To UBound(Tags)
            'If Tags(I).Name
            Select Case Tags(I).Name
                Case "COMM" 'Comment
                    Tags(I).Data = Chr(0) & "eng" & Chr(0) & id3v2Info.strComments
                    
                Case "TALB" 'Album
                    Tags(I).Data = Chr(0) & id3v2Info.strAlbum
                        
                Case "TIT2" 'Title
                    Tags(I).Data = Chr(0) & id3v2Info.strTitle
                    
                Case "TPE1" 'Artist
                    Tags(I).Data = Chr(0) & id3v2Info.strArtist
                                     
                Case "TYER" 'Year
                    Tags(I).Data = Chr(0) & id3v2Info.strYear
                    
                Case "TCON" 'Genre
                    Tags(I).Data = Chr(0) & id3v2Info.Genre
                    
                Case Else
                    'Tags(I).Data = Chr(0)
            End Select
            v2Tag = v2Tag & Tags(I).Name & String(4 - Len(v2TaglenStr(Len(Tags(I).Data))), Chr(0)) & v2TaglenStr(Len(Tags(I).Data)) & Tags(I).Flags & Tags(I).Data
        Next I
        v2Tag = "ID3" & Chr(3) & Chr(0) & Chr(0) & String(4 - Len(v2TaglenStr(Len(v2Tag))), Chr(0)) & v2TaglenStr(Len(v2Tag)) & v2Tag
        ClearTag
        Open mvarFilename For Binary As #FF
            FileDat = String(LOF(FF), Chr(0))
            Get #FF, , FileDat
        Close #FF
        'FF = FreeFile
        Open mvarFilename For Output As #FF
        Close #FF
        'FF = FreeFile
        Open mvarFilename For Binary Access Write As #FF
            Put #FF, , v2Tag
            Put #FF, , FileDat
        Close #FF
    Else
        If Not ExistFile(mvarFilename) Then WriteTag = False: Exit Function
        
        Tag = "TAG"
        FF = FreeFile
        
        Open mvarFilename For Binary Access Write As #FF
            If HasTag Then
                Seek #FF, FileLen(mvarFilename) - 127
            Else
                Seek #FF, FileLen(mvarFilename)
            End If
            Put #FF, , Tag
            Put #FF, , id3Info
            Close #FF
        
            WriteTag = True
        
    End If
    Close
    Exit Function
writeError:
    Debug.Print Err.Number & " " & Err.Description & " *ID3v" & IIf(Me.ID3v2, 2, 1) & "* File:" & Me.Filename
    AddDebugLine ErrString & Err.Number & " " & Err.Description & " *ID3v" & IIf(Me.ID3v2, 2, 1) & "* File:" & Me.Filename & vbCrLf
    writeError = True
    WriteTag = False
End Function
'----------------------------------------------------

'----------------------------------------------------
Public Function GenreText(iGenreID As Integer) As String
Attribute GenreText.VB_Description = "Returns genre item string based on input number from 0 to 147."
Dim aGenre() As Variant

  'If iGenreID < 0 Or iGenreID > 147 Then
    'GenreText = ""

  'Else
  
  aGenre = Array("", "Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
    "Hip-Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", _
    "Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
    "Soundtrack", "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", _
    "Trance", "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", _
    "Noise", "Alt. Rock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
    "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
    "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
    "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", _
    "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", "Acid Punk", "Acid Jazz", _
    "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "National Folk", _
    "Swing", "Fast-Fusion", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
    "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
    "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", _
    "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
    "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
    "Euro-House", "Dance Hall", "Goa", "Drum & Bass", "Club-House", "Hardcore", "Terror", "Indie", "BritPop", _
    "Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
    "Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop")

  GenreText = aGenre(iGenreID)

  'End If
  
End Function
'----------------------------------------------------

'----------------------------------------------------
Public Sub FillComboGenre(inCombo As ComboBox, Optional iWhich As Integer)
On Error Resume Next
Dim I As Integer

  If iWhich < 0 Or iWhich > 147 Then iWhich = -1

  inCombo.Clear
  
  For I = 0 To 148
    inCombo.AddItem GenreText(I)
  Next I
  
  inCombo.ListIndex = iWhich + 1
  Exit Sub
  
End Sub
'----------------------------------------------------

'----------------------------------------------------
'Test file existence with error trapping
Private Function ExistFile(ByVal sSpec As String) As Boolean
On Error Resume Next
  Err.Clear
  Call FileLen(sSpec)
  ExistFile = (Err = 0)

End Function
'----------------------------------------------------

'----------------------------------------------------
'This function converts 4*1 byte array to 4*8 bits
Private Function ByteToBit(ByteArray) As String
Dim z As Integer, I As Integer

  ByteToBit = ""
  For z = 1 To 4
    For I = 7 To 0 Step -1
      If Int(ByteArray(z) / (2 ^ I)) = 1 Then
        ByteToBit = ByteToBit & "1"
        ByteArray(z) = ByteArray(z) - (2 ^ I)
      Else
        If ByteToBit <> "" Then
          ByteToBit = ByteToBit & "0"
        End If
      End If
    Next I
  Next z
  
End Function
'----------------------------------------------------

'----------------------------------------------------
'This function converts Binary string to decimal integer
Private Function BinToDec(BinValue As String) As Long
Dim I As Integer

  BinToDec = 0
  For I = 1 To Len(BinValue)
    If Mid(BinValue, I, 1) = 1 Then
      BinToDec = BinToDec + 2 ^ (Len(BinValue) - I)
    End If
  Next I
  
End Function
'----------------------------------------------------
Private Sub Class_Terminate()

End Sub

Private Function v2TagLen(Length As String) As Double
    Dim I As Integer
    For I = 1 To Len(Length)
        v2TagLen = v2TagLen + Asc(Mid$(Length, I, 1)) * 2 ^ (7 * (Len(Length) - I))
    Next I
End Function

Private Function v2TaglenStr(Length As Long) As String
    Dim I As Integer
    Do Until Length = 0
        v2TaglenStr = Chr(Length Mod 128) & v2TaglenStr
        Length = (Length - (Length Mod 128)) / 128
    Loop
End Function

Private Sub ShowData()
    Dim I As Integer
    For I = 1 To UBound(Tags)
        Debug.Print Tags(I).Name & " " & Tags(I).Data
    Next I
End Sub


Private Sub CheckTag()
    Dim bComment As Boolean
    Dim bAlbum As Boolean
    Dim bTitle As Boolean
    Dim bArtist As Boolean
    Dim bYear As Boolean
    Dim bGenre As Boolean
    Dim I As Integer
    bComment = False
    bAlbum = False
    bTitle = False
    bArtist = False
    bYear = False
    bGenre = False
    For I = LBound(Tags) To UBound(Tags)
        Select Case Tags(I).Name
            Case "COMM" 'Comment
                bComment = True
            Case "TALB" 'Album
                bAlbum = True
            Case "TIT2" 'Title
                bTitle = True
            Case "TPE1" 'Artist
                bArtist = True
            Case "TYER" 'Year
                bYear = True
            Case "TCON" 'Genre
                bGenre = True
            Case Else
                Debug.Print Tags(I).Data
        End Select
    Next I
    If Not bComment Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = Comments
            .Flags = Chr(0) & Chr(0)
            .Name = "COMM"
        End With
    End If
    If Not bAlbum Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = Album
            .Flags = Chr(0) & Chr(0)
            .Name = "TALB"
        End With
    End If
    If Not bTitle Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = Title
            .Flags = Chr(0) & Chr(0)
            .Name = "TIT2"
        End With
    End If
    If Not bArtist Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = Artist
            .Flags = Chr(0) & Chr(0)
            .Name = "TPE1"
        End With
    End If
    If Not bYear Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = Year
            .Flags = Chr(0) & Chr(0)
            .Name = "TYER"
        End With
    End If
    If Not bGenre Then
        ReDim Preserve Tags(0 To UBound(Tags) + 1)
        With Tags(UBound(Tags))
            .Data = GenreV2
            .Flags = Chr(0) & Chr(0)
            .Name = "TCON"
        End With
    End If
End Sub

