VERSION 5.00
Begin VB.Form VBSYXSEN 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "VB SYSEX Send SYSEX"
   ClientHeight    =   720
   ClientLeft      =   1815
   ClientTop       =   3000
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "vbsyxsen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   720
   ScaleWidth      =   5160
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Test Tone"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Send Sysex"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Output Port"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "VBSYXSEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' 32 bit
'
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function MIDIOutOpen Lib "winmm.dll" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Dim OutDevCount As Integer

Private Sub Command1_Click()
    VBSYXMID.Show 1
End Sub

Private Sub Command2_Click()
    SYXSend
End Sub

Private Sub Command3_Click()
    Dim rc As Integer, tm
    
    If VBSYXMID.Text.Text <> "N" Then
        mDev = (Val(VBSYXMID.Text.Text)) - 1
        rc = MIDIOutOpen(hMidi, mDev, 0&, 0&, 0&)
        If rc <> 0 Then
            Call MidiErr("Open", rc)
            Exit Sub
        End If
        rc = midiOutShortMsg(hMidi, &H7F3C90) ' middle c note on velocity 127
        tm = Timer
        For zz = 1 To 32760
            If tm + 1 < Timer Then Exit For
        Next
        rc = midiOutShortMsg(hMidi, &H7F3C80) ' middle c note off velocity 127
        rc = midiOutClose(hMidi)
        If rc <> 0 Then
            Call MidiErr("Close", rc)
        End If
    End If
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim zz As Integer
    Dim OutCaps As MIDIOUTCAPS
    
    OutDevCount = midiOutGetNumDevs()
    VBSYXMID.List1.List(0) = "Device not enabled"
    For zz = 0 To OutDevCount - 1            ' Midi Mapper = -1
        vntRet = midiOutGetDevCaps(zz, OutCaps, Len(OutCaps))
        If vntRet <> 0 Then
            MsgBox "midiOutGetDevCaps Error: " & vntRet
            Exit For
        End If
        VBSYXMID.List1.List(zz + 1) = OutCaps.szPname
    Next zz
    VBSYXMID.Show 1
    OpenSYX
End Sub

Private Sub LongMidiMessage(InString As String)
    Dim mHdr As MIDIHDR
    Dim rc As Integer
    Dim Length As Integer
    Dim Checks As Integer
    Dim midistring(255) As Byte ' Make sure this is big enough!
    
    ' This is wrong - you cannot use strings in MIDIHDR
    'Length = Len(InString)
    'mHdr.lpData = InString
    '
    ' Let's do this instead
    Length = 255
    For rc = 1 To Len(InString)
        midistring(rc - 1) = Asc(Mid$(InString, rc, 1))
    Next rc
    midistring(rc - 1) = 0 ' end of string - just in case :-)
    mHdr.lpData = VarPtr(midistring(0)) ' Undocumented feature!
    '
    mHdr.dwBufferLength = Length
    mHdr.dwBytesRecorded = 0 ' Was Length - only used for MIDI in
    mHdr.dwUser = 0
    mHdr.dwFlags = 0
    ' this next line has caused an error on one user's machine under VB5 - who knows why?
    rc = midiOutPrepareHeader(hMidi, mHdr, LenB(mHdr))
    If rc <> 0 Then
        MsgBox "Prepare rc = " & rc
        Exit Sub
    End If
    ' send long message
    rc = midiOutLongMsg(hMidi, mHdr, LenB(mHdr))
    If rc <> 0 Then
        MsgBox "Send Long Message rc= " & rc
        Exit Sub
    End If
    ' this next line is only required under VB5 IF
    ' you declare the mHdr.lpData as a string
    ' In this new code, mHdr.lpData is a ptr to byte array
    ' and thus this kludge is not required :)
    ' mHdr.dwFlags = 0 ' this line not required anymore
    ' this next line now works under VB5
    rc = midiOutUnprepareHeader(hMidi, mHdr, LenB(mHdr))
    If rc <> 0 Then
        MsgBox "Unprepare rc= " & rc
        Exit Sub
    End If
End Sub

Private Sub MidiErr(mOpt As String, rc As Integer)
    Dim msgText As String * 132
    
    vntRet = midiOutGetErrorText(rc, msgText, 128)
    MsgBox "Operation: " & mOpt & Chr(13) & Chr(10) & msgText
End Sub

Private Sub OpenSYX()
    Dim syx As String * 210
    
    Fname = "E:\vb\midi\midi1\ok.syx"
    ' Fname = "E:\vb\midi\midi1\gsreset.syx"
    Fnum = FreeFile ' Determine file number.
    Open Fname For Binary Access Read As Fnum   ' Open file.
    syxlen = LOF(Fnum) ' Get number of bytes in file
    If syxlen > 200 Then
        Response = MsgBox("Sorry, I don't handle SYSEX files greater than 200 bytes in length. You wouldn't want to decode them anyway!", 16, "File too large!")
        Exit Sub
    End If
    Get Fnum, 1, syx ' read as many bytes as we can into syx string
    Close   ' Close all Files
    outSYX = Left$(syx, syxlen)
End Sub

Private Sub SYXSend()
    Dim rc As Integer
    
    If VBSYXMID.Text.Text <> "N" Then
        mDev = (Val(VBSYXMID.Text.Text)) - 1
        rc = MIDIOutOpen(hMidi, mDev, 0&, 0&, 0&)
        If rc <> 0 Then
            Call MidiErr("Open", rc)
            Exit Sub
        End If
        LongMidiMessage (outSYX)
        rc = midiOutClose(hMidi)
        If rc <> 0 Then
            Call MidiErr("Close", rc)
        End If
    End If
End Sub

