VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright (C)1997-98 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  Warning: This computer program is protected by copyright law and
'  international treaties. Unauthorized reproduction or distribution
'  of this program, or any portion of it, may result in severe civil
'  and criminal penalties, and will be prosecuted to the maximum
'  extent possible under the law.
' *********************************************************************
Option Explicit
'
' Win32 APIs to determine OS information.
'
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'
' Member variables.
'
Private m_os As OSVERSIONINFO
Private m_NT As Boolean
Private m_95 As Boolean
Private m_98 As Boolean

' ================================
'  Initialize
' ================================
Private Sub Class_Initialize()
   Dim nNull As Long
   '
   ' Retrieve version data for OS.
   '
   m_os.dwOSVersionInfoSize = Len(m_os)
   Call GetVersionEx(m_os)
   '
   ' Trim CSDVersion string at first null
   '
   nNull = InStr(m_os.szCSDVersion, vbNullChar)
   If nNull > 1 Then
      m_os.szCSDVersion = Left(m_os.szCSDVersion, _
                               nNull - 1)
   ElseIf nNull = 1 Then
      m_os.szCSDVersion = ""
   End If
   '
   ' Determine and store values likely to be
   ' referenced often.  VB4/32 will not run
   ' in Win32s, so needn't test for that.
   '
   Select Case m_os.dwPlatformId
      Case VER_PLATFORM_WIN32_WINDOWS
         If m_os.dwMinorVersion >= 10 Then
            m_95 = False
            m_98 = True
         Else
            m_95 = True
            m_98 = False
         End If
         m_NT = False
      Case VER_PLATFORM_WIN32_NT
         m_95 = False
         m_98 = False
         m_NT = True
   End Select
End Sub

' ================================
'  Public Properties
' ================================
Public Property Get MajorVersion() As Long
   MajorVersion = m_os.dwMajorVersion
End Property

Public Property Get MinorVersion() As Long
   MinorVersion = m_os.dwMinorVersion
End Property

Public Property Get BuildNumber() As Long
   BuildNumber = WordLo(m_os.dwBuildNumber)
End Property

Public Property Get PlatformID() As Long
   PlatformID = m_os.dwPlatformId
End Property

Public Property Get IsWinNT() As Boolean
   IsWinNT = m_NT
End Property

Public Property Get IsWin95() As Boolean
   IsWin95 = m_95
End Property

Public Property Get IsWin98() As Boolean
   IsWin98 = m_98
End Property

Public Property Get Platform() As String
   If m_95 Then
      Platform = "Windows 95"
   ElseIf m_98 Then
      Platform = "Windows 98"
   Else 'm_NT
      Platform = "Windows NT"
   End If
End Property

Public Property Get Version() As String
   '
   ' Build and return version info string.
   '
   Version = Platform & _
             " v" & MajorVersion & _
             "." & MinorVersion & _
             ", Build " & BuildNumber
End Property

Public Property Get CSDVersion() As String
   CSDVersion = Trim(m_os.szCSDVersion)
End Property

' ================================
'  Private Methods
' ================================
Private Function WordLo(LongIn As Long) As Integer
   '
   ' Low word retrieved by masking off high word.
   ' If low word is too large, twiddle sign bit.
   '
   If (LongIn And &HFFFF&) > &H7FFF Then
      WordLo = (LongIn And &HFFFF&) - &H10000
   Else
      WordLo = LongIn And &HFFFF&
   End If
End Function

