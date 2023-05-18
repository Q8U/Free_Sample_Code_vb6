Attribute VB_Name = "modMakeINI"

' Two Windows API calls used to read and write private .INI files.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long





'Info to Write to .ini file                                 Info in .ini file

Global Const SECTION = "BJ"                                 '[BJ]
Global Const ENTRY = "Entry"                                'Entry=
Global Const ENTRY1 = "Entry1"                              'Entry1=
Global Const ENTRY2 = "Entry2"                              'Entry2=
Global Const SECTION1 = "Section1"                          '[Section1]
Global Const ENTRY3 = "Entry3"                              'Entry3=
Global Const SECTION2 = "Section2"                          '[Section2]
Global Const ENTRY4 = "Entry4"                              'Entry4=
Global Const SECTION3 = "Section3"                          '[Section3]
Global Const ENTRY5 = "Entry5"                              'Entry5=
Global Const ENTRY6 = "Entry6"                              'Entry6=
Global Const SECTION4 = "Program Information"               '[Program Information]
Global Const ENTRY7 = "Application Path"                    'Application Path=
Global Const ENTRY8 = "Application EXE Name"                'Application EXE Name=
Global Const ENTRY9 = "Application Version"                 'Application Version=
Global Const INI_FILE = "BJ's How to Get... ini file.ini"   '.ini file itself found
                                                            'in your Windows Directory

