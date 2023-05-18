VERSION 5.00
Begin VB.UserControl ReReg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ReReg.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ReReg.ctx":030A
   Begin VB.Timer ReRegTimer 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ReReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ############################################################
' #                                                          #
' # Date: 7-6-2000                                           #
' #                                                          #
' # ReReg was developed by:                                  #
' #                                                          #
' # Name:     Bas van de Ree                                 #
' # Company:  N.V. Interpolis                                #
' # Place:    Tilburg, The Netherlands                       #
' # Email:    b.vd.ree@interpolis.nl                         #
' #                                                          #
' # If you have any comments or enhancements, please mail me #
' #                                                          #
' ############################################################

Option Explicit

'Set API constants
Const TH32CS_SNAPPROCESS As Long = 2&
Const MAX_PATH As Integer = 260
Const SMTO_BLOCK = &H1
Const SMTO_ABORTIFHUNG = &H2
Const WM_NULL = &H0
Const WM_CLOSE = &H10
Const PROCESS_ALL_ACCESS = &H1F0FFF

'Define API type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

'Declare API functions
Private Declare Function GetWindowsDirectoryA Lib "Kernel32" (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" _
    Alias "CreateToolhelp32Snapshot" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "Kernel32" _
    Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "Kernel32" _
    Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "Kernel32" _
   (ByVal hPass As Long)
   
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, _
   ByVal uExitCode As Long) As Long
   
Public Function ReloadRegistry()

Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
Dim Naam As String
    
    ReRegTimer.Enabled = True
    
    'Make a snapshot of the current processes
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then
    Exit Function
    End If
    'Walk trough the processes and kill all Explorer processes
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Do While r
        Naam = Trim(uProcess.szExeFile)
        Naam = Left(Naam, InStr(1, Naam, Chr$(0)) - 1)
        If Naam = WinDir & "\EXPLORER.EXE" Then
            TerminatePID (uProcess.th32ProcessID)
        End If
        r = ProcessNext(hSnapShot, uProcess)
        Loop
    Call CloseHandle(hSnapShot)
    
End Function

Private Sub ReRegTimer_Timer() 'Timer Sub to initialize explorer after it has been terminated
    
    Shell (WinDir & "\EXPLORER.EXE")
    ReRegTimer.Enabled = False
        
End Sub
Public Function WinDir() As String 'Function to retrieve the location of the windows folder

    Dim sBuf As String
    Dim cSize As Long
    Dim retval As Long
      
    sBuf = String(255, 0)
    cSize = 255
    
    retval = GetWindowsDirectoryA(sBuf, cSize)
    sBuf = Left(sBuf, retval)
    
    WinDir = sBuf

End Function

Public Function TerminatePID(PID As Long) 'Function to kill a process given it's ProcessId

    Dim lngReturnValue As Long
    Dim lngProcess As Long

    lngProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, PID)
    lngReturnValue = TerminateProcess(lngProcess, 0&)

End Function
