Attribute VB_Name = "MemCap"
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ejbantz@usa.net
'* Web: http://www.inlink.com/~ejbantz

'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------

Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const HWND_BOTTOM = 1

'// Memory manipulation
Declare Function lStrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Declare Function lStrCpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub hmemcpy Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    
'// Window manipulation
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean

Function MyFrameCallback(ByVal lwnd As Long, ByVal lpVHdr As Long) As Long

    Debug.Print "FrameCallBack"
    
    Dim VideoHeader As VIDEOHDR
    Dim VideoData() As Byte
    
    '//Fill VideoHeader with data at lpVHdr
    RtlMoveMemory VarPtr(VideoHeader), lpVHdr, Len(VideoHeader)
    
    '// Make room for data
    ReDim VideoData(VideoHeader.dwBytesUsed)
    
    '//Copy data into the array
    RtlMoveMemory VarPtr(VideoData(0)), VideoHeader.lpData, VideoHeader.dwBytesUsed

    Debug.Print VideoHeader.dwBytesUsed
    
    
    Debug.Print VideoData
    
End Function

Function MyYieldCallback(lwnd As Long) As Long

    Debug.Print "Yield"

End Function

Function MyErrorCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long
    
    If iID = 0 Then Exit Function
    
    Dim sStatusText As String
    Dim usStatusText As String
    
    'Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
            
    Debug.Print "Error: ", usStatusText, iID

End Function

Function MyStatusCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long

    If iID = 0 Then Exit Function
   
    Dim sStatusText As String
    Dim usStatusText As String
    
    '// Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
    
    Debug.Print "Status: ", usStatusText, iID

End Function

Function MyVideoStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Beep  '// Replace this with your code!
  
End Function

Function MyWaveStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Debug.Print "WaveStream"

End Function

