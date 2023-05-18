Attribute VB_Name = "NotSoundAnymore"
Public writeError As Boolean
Public DebugData As String
Public v2Last As Boolean
Public PreemptFolder As String

Public Sub AddDebugLine(Str As String)
    DebugData = DebugData & Str & vbCrLf
End Sub

Public Sub Err_Show(modalFrm As Form)
    If DebugData = "" Then Exit Sub
    errWin.Visible = False
    errWin.tErr.Text = DebugData
    errWin.Show vbModal, modalFrm
    Err_Clear
    Unload errWin
End Sub

Public Sub Err_Clear()
    DebugData = ""
End Sub


Public Sub ERR_Add(Str As String)
    DebugData = DebugData & Str & vbCrLf
End Sub

Public Sub ClearDebug()
    DebugData = ""
End Sub
