Attribute VB_Name = "mdbMsgBox"
Public SystemRootS As String
Global Result As Integer

Public Enum Buttons
    OKOnly = 0
    OKCancel = 1
    YesNo = 2
    YesNoCancel = 3
    RetryCancel = 4
    AbortRetryIgnore = 5
End Enum

Public Enum Icons
    Critical = 0
    Question = 1
    Warning = 2
    Information = 3
    None = 4
End Enum

Public Function MessageBox(Mensagem As String, Buttons As Buttons, Icon As Icons)
    frmMessageBox.lblMessage = Mensagem
    
    'Erase all icons
    For Index = 0 To 3
        frmMessageBox.imgIcon(Index).Visible = False
    Next Index
    
    If Icon <> None Then frmMessageBox.imgIcon(Icon).Visible = True
    
    'Erase all buttons
    For Index = 0 To 2
        frmMessageBox.pctButton(Index).Visible = False
        frmMessageBox.lblCaption(Index).FontBold = False
    Next Index
    
    
    If Buttons = AbortRetryIgnore Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "Retry"
        frmMessageBox.lblCaption(0).FontBold = True
        
        frmMessageBox.pctButton(1).Visible = True
        frmMessageBox.lblCaption(1).Caption = "Ignore"
        
        frmMessageBox.pctButton(2).Visible = True
        frmMessageBox.lblCaption(2).Caption = "Abort"
    
    ElseIf Buttons = OKCancel Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "OK"
        frmMessageBox.lblCaption(0).FontBold = True
        
        frmMessageBox.pctButton(1).Visible = True
        frmMessageBox.lblCaption(1).Caption = "Cancel"
    
    ElseIf Buttons = OKOnly Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "OK"
        frmMessageBox.lblCaption(0).FontBold = True
        
    ElseIf Buttons = RetryCancel Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "Retry"
        frmMessageBox.lblCaption(0).FontBold = True
        
        frmMessageBox.pctButton(1).Visible = True
        frmMessageBox.lblCaption(1).Caption = "Cancel"
    
    ElseIf Buttons = YesNoCancel Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "Yes"
        frmMessageBox.lblCaption(0).FontBold = True
        
        frmMessageBox.pctButton(1).Visible = True
        frmMessageBox.lblCaption(1).Caption = "No"
    
        frmMessageBox.pctButton(2).Visible = True
        frmMessageBox.lblCaption(2).Caption = "Cancel"
    
    ElseIf Buttons = YesNo Then
        frmMessageBox.pctButton(0).Visible = True
        frmMessageBox.lblCaption(0).Caption = "Yes"
        frmMessageBox.lblCaption(0).FontBold = True
        
        frmMessageBox.pctButton(1).Visible = True
        frmMessageBox.lblCaption(1).Caption = "No"
    
    End If
    
    frmMessageBox.Show 1
End Function

Function FileExist(FileName As String) As Boolean
    On Error GoTo Erro
    If FileLen(FileName) <> 0 Then
        FileExist = True
    Else
        FileExist = False
    End If
    Exit Function
Erro:
    If Err = 76 Or Err = 53 Then FileExist = False
End Function
