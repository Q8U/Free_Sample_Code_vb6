VERSION 5.00
Begin VB.UserControl ScrllngFrm 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "ScrllngFrm.ctx":0000
   ScaleHeight     =   1815
   ScaleWidth      =   2655
   ToolboxBitmap   =   "ScrllngFrm.ctx":0050
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   720
      Top             =   720
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1575
      Left            =   2400
      Max             =   115
      SmallChange     =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      Max             =   80
      SmallChange     =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox pCorner 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1560
      Width           =   315
   End
   Begin VB.PictureBox pView 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   2595
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "ScrllngFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Name: ScrllngFrm
'
'Description: This control is very useful for
' those who need more space on their forms.
' Run this example.
'
'How to use:
'
' 1. Insert a ScrllngFrm Control into your Form.
'
' 2. Insert one or more Picture Boxes into the
'    ScrllngFrm Control.
'
' 3. Set the visible property of each Picture Box
'    to False.
'
' 4. Insert other controls (Such us Command Buttons,
'    Text Boxes...) into each Picture Box.
'
'    TIP: Right-click the Picture Boxes and select
'         "Bring To Front" or "Send To Back" so
'         you can edit the controls contained by
'         each PictureBox more comfortably.
'
' 5. If you added Command Buttons to your Picture
'    Boxes (the pages) you should set their Style
'    property to Graphical.
'
' 6. On the Form_Load Event call the AddPage function.
'    Each Picture Box will correspond to a page.
'
'
'Notes:
'   The Control captures the events of the Picture Box,
'   so, if you resize the Picture Box, the control
'   adjust the scrollbars. Also, if you resize the
'   ScrllngFrm Control, it adjust its properties.
'   You can, also, have more than one PictureBox added
'   to the control. This control will manage each
'   PictureBox as if each one was a page.
'
'Acknowledgment:
'
'   Some parts of this ActiveX Control were based
'   on codes submitted by other programmers on
'   Planet-Source-Code.
'   I would like to give many thanks to:
'
'   Fred_Cpp
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=31896&lngWId=1
'
'   TopCoder
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13566&lngWId=1
'
'Author:
'   Elias Barbosa
'   Date: 02/19/2002
'   Updated: 03/18/2002
'   Updated: 03/28/2002
'   e-mail: elias@eb8.com
'   http://www.planet-source-code.com/vb/default.asp?lngCId=32374&lngWId=1

Option Explicit

Private intChanged As Boolean
Private Gpast As Variant
Private Gcurrent As Object
Private lPrevParent As Long
Private WithEvents pChild As PictureBox
Attribute pChild.VB_VarHelpID = -1
Private currPage As Integer
Private FirstControl()
Private intSetFocus As Boolean

'Default Property Values:
Const m_def_MemorizeField = True
Const m_def_MemorizeScroll = True
Const m_def_NextEnabled = False
Const m_def_PreviousEnabled = False
Const m_def_CurrentPage = 0
Const m_def_HowManyPages = 0
Const m_def_SelectText = True
Const m_def_HighPicture = False
Const m_def_HighlightColor = &HFFC0C0
Const m_def_Highlight = True
Const m_def_BackColor = &H8000000C

'Property Variables:
'
'============================
'These properties are related
'to page navigation.
'============================
Dim m_NextEnabled As Boolean
Dim m_PreviousEnabled As Boolean
Dim m_CurrentPage As Integer
Dim m_HowManyPages As Integer

'============================
'These properties are related
'to field selection behavior.
'============================
Dim m_MemorizeField As Boolean
Dim m_MemorizeScroll As Boolean
Dim m_SelectText As Boolean
Dim m_Highlight As Boolean
Dim m_HighPicture As Boolean
Dim m_HighlightColor As OLE_COLOR

'============================
'These properties are related
'to UserControl appearance.
'============================
Dim m_BackPicture As Picture
Dim m_BackColor As OLE_COLOR

'Event Declarations:
Event Resize()
Event Scroll()
Public Event FocusMoved()
Public Event PageChanged()

'API Declarations
Private Declare Function SetParent _
    Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long

Private Sub UserControl_Initialize()
    'This Array will manage each page (PictureBox) added to the control.
    'FirstControl(1, 1) = hWnd of current page (PictureBox attached).
    'FirstControl(2, 1) = Name of the first control on TabStop Index that is contained by current page.
    'FirstControl(3, 1) = Index number of the first control on TabStop Index that is contained by current page.
    'FirstControl(4, 1) = hWnd of the first control on TabStop Index that is contained by current page.
    'FirstControl(5, 1) = Last value of the VScroll Scroll Bar on current page.
    'FirstControl(6, 1) = Last value of the HScroll Scroll Bar on current page.
    'FirstControl(7, 1) = Name of PictureBox that represent current page.
    'FirstControl(8, 1) = Index Number of PictureBox that represent current page.
    ReDim FirstControl(8, 1)
    
End Sub

'=======================================================
'======= Following are some Subs that are called =======
'======= by events associated with controls      =======
'======= contained in this UserControl or with   =======
'======= the UserControl itself.                 =======
'=======================================================

'Some of the most important tasks executed on this
'control are taken care of on the Resize Sub.
'For example:
'    * The scrolling size adjustment.
'    * The necessity or not of having a horizontal
'      or vertical scroll bar visible.
'    * The maximum and minimum values of each Scroll
'      Bar after each resize of the form...
Private Sub UserControl_Resize()
    Dim loff As Integer
    Dim loffV As Integer
    Dim loffH As Integer
    Dim sV As Single
    Dim sH As Single
    On Error Resume Next
    
    'Vertical additional space...
    loffV = 39
    'Horizontal addidional space...
    loffH = 45
    
    'The following subs will be called
    're-dimension the UserControl window
    'according to the new size of the new
    'UserControl size.
    Call VScroll.Move(UserControl.Width - VScroll.Width - loffV, 0, VScroll.Width, UserControl.Height - HScroll.Height - loffH)
    Call HScroll.Move(0, UserControl.Height - HScroll.Height - loffH, UserControl.Width - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(UserControl.Width - VScroll.Width - loffV, UserControl.Height - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(0, 0, Width - VScroll.Width, Height - HScroll.Height)
    
    HScroll.Min = 1
    VScroll.Min = 1
    
    sH = pChild.Width - pView.Width
    sV = pChild.Height - pView.Height
    
    'Modify Vertical ScrollBar.
    If (sV = 0) Then
        VScroll.Max = 1
        VScroll.Width = 0
        VScroll.Left = UserControl.Width
        loffV = 37
    ElseIf (sV < 0) Then
        VScroll.Max = 1 ' -sV
        VScroll.Width = 0
        VScroll.Left = UserControl.Width
        loffV = 37
    Else
        VScroll.Max = sV
        VScroll.Width = 255
    End If
    
    'Modify Horizontal Scrollbar.
    If (sH = 0) Then
        HScroll.Max = 1
        HScroll.Height = 0
        loffH = 25
    ElseIf (sH < 0) Then
        HScroll.Max = 1 '-sH
        HScroll.Visible = False
        HScroll.Height = 0
        loffH = 25
    Else
        HScroll.Max = sH
        HScroll.Visible = True
        HScroll.Height = 255
        
    End If
    
    'The following subs will be called again
    'because, depending on the new size of the
    'UserControl, one of the Scrolling Bars may
    'be hidden. On this event the UserControl
    'window will have to be re-dimensioned to
    'adjust to this new circumstance.
    Call VScroll.Move(UserControl.Width - VScroll.Width - loffV, 0, VScroll.Width, UserControl.Height - HScroll.Height - loffH)
    Call HScroll.Move(0, UserControl.Height - HScroll.Height - loffH, UserControl.Width - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(UserControl.Width - VScroll.Width - loffV, UserControl.Height - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(0, 0, Width - VScroll.Width, Height - HScroll.Height)
    
    VScroll.Max = pChild.Height - pView.Height
    HScroll.Max = pChild.Width - pView.Width
    
    HScroll.LargeChange = UserControl.Width
    VScroll.LargeChange = UserControl.Height
    
    RaiseEvent Resize
    
End Sub

Private Sub pChild_Resize()
    Call UserControl_Resize
    
End Sub

'Both events, the Change event and the Scroll events
'are used for the following reason:
'   Change Event: This event will be called if the
'      Scroll Bar value has changed. This event will
'      be called even if the change was made by code.
'   Scroll Event: This event will be called while the
'      user interacts with the Scroll Bar and not after
'      the user have changed the scroll bar position.
Private Sub VScroll_Change()
    UpdatePos
    FirstControl(5, currPage) = VScroll.Value
    
End Sub

Private Sub VScroll_Scroll()
    UpdatePos
    
End Sub

Private Sub HScroll_Change()
    UpdatePos
    FirstControl(6, currPage) = HScroll.Value
    
End Sub

Private Sub HScroll_Scroll()
    UpdatePos
    
End Sub

Private Sub pChild_GotFocus()
    'The pChild_GotFocus event is raised whenever the
    'user sets focus to the Picture Box that was set
    'as the current page. Because setting focus to the
    'page (Picture Box) was not a desirable behavior,
    'I decided to redirect the focus to the last
    'selected field on the current page.
    Call subSelectFirst
    
End Sub

Private Sub UserControl_EnterFocus()
    'The UserControl_EnterFocus event is raised whenever
    'the user sets focus to the User Control itself.
    'Because setting focus to the User Control was not
    'a desirable behavior, I decided to redirect the
    'focus to the last selected field on the current
    'page. However, I would get an error if I tried to
    'do this while there is no page attached to the
    'user control. That's why I used the variable
    'intSetFocus. When this variable is set to true
    'and there is no page attached, the User Control
    'will not try to set focus to any field.
    intSetFocus = True
    Call subSelectFirst
    
End Sub

'==================================
'======= Following are some  ======
'======= complementary Subs. ======
'==================================

Private Sub UpdatePos()
    'Called when Scrolls have Changed
    On Error Resume Next
    pChild.Move -HScroll.Value, -VScroll.Value
    RaiseEvent Scroll
    
End Sub

Public Function hwnd()
    hwnd = UserControl.hwnd
    
End Function

Public Sub AddPage(NewPictureBox)
    Dim intPage As Integer
    Dim intTempIndex As String
    Dim i As Integer
    
    If (Len(FirstControl(1, 0)) > 0) Then
        
        For i = 0 To m_HowManyPages - 1
            If (Len(FirstControl(1, i)) = 0) _
            And (i = 0) Then
                intPage = 0
                Exit For
                
            ElseIf (FirstControl(1, i) = NewPictureBox.hwnd) Then
                intPage = i
                Exit For
                
            ElseIf (i = m_HowManyPages - 1) Then
                intPage = m_HowManyPages
                FirstControl(1, intPage) = NewPictureBox.hwnd
                
            End If
            
        Next i
        
        'If another page is added, re-dimension
        'the array to hold this extra information.
        If (intPage = m_HowManyPages) Then
            ReDim Preserve FirstControl(8, intPage + 1)
            m_HowManyPages = m_HowManyPages + 1
            
            FirstControl(2, intPage) = "This is a new Page."
            FirstControl(7, intPage) = NewPictureBox.Name
            
            On Error Resume Next
            intTempIndex = CStr(pChild.Index)
            
            FirstControl(8, intPage) = intTempIndex
        End If
        
        Call subUpdateNav
        RaiseEvent PageChanged
        
    Else
        m_HowManyPages = 1
        Call subAttach(NewPictureBox)
    End If
    
End Sub

Public Sub DeletePage(my_PageNumber As Integer)
    Dim intTempArray()
    Dim i As Integer
    Dim j  As Integer
    
    If (my_PageNumber > m_HowManyPages) _
    Or (my_PageNumber < 1) Then
        Exit Sub
        
    End If
    If (m_HowManyPages = 1) Then
        'If there was another PictureBox attached,
        'restore the "parenthood" for this PictureBox
        'before Attaching the new PictureBox.
        If (lPrevParent <> 0) Then
            pChild.Visible = False
            'Restore "parenthood" of the Picture Box! :)
            Call SetParent(pChild.hwnd, lPrevParent)
            
            'Release computer resources...
            Set pChild = Nothing
            lPrevParent = 0
            ReDim FirstControl(8, 1)
            m_HowManyPages = 0
            currPage = 0
            
        End If
        
    ElseIf (m_HowManyPages > 1) Then
        
        m_HowManyPages = m_HowManyPages - 1
        
        ReDim intTempArray(8, m_HowManyPages)
        
        For i = 0 To m_HowManyPages - 1
            If (i = my_PageNumber - 1) Then
                j = j + 1
            End If
            
            intTempArray(1, i) = FirstControl(1, j)
            intTempArray(2, i) = FirstControl(2, j)
            intTempArray(3, i) = FirstControl(3, j)
            intTempArray(4, i) = FirstControl(4, j)
            intTempArray(5, i) = FirstControl(5, j)
            intTempArray(6, i) = FirstControl(6, j)
            intTempArray(7, i) = FirstControl(7, j)
            intTempArray(8, i) = FirstControl(8, j)
            
            j = j + 1
            
        Next i
        
        ReDim FirstControl(8, m_HowManyPages)
        
        For i = 0 To m_HowManyPages - 1
            FirstControl(1, i) = intTempArray(1, i)
            FirstControl(2, i) = intTempArray(2, i)
            FirstControl(3, i) = intTempArray(3, i)
            FirstControl(4, i) = intTempArray(4, i)
            FirstControl(5, i) = intTempArray(5, i)
            FirstControl(6, i) = intTempArray(6, i)
            FirstControl(7, i) = intTempArray(7, i)
            FirstControl(8, i) = intTempArray(8, i)
            
        Next i
        
        If (my_PageNumber = currPage + 1) Then
            
            If (currPage + 1 > m_HowManyPages) Then
                currPage = m_HowManyPages - 1
                
            End If
            
            Call subSetPage(currPage + 1)
        End If
    End If
    
    Call subUpdateNav
    RaiseEvent PageChanged
    
End Sub

Private Sub subAttach(newChild)
    'If there was another PictureBox attached,
    'restore the "parenthood" for this PictureBox
    'before Attaching the new PictureBox.
    If (lPrevParent <> 0) Then
        pChild.Visible = False
        'Restore "parenthood" of the Picture Box! :)
        SetParent pChild.hwnd, lPrevParent
        
        'Release computer resources...
        Set pChild = Nothing
    End If
    
    Set pChild = newChild
    
    pChild.Visible = True
    Set pChild.Picture = m_BackPicture
    pChild.BackColor = pView.BackColor
    
    'To avoid any error, check if the container of
    'the PictureBox been attached is the ScrllngPic1.
    If (pChild.Container.hwnd = UserControl.hwnd) Then
        'Set the TabStop of the UserControl to
        'False. The only way to access the
        'TabStop property of the UserControl
        'is by changing the TabStop Property
        'of the Container of the Picture Box
        'that is been attached. Therefore, make
        'sure that the container of the Picture
        'Box been attached is the UserControl.
        pChild.Container.TabStop = False
    End If
    
    'Set the TabStop of the Picture Box
    'to False.
    pChild.TabStop = False
    
    lPrevParent = SetParent(pChild.hwnd, pView.hwnd)
    pChild.Move 0, 0
    
    Call UserControl_Resize
    Call subSelectFirst
    Call UpdatePos
    
    Call subUpdateNav
    RaiseEvent PageChanged
    
    Timer1.Enabled = True
    
End Sub

'This Timer will change the background
'color of the Control with Focus. This
'Timer will, also, check whether the
'Control with Focus is out of site or
'not. The Timer will scroll the form
'to the position of Control with Focus
'only if the Control is found to be out
'of site.
Private Sub Timer1_Timer()
    Dim Gcurrent2 As Object
    Dim intObjectTop As Integer
    Dim intTop As Integer
    Dim intObjectLeft As Integer
    Dim intLeft As Integer
    Dim intTemp As Variant
    Dim intCtrlName As String
    Dim intCtrlIndex As String
    Dim intForm As Form
    
    'This code is based on a submission by TopCoder:
    'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13566&lngWId=1
    
    If (m_HowManyPages = 0) Then
        Exit Sub
    End If
    
    If (Gcurrent Is Nothing) Then
        Set Gcurrent = UserControl.Parent.ActiveControl
        Gpast = Gcurrent.BackColor
        intChanged = True
        
    Else
        If (UserControl.Parent.WindowState = 1) Then
            Exit Sub
        End If
        
        Set Gcurrent2 = Gcurrent
        Set Gcurrent = UserControl.Parent.ActiveControl
        
        If (Gcurrent2.hwnd <> Gcurrent.hwnd) Then
            intChanged = True
        End If
        
        'There is no need to do anything if
        'the conrtrol has not been changed...
        If (intChanged) Then
            intChanged = False
            
            'Don't do anything if the previously
            'selected control was the ScrllngFrm
            'itself or the Picture1.
            If (Gcurrent2.hwnd <> UserControl.hwnd) _
            And (Gcurrent2.hwnd <> pChild.hwnd) Then
                'If you don't  want to have any
                'Picture Box to be highlighted...
                If Not (m_HighPicture) _
                And (TypeOf Gcurrent2 Is PictureBox) Then
                    'Keep going...
                    
                'Check if the object with focus has the
                'Main Picture Box as their Container...
                ElseIf (Gcurrent2.Container.hwnd = pChild.hwnd) Then
                    Gcurrent2.BackColor = Gpast
                    
                'Don't do anything if the Container of the
                'object is the form itself. Only objects
                'within the Scrolling Form Control should
                'be considered.
                ElseIf (Gcurrent2.Container.hwnd = Gcurrent2.Parent.hwnd) Then
                    'Keep going...
                    
                'The next ElseIf Statement will allow
                'for objects within a frame or another
                'Picture Box to be considered as valid
                'objects and have their BackColor changed.
                ElseIf (Gcurrent2.Container.Container.hwnd = pChild.hwnd) Then
                    Gcurrent2.BackColor = Gpast
                    
                End If
                
            End If
            
            'Don't do anything if the currently
            'selected controls is the ScrllngFrm
            'itself or the Picture1.
            If (Gcurrent.hwnd <> UserControl.hwnd) _
            And (Gcurrent.hwnd <> pChild.hwnd) Then
                
                'If you don't  want to have any
                'Picture Box to be highlighted...
                If Not (m_HighPicture) _
                And (TypeOf Gcurrent Is PictureBox) Then
                    Exit Sub
                    
                'Check if the object with focus has the
                'Main Picture Box as their Container...
                ElseIf (Gcurrent.Container.hwnd = pChild.hwnd) Then
                    
                    Gpast = Gcurrent.BackColor
                    
                    intObjectTop = Gcurrent.Top + pChild.Top
                    intObjectLeft = Gcurrent.Left + pChild.Left
                    intTop = Gcurrent.Top
                    intLeft = Gcurrent.Left
                    
                'Exit sub if the Container of the object
                'is the form itself. Only objects within
                'the Scrolling Form Control should be
                'considered.
                ElseIf (Gcurrent.Container.hwnd = Gcurrent.Parent.hwnd) Then
                    Exit Sub
                    
                'The next ElseIf Statement will allow
                'for objects within a frame or another
                'PictureBox to be considered as valid
                'objects and have their BackColor changed.
                ElseIf (Gcurrent.Container.Container.hwnd = pChild.hwnd) Then
                    Gpast = Gcurrent.BackColor
                    
                    intObjectTop = (Gcurrent.Top + Gcurrent.Container.Top) + pChild.Top
                    intObjectLeft = (Gcurrent.Left + Gcurrent.Container.Left) + pChild.Left
                    intTop = Gcurrent.Top + Gcurrent.Container.Top
                    intLeft = Gcurrent.Left + Gcurrent.Container.Left
                    
                Else
                    Exit Sub
                End If
                
                'If the user wants to highlight the
                'BackColor of the objects...
                If (m_Highlight) Then
                    Gcurrent.BackColor = m_HighlightColor
                End If
                
                'If the user wants to select the text
                'of every TextBox...
                If (m_SelectText) _
                And (TypeOf Gcurrent Is TextBox) Then
                    Gcurrent.SelStart = 0
                    Gcurrent.SelLength = Len(Gcurrent.Text)
                    
                End If
                
                RaiseEvent FocusMoved
                
                'Keep track of the currently selected
                'field for this page.
                FirstControl(2, currPage) = Gcurrent.Name
                
                'Now, I will check if the currently selected
                'object is part of an Array or not. I just
                'check if the currently selected control
                'has an Index number. If it has an Index
                'number, I will know that it is part of an
                'array. However, if it doesn't have a number,
                'it will give me an error. That's why I put
                'this error handling line. There might be a
                'cleaner way to figure it out. However, after
                'doing a thorough research, I didn't find
                'anything better then this. If you know a
                'better way of figuring out whether a control
                'is par of an array or not, please, send me
                'an e-mail or post a feedback. :)
                On Error Resume Next
                intTemp = Gcurrent.Index
                FirstControl(3, currPage) = CStr(intTemp)
                FirstControl(4, currPage) = Gcurrent.hwnd
                
                'Check if Control is out of the view...
                If ((intObjectTop + Gcurrent.Height) > VScroll.Height) Then
                    'Go down one page.
                    If ((VScroll.Value + VScroll.Height) > intTop) _
                    And (intTop < VScroll.Max) Then
                        VScroll.Value = intTop
                        Gcurrent.SetFocus
                    Else
                        VScroll.Value = VScroll.Max
                        Gcurrent.SetFocus
                    End If
                    
                ElseIf (intObjectTop < 0) Then
                    'Go up one field.
                    If (intObjectTop + (VScroll.Height + 150) > 1) _
                    And (intTop > VScroll.Min) Then
                        VScroll.Value = intTop
                        Gcurrent.SetFocus
                    Else
                        VScroll.Value = 1
                        Gcurrent.SetFocus
                    End If
                End If
                
                'Check if object is out of the view...
                If ((intObjectLeft + Gcurrent.Width) > HScroll.Width) Then
                    'Go right one screen.
                    If ((HScroll.Value + HScroll.Width) > intLeft) _
                    And (intLeft < HScroll.Max) Then
                        HScroll.Value = intLeft
                        Gcurrent.SetFocus
                    Else
                        HScroll.Value = HScroll.Max
                        Gcurrent.SetFocus
                    End If
                    
                ElseIf (intObjectLeft < 0) Then
                    'Go left one field.
                    If (intObjectLeft + (HScroll.Width + 150) > 1) _
                    And (intLeft > HScroll.Min) Then
                        HScroll.Value = intLeft
                        Gcurrent.SetFocus
                    Else
                        HScroll.Value = 1
                        Gcurrent.SetFocus
                    End If
                End If
                
                'Memorize Scroll Bar position for current page.
                FirstControl(5, currPage) = VScroll.Value
                FirstControl(6, currPage) = HScroll.Value
            End If
        End If
    End If

End Sub

'The following sub will go trough all the visible
'controls of the form and will try to figure which
'ones are within the PictureBox that corresponds
'to the current page. Then, the sub will try to
'determine which control has the smallest TabIndex
'and will select this control as long as this is the
'first time the user is opening the current page
'(PuctureBox). However, if this is not the first
'time that the user has visited the current page
'(PictureBox), this sub will select the Control
'last selected on the current page.
Private Sub subSelectFirst()
    Dim intForm As Form
    Dim intControl As Control
    Dim intContName As String
    Dim intContIndex As String
    Dim intContHWnd As Long
    Dim intContTabInd As Integer
    Dim intContTabStop As Boolean
    Dim intFirstControl As Boolean
    Dim intPage As Integer
    Dim intHasHWnd As Variant
    Dim intIsNewPage As Boolean
    Dim intTempIndex As String
    Dim intContainerName As String
    Dim intContanierIndex As String
    Dim intContanier As Control
    Dim i As Integer
    
    intFirstControl = True
    
    'This If Then statement will only be true if there
    'was no page (PictureBox) added to the Control.
    If (m_HowManyPages = 0) Then
        If Not (intSetFocus) Then
            m_HowManyPages = 1
        Else
            intSetFocus = False
            Exit Sub
        End If
    End If
    
    'Find out whether the user is adding a new page
    'or is trying to go to a page that has already
    'been added.
    For i = 0 To m_HowManyPages - 1
        If (Len(FirstControl(1, i)) = 0) _
        And (i = 0) Then
            intPage = 0
            intIsNewPage = True
            Exit For
            
        ElseIf (FirstControl(1, i) = pChild.hwnd) Then
            intPage = i
            Exit For
            
        ElseIf (i = m_HowManyPages - 1) Then
            intPage = m_HowManyPages
            intIsNewPage = True
            
        End If
        
    Next i
    
    'If another page is added, re-dimension
    'the array to hold this extra information.
    If (intPage = m_HowManyPages) Then
        ReDim Preserve FirstControl(8, intPage + 1)
        m_HowManyPages = m_HowManyPages + 1
    End If
    
    'Update the current page...
    currPage = intPage
    
    'Get a hold on the Form where this
    'UserControl is.
    Set intForm = pChild.Parent
    
    'Loop through all the controls on
    'the Parent Form.
    For i = 0 To intForm.Controls.Count - 1
        Set intControl = intForm.Controls(i)
        
        'Some properties may not be available, depending on
        'the control that is selected. For example, if the
        'selected control is a Timer, you cannot check its
        'type without getting an error. As a result, I added
        'the following error handling line. It is worth
        'mentioning that I didn't use the common error
        'handling line "On Error Resume Next". If I did,
        'every time I got an error on the If Then conditional
        'I would have an undesirable behavior. Instead of
        'skipping the conditional, the process would go right
        'through it! To avoid this, I stated on my error
        'handling line that I wanted the process to jump to
        'the end of the For Next loop (thus skipping the If
        'Then statement) and, then, Resume Next from there.
        On Error GoTo NextLoop
        If (intControl.Visible) _
        And Not (TypeOf intControl Is Label) _
        And Not (TypeOf intControl Is Frame) Then
            
            'There are some controls like Labels that
            'don't have the hWnd property. I used this
            'error handling to bypass an error that
            'was going to occur if I tried to read this
            'property form such controls.
            On Error Resume Next
            intHasHWnd = Empty
            intHasHWnd = intControl.hwnd
            
            'If Control has a Windows handle number (hWnd)...
            If Not (IsEmpty(intHasHWnd)) Then
                'I was having problems when the selected
                'Control was a Timer. The Timer has no container.
                'Therefore, I needed to check the name of the
                'container before going any further.
                intContainerName = ""
                intContainerName = CStr(intControl.Container.Name)
                
                If (Len(intContainerName) > 0) Then
                    
                    'The control should be the contained by
                    'the PictureBox (Current Page).
                    If (intControl.Container.hwnd <> intForm.hwnd) Then
                        If (intControl.Container.hwnd = pChild.hwnd) Then
                            
                            'There are some controls like Frames and
                            'Labels that don't have the TabStop property.
                            'If I tried to read this property from
                            'such controls, I would get an error message.
                            'However, because I wrote the error handling
                            'line before, this error will be bypassed.
                            intContTabStop = False
                            intContTabStop = intControl.TabStop
                            If (intFirstControl) _
                            And (intContTabStop) Then
                                intContTabInd = intControl.TabIndex
                                intFirstControl = False
                                
                            End If
                            
                            'Check whether the currently selected control
                            'has a TabIndex smaller then the previously
                            'saved control or not.
                            If (intControl.TabIndex <= intContTabInd) _
                            And (intContTabStop) Then
                                intContTabStop = False
                                intContTabInd = intControl.TabIndex
                                intContName = intControl.Name
                                intContHWnd = intControl.hwnd
                                
                                intContIndex = ""
                                intContIndex = CStr(intControl.Index)
                                
                            End If
                            
                        'If the container of the currently selected
                        'control is not the PictureBox(Current Page),
                        'select the container of this Control and verify
                        'if the container of this container is the
                        'PictureBox. If so, it will mean that, even
                        'though this control is not directly an object
                        'contained by the Page (PictureBox), it will
                        'still be considered as a member of the page
                        'because its container is part of the PictueBox
                        'collection.
                        'Summarizing, this function will allow
                        'the user to use frames on each page.
                        Else
                            
                            'Now, I will check if the currently selected
                            'object is part of an Array or not. I just
                            'check if the currently selected control has
                            'an Index number. If it has an Index number,
                            'I will know that it is part of an array. However,
                            'if it doesn't have a number, it will give me an
                            'error. That's why I put this error handling line.
                            'There might be a cleaner way to figure it out.
                            'However, after doing a thorough research, I didn't
                            'find anything better then this. If you know a
                            'better way of figuring out whether a control is
                            'par of an array or not, please, send me an e-mail
                            'or post a feedback. :)
                            intContanierIndex = ""
                            intContanierIndex = CStr(intControl.Container.Index)
                            
                            If (Len(intContanierIndex) = 0) Then
                                'Select the control using the following statement
                                'if the control is NOT part of an array.
                                Set intContanier = intForm.Controls(intContainerName)
                                
                            Else
                                'Select the control using the following statement
                                'if the control is part of an array.
                                Set intContanier = intForm.Controls(intContainerName).Item(CInt(intContanierIndex))
                                
                            End If
                            
                            If (intContanier.Container.hwnd = pChild.hwnd) Then
                                
                                'There are some controls like Frames and Labels
                                'that don't have the TabStop property. If I tried
                                'to read this property from such controls, I would
                                'get an error message. However, because I wrote
                                'the error handling line before, this error will
                                'be bypassed.
                                intContTabStop = False
                                intContTabStop = intControl.TabStop
                                
                                If (intFirstControl) _
                                And (intContTabStop) Then
                                    
                                    intContTabInd = intControl.TabIndex
                                    intFirstControl = False
                                    
                                End If
                                
                                'Check whether the currently selected control has
                                'a TabIndex smaller then the previously saved
                                'control or not.
                                If (intControl.TabIndex <= intContTabInd) _
                                And (intContTabStop) Then
                                    intContTabStop = False
                                    intContTabInd = intControl.TabIndex
                                    intContName = intControl.Name
                                    intContHWnd = intControl.hwnd
                                    
                                    'Now, I will check if the currently selected
                                    'object is part of an Array or not. I just check
                                    'if the currently selected control has an Index
                                    'number. If it has an Index number, I will know
                                    'that it is part of an array. However, if it
                                    'doesn't have a number, it will give me an error.
                                    'That's why I put this error handling line. There
                                    'might be a cleaner way to figure it out. However,
                                    'after doing a thorough research, I didn't find
                                    'anything better then this. If you know a better
                                    'way of figuring out whether a control is par of
                                    'an array or not, please, send me an e-mail or
                                    'post a feedback. :)
                                    intContIndex = ""
                                    intContIndex = CStr(intControl.Index)
                                    
                                End If
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Resume Next
    
    Next i
    
    'If it is a page that has just been added,
    'select the control with the smallest TabIndex
    'contained by the current page (PictureBox).
    If ((intIsNewPage) _
    Or (FirstControl(2, intPage) = "This is a new Page.")) Then
        FirstControl(1, intPage) = pChild.hwnd
        FirstControl(2, intPage) = intContName
        FirstControl(3, intPage) = intContIndex
        FirstControl(4, intPage) = intContHWnd
        FirstControl(5, intPage) = 1
        FirstControl(6, intPage) = 1
        FirstControl(7, intPage) = pChild.Name
        
        'Now, I will check if the currently selected
        'object is part of an Array or not. I just check
        'if the currently selected control has an Index
        'number. If it has an Index number, I will know
        'that it is part of an array. However, if it
        'doesn't have a number, it will give me an error.
        'That's why I put this error handling line. There
        'might be a cleaner way to figure it out. However,
        'after doing a thorough research, I didn't find
        'anything better then this. If you know a better
        'way of figuring out whether a control is par of
        'an array or not, please, send me an e-mail or
        'post a feedback. :)
        On Error Resume Next
        intTempIndex = ""
        intTempIndex = CStr(pChild.Index)
        FirstControl(8, intPage) = intTempIndex
        
        VScroll.Value = FirstControl(5, intPage)
        HScroll.Value = FirstControl(6, intPage)
    
    'If it is a page that has been visited before,
    'select the control was last select on the current
    'page (PictureBox).
    Else
        'If the user chose not to have the UserControl
        'memorizing the position of the Scroll Bars on
        'last visit to the page, reset these values on
        'the array.
        If Not (m_MemorizeScroll) Then
            FirstControl(5, intPage) = 1
            FirstControl(6, intPage) = 1
        End If
        
        'If the user choses not to have the UserControl
        'memorizing the last selected Control on last
        'visit to the page current page, select the control
        'with the smallest TabIndex contained by the
        'current page (PictureBox).
        If Not (m_MemorizeField) Then
            FirstControl(2, intPage) = intContName
            FirstControl(3, intPage) = intContIndex
        End If
        
        intContName = FirstControl(2, intPage)
        intContIndex = FirstControl(3, intPage)
        VScroll.Value = FirstControl(5, intPage)
        HScroll.Value = FirstControl(6, intPage)
    End If
    
    
    If (Len(intContIndex) = 0) Then
        'Select the control using the following
        'statement if the control is NOT part of
        'an array.
        intForm.Controls(intContName).SetFocus
        
    Else
        'Select the control using the following
        'statement if the control is part of an
        'array.
        Set intControl = intForm.Controls(intContName).Item(CInt(intContIndex))
        intControl.SetFocus
        
    End If
    
End Sub

'======================================
'==== The following subs are used  ====
'==== to navigate through each of  ====
'==== the added pages (PictureBox) ====
'======================================

'Validate the NextPage request and call the
'subCallAttach sub only if going to next page
'is possible.
Public Sub NextPage()
    Dim intPage As Integer
    
    If Not (Len(FirstControl(1, 0)) > 0) Then
        Exit Sub
        
    ElseIf (currPage + 1 < m_HowManyPages) Then
        intPage = currPage + 1
        
    Else
        intPage = currPage
        
    End If
    
    Call subCallAttach(intPage)
    
End Sub

'Validate the PreviousPage request and call
'the subCallAttach sub only if going to
'previous page is possible.
Public Sub PreviousPage()
    Dim intPage As Integer
    
    If Not (Len(FirstControl(1, 0)) > 0) Then
        Exit Sub
        
    ElseIf (currPage > 0) Then
        intPage = currPage - 1
        
    Else
        intPage = currPage
        
    End If
    
    Call subCallAttach(intPage)
    
End Sub

'Validate the FirstPage request and call the
'subCallAttach sub only if going to next
'page is possible.
Public Sub FirstPage()
    Dim intPage As Integer
    
    If Not (Len(FirstControl(1, 0)) > 0) Then
        Exit Sub
    End If
    
    Call subCallAttach(intPage)
    
End Sub

'Validate the LastPage request and call the
'subCallAttach sub only if going to last
'page is possible.
Public Sub LastPage()
    Dim intPage As Integer
    
    If Not (Len(FirstControl(1, 0)) > 0) Then
        Exit Sub
        
    ElseIf (currPage < m_HowManyPages - 1) Then
        intPage = m_HowManyPages - 1
        
    Else
        intPage = currPage
        
    End If
    
    Call subCallAttach(intPage)
    
End Sub

'When the user sets a new number for the
'CurrentPage property, this sub is called.
'At this point validate the SetPage request
'and call the subCallAttach sub only if
'going to the requested page is possible.
Private Sub subSetPage(m_Page As Integer)
    Dim intPage As Integer
    
    If Not (Len(FirstControl(1, 0)) > 0) Then
        Exit Sub
        
    ElseIf (m_HowManyPages > 1) _
    And (m_Page <> currPage + 1) Then
        If (m_Page > m_HowManyPages) Then
            intPage = m_HowManyPages - 1
            
        ElseIf (m_Page < 1) Then
            intPage = 0
            
        Else
            intPage = m_Page - 1
            
        End If
    Else
        intPage = currPage
        
    End If
    
    Call subCallAttach(intPage)
    
End Sub

'This sub will select the PictureBox
'corresponding to the "intPage" page number
'and will call the attach function.
Public Sub subCallAttach(intPage As Integer)
    Dim intControl As Control
    Dim intForm As Form
    Dim intPicBoxName As String
    Dim intPicBoxIndex As Integer
    
    Set intForm = pChild.Parent
    intPicBoxName = FirstControl(7, intPage)
    
    If (Len(FirstControl(8, intPage)) = 0) Then
        Set intControl = intForm.Controls(intPicBoxName)
        
    'Execute the following lines of code only
    'if the Picture box is a member of a
    'PictuerBox Array.
    Else
        intPicBoxIndex = CInt(FirstControl(8, intPage))
        Set intControl = intForm.Controls(intPicBoxName).Item(intPicBoxIndex)
        
    End If
    
    Call subAttach(intControl)
    
End Sub

'This sub will enable or disable the buttons
'that will be used to navigate to all the
'pages added.
Private Sub subUpdateNav()
    If (m_HowManyPages > 1) _
    And (currPage + 1 < m_HowManyPages) Then
        m_NextEnabled = True
        
    Else
        m_NextEnabled = False
        
    End If
    
    If (m_HowManyPages > 1) _
    And (currPage + 1 > 1) Then
        m_PreviousEnabled = True
        
    Else
        m_PreviousEnabled = False
        
    End If
    
End Sub

'==================================================
'======= Following are the Subs that will    ======
'======= initialize and save the Properties. ======
'==================================================

'Get property values from property bags...
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '============================
    'These properties are related
    'to UserControl appearance.
    '============================
    pView.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set m_BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
    Set pView.Picture = m_BackPicture
    
    '============================
    'These properties are related
    'to field selection behavior.
    '============================
    m_Highlight = PropBag.ReadProperty("Highlight", m_def_Highlight)
    m_HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
    m_HighPicture = PropBag.ReadProperty("HighPicture", m_def_HighPicture)
    m_SelectText = PropBag.ReadProperty("SelectText", m_def_SelectText)
    m_MemorizeField = PropBag.ReadProperty("MemorizeField", m_def_MemorizeField)
    m_MemorizeScroll = PropBag.ReadProperty("MemorizeScroll", m_def_MemorizeScroll)
    
    '============================
    'These properties are related
    'to page navigation.
    '============================
    m_CurrentPage = PropBag.ReadProperty("CurrentPage", m_def_CurrentPage)
    m_HowManyPages = PropBag.ReadProperty("HowManyPages", m_def_HowManyPages)
    m_NextEnabled = PropBag.ReadProperty("NextEnabled", m_def_NextEnabled)
    m_PreviousEnabled = PropBag.ReadProperty("PreviousEnabled", m_def_PreviousEnabled)
    
End Sub

'Write the property values to the property bags...
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '============================
    'These properties are related
    'to UserControl appearance.
    '============================
    Call PropBag.WriteProperty("BackPicture", m_BackPicture, Nothing)
    Call PropBag.WriteProperty("BackColor", pView.BackColor, &H8000000F)
    
    '============================
    'These properties are related
    'to field selection behavior.
    '============================
    Call PropBag.WriteProperty("Highlight", m_Highlight, m_def_Highlight)
    Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)
    Call PropBag.WriteProperty("HighPicture", m_HighPicture, m_def_HighPicture)
    Call PropBag.WriteProperty("SelectText", m_SelectText, m_def_SelectText)
    Call PropBag.WriteProperty("MemorizeField", m_MemorizeField, m_def_MemorizeField)
    Call PropBag.WriteProperty("MemorizeScroll", m_MemorizeScroll, m_def_MemorizeScroll)
    
    '============================
    'These properties are related
    'to page navigation.
    '============================
    Call PropBag.WriteProperty("NextEnabled", m_NextEnabled, m_def_NextEnabled)
    Call PropBag.WriteProperty("PreviousEnabled", m_PreviousEnabled, m_def_PreviousEnabled)
    Call PropBag.WriteProperty("CurrentPage", m_CurrentPage, m_def_CurrentPage)
    Call PropBag.WriteProperty("HowManyPages", m_HowManyPages, m_def_HowManyPages)
    
End Sub

Private Sub UserControl_InitProperties()
    '============================
    'These properties are related
    'to UserControl appearance.
    '============================
    Set m_BackPicture = LoadPicture("")
    
    '============================
    'These properties are related
    'to field selection behavior.
    '============================
    m_HighlightColor = m_def_HighlightColor
    m_Highlight = m_def_Highlight
    m_HighPicture = m_def_HighPicture
    m_SelectText = m_def_SelectText
    m_MemorizeField = m_def_MemorizeField
    m_MemorizeScroll = m_def_MemorizeScroll
    
    '============================
    'These properties are related
    'to page navigation.
    '============================
    m_CurrentPage = m_def_CurrentPage
    m_HowManyPages = m_def_HowManyPages
    m_NextEnabled = m_def_NextEnabled
    m_PreviousEnabled = m_def_PreviousEnabled
    
End Sub

'============================================
'======= Following is the description ======
'======= of the Public Properties.    ======
'============================================

'############################
'These properties are related
'to UserControl appearance.
'############################

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pView,pView,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of the ScrllngFrm Control. This color will be used as the backgroud color of the added pages at Run Time."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = pView.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    pView.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,2,2,0
Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "Returns/sets the background picture of the ScrllngFrm Control. This picture will be used as the backgroud picture for the added pages at Run Time."
Attribute BackPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'If Ambient.UserMode Then Err.Raise 393
    Set BackPicture = m_BackPicture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
    'If Ambient.UserMode = False Then Err.Raise 383
    Set m_BackPicture = New_BackPicture
    Set pView.Picture = New_BackPicture
    PropertyChanged "BackPicture"
End Property

'############################
'These properties are related
'to field selection behavior.
'############################

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Highlight() As Boolean
Attribute Highlight.VB_Description = "Returns/set a value that determines whether the control on focus will have its background color changed or not."
Attribute Highlight.VB_ProcData.VB_Invoke_Property = "Behavior;Behavior"
    Highlight = m_Highlight
End Property

Public Property Let Highlight(ByVal New_Highlight As Boolean)
    m_Highlight = New_Highlight
    PropertyChanged "Highlight"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFC0C0&
Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Returns/sets the color that will be used to highlight the Control with focus."
Attribute HighlightColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal New_HighlightColor As OLE_COLOR)
    m_HighlightColor = New_HighlightColor
    PropertyChanged "HighlightColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get HighPicture() As Boolean
Attribute HighPicture.VB_Description = "Returns/set a value that determines whether any Picture Box on focus will have its background color changed or not."
Attribute HighPicture.VB_ProcData.VB_Invoke_Property = "Behavior;Behavior"
    HighPicture = m_HighPicture
    
End Property

Public Property Let HighPicture(ByVal New_HighPicture As Boolean)
    m_HighPicture = New_HighPicture
    PropertyChanged "HighPicture"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get SelectText() As Boolean
Attribute SelectText.VB_Description = "Returns/set a value that determines whether any Text Box on focus will have its text selected."
Attribute SelectText.VB_ProcData.VB_Invoke_Property = "Behavior;Behavior"
    SelectText = m_SelectText
    
End Property

Public Property Let SelectText(ByVal New_SelectText As Boolean)
    m_SelectText = New_SelectText
    PropertyChanged "SelectText"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MemorizeField() As Boolean
Attribute MemorizeField.VB_Description = "Returns/set a value that determines whether the last field selected on each page will get focus on the next time that the corresponding page is opened."
Attribute MemorizeField.VB_ProcData.VB_Invoke_Property = "Behavior;Behavior"
    MemorizeField = m_MemorizeField
End Property

Public Property Let MemorizeField(ByVal New_MemorizeField As Boolean)
    m_MemorizeField = New_MemorizeField
    PropertyChanged "MemorizeField"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MemorizeScroll() As Boolean
Attribute MemorizeScroll.VB_Description = "Returns/set a value that determines whether the last position of the scroll bars on each page will be restore on the next time that the corresponding page is opened."
Attribute MemorizeScroll.VB_ProcData.VB_Invoke_Property = "Behavior;Behavior"
    MemorizeScroll = m_MemorizeScroll
    
End Property

Public Property Let MemorizeScroll(ByVal New_MemorizeScroll As Boolean)
    m_MemorizeScroll = New_MemorizeScroll
    PropertyChanged "MemorizeScroll"
    
End Property

'############################
'These properties are related
'to page navigation.
'############################

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CurrentPage() As Integer
Attribute CurrentPage.VB_Description = "Returns/set the current page been viewed at Run Time only."
Attribute CurrentPage.VB_ProcData.VB_Invoke_Property = "Navigation;Behavior"
    If (m_HowManyPages > 0) Then
        CurrentPage = currPage + 1 'm_CurrentPage
    Else
        CurrentPage = 0
    End If
End Property

Public Property Let CurrentPage(ByVal New_CurrentPage As Integer)
    Call subSetPage(New_CurrentPage)
    
    PropertyChanged "CurrentPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HowManyPages() As Integer
Attribute HowManyPages.VB_Description = "Returns the total number of pages added to the ScrllingFrm Control at Run Time."
Attribute HowManyPages.VB_ProcData.VB_Invoke_Property = "Navigation;Behavior"
    HowManyPages = m_HowManyPages
End Property

Public Property Let HowManyPages(ByVal New_HowManyPages As Integer)
    m_HowManyPages = New_HowManyPages
    PropertyChanged "HowManyPages"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,1,False
Public Property Get NextEnabled() As Boolean
Attribute NextEnabled.VB_Description = "Returns True if there is more then one page added to the ScrllingFrm Control and the current page is not the first page."
Attribute NextEnabled.VB_ProcData.VB_Invoke_Property = "Navigation;Behavior"
    NextEnabled = m_NextEnabled
End Property

Public Property Let NextEnabled(ByVal New_NextEnabled As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_NextEnabled = New_NextEnabled
    PropertyChanged "NextEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,1,False
Public Property Get PreviousEnabled() As Boolean
Attribute PreviousEnabled.VB_Description = "Returns True if there is more then one page added to the ScrllingFrm Control and the current page is not the last page."
Attribute PreviousEnabled.VB_ProcData.VB_Invoke_Property = "Navigation;Behavior"
    PreviousEnabled = m_PreviousEnabled
End Property

Public Property Let PreviousEnabled(ByVal New_PreviousEnabled As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_PreviousEnabled = New_PreviousEnabled
    PropertyChanged "PreviousEnabled"
End Property
