VERSION 5.00
Begin VB.UserControl InvFrame 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   9
      X2              =   9
      Y1              =   232
      Y2              =   8
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   8
      X2              =   8
      Y1              =   232
      Y2              =   8
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   8
      X2              =   311
      Y1              =   233
      Y2              =   233
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   8
      X2              =   311
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   312
      X2              =   312
      Y1              =   232
      Y2              =   8
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   311
      X2              =   311
      Y1              =   232
      Y2              =   8
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   56
      X2              =   312
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   56
      X2              =   312
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   8
      X2              =   24
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   8
      X2              =   24
      Y1              =   8
      Y2              =   8
   End
End
Attribute VB_Name = "InvFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const CapLeft = 20
Const CapTop = 0
Dim CapHeight As Long
Dim CapWidth As Long

Public Enum Vis
    Transparent = 0
    Visible = 1
End Enum

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
'Default Property Values:
Const m_def_Caption = "InvFrame"
Const m_def_Enabled = True
'Property Variables:
Dim m_Caption As String
Dim m_Enabled As Boolean




Private Sub UserControl_Resize()
    Line1(0).X1 = 0
    Line1(0).X2 = CapLeft - 2
    Line1(0).Y1 = CapHeight / 2
    Line1(0).Y2 = CapHeight / 2
    Line1(1).X1 = CapLeft + CapWidth + 2
    Line1(1).X2 = UserControl.ScaleWidth - 2
    Line1(1).Y1 = CapHeight / 2
    Line1(1).Y2 = CapHeight / 2
    Line1(2).X1 = Line1(1).X2
    Line1(2).X2 = Line1(1).X2
    Line1(2).Y1 = Line1(1).Y2
    Line1(2).Y2 = UserControl.ScaleHeight - 2
    Line1(3).X1 = Line1(2).X2
    Line1(3).X2 = 0
    Line1(3).Y1 = Line1(2).Y2
    Line1(3).Y2 = Line1(2).Y2
    Line1(4).X1 = Line1(3).X2
    Line1(4).X2 = Line1(3).X2
    Line1(4).Y1 = Line1(3).Y2
    Line1(4).Y2 = Line1(0).Y1

    Line2(0).X1 = Line1(0).X1 + 1
    Line2(0).X2 = Line1(0).X2
    Line2(0).Y1 = Line1(0).Y1 + 1
    Line2(0).Y2 = Line1(0).Y1 + 1
    Line2(1).X1 = Line1(1).X1
    Line2(1).X2 = Line1(1).X2
    Line2(1).Y1 = Line1(1).Y1 + 1
    Line2(1).Y2 = Line1(1).Y2 + 1
    Line2(2).X1 = Line1(2).X1 + 1
    Line2(2).X2 = Line1(2).X2 + 1
    Line2(2).Y1 = Line1(2).Y1
    Line2(2).Y2 = Line1(2).Y2 + 1
    Line2(3).X1 = Line1(3).X1 + 1
    Line2(3).X2 = Line1(3).X2 - 1
    Line2(3).Y1 = Line1(3).Y1 + 1
    Line2(3).Y2 = Line1(3).Y2 + 1
    Line2(4).X1 = Line1(4).X1 + 1
    Line2(4).X2 = Line1(4).X2 + 1
    Line2(4).Y1 = Line1(4).Y1 - 1
    Line2(4).Y2 = Line1(4).Y2 + 1

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,InvFrame
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Dim TextSize As POINTAPI
    Label1.Caption = New_Caption
    Label1.Left = CapLeft
    GetTextExtentPoint32 UserControl.hdc, m_Caption, Len(m_Caption), TextSize
    CapWidth = TextSize.X
    CapHeight = TextSize.Y
    UserControl.Cls
    UserControl.CurrentX = CapLeft
    UserControl.CurrentY = CapTop
    UserControl.Print m_Caption
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = New_Enabled
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Me.Caption = m_def_Caption
    m_Enabled = m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Me.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Vis
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Vis)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

