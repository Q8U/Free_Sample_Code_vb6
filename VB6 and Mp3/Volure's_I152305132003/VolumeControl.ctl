VERSION 5.00
Begin VB.UserControl Sound2 
   CanGetFocus     =   0   'False
   ClientHeight    =   3597
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4796
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "VolumeControl.ctx":0000
   ScaleHeight     =   3597
   ScaleWidth      =   4796
   ToolboxBitmap   =   "VolumeControl.ctx":0282
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   360
   End
End
Attribute VB_Name = "Sound2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Enumerations:
Public Enum ControlDevices
    mMaster = 0
    mLineIn = 1
    mMicrophone = 2
    mSynthesizer = 3
    mCD = 4
    mWave = 5
    mAuxiliary = 6
End Enum
'Default Property Values:
Const m_def_Mute = False
Const m_def_DeviceToControl = mWave
Const m_def_Volume = 0
'Property Variables:
Dim m_Mute As Boolean
Dim m_DeviceToControl As ControlDevices
Dim m_Volume As Long
'Event Declarations:
Event MuteChanged(NewMute As Boolean)
Event VolumeChanged(NewVolume As Long)




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
UserControl.Enabled() = New_Enabled
PropertyChanged "Enabled"
Timer1.Enabled = New_Enabled

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Volume() As Long
Volume = m_Volume

End Property

Public Property Let Volume(ByVal New_Volume As Long)
If UserControl.Enabled Then
    m_Volume = New_Volume
    SetVolume m_DeviceToControl, m_Volume
    PropertyChanged "Volume"
End If

End Property

Private Sub Timer1_Timer()
Dim vv As Long
Dim mm As Boolean

vv = GetVolume(m_DeviceToControl)
If vv <> m_Volume Then
    RaiseEvent VolumeChanged(vv)
End If
m_Volume = vv
mm = GetMute(m_DeviceToControl + 7)
If mm <> m_Mute Then
    RaiseEvent MuteChanged(mm)
End If
m_Mute = mm

End Sub

Private Sub UserControl_Initialize()
OpenMixer (0)
UserControl.Width = 480
UserControl.Height = 480

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
m_DeviceToControl = m_def_DeviceToControl
m_Volume = GetVolume(m_DeviceToControl)
m_Mute = GetMute(m_DeviceToControl + 7)
UserControl.Width = 480
UserControl.Height = 480

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
m_Volume = PropBag.ReadProperty("Volume", m_def_Volume)
m_Mute = PropBag.ReadProperty("mute", m_def_Mute)
m_DeviceToControl = PropBag.ReadProperty("DeviceToControl", m_def_DeviceToControl)
Timer1.Enabled = UserControl.Enabled

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 480
UserControl.Height = 480

End Sub

Private Sub UserControl_Terminate()
CloseMixer

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("Volume", m_Volume, m_def_Volume)
Call PropBag.WriteProperty("mute", m_Mute, m_def_Mute)
Call PropBag.WriteProperty("DeviceToControl", m_DeviceToControl, m_def_DeviceToControl)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get Mute() As Boolean
Mute = m_Mute

End Property

Public Property Let Mute(ByVal New_mute As Boolean)
If UserControl.Enabled Then
    m_Mute = New_mute
    SetMute m_DeviceToControl + 7, m_Mute
    PropertyChanged "Mute"
End If

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get DeviceToControl() As ControlDevices
DeviceToControl = m_DeviceToControl

End Property

Public Property Let DeviceToControl(ByVal New_DeviceToControl As ControlDevices)
If UserControl.Enabled Then
    m_DeviceToControl = New_DeviceToControl
    PropertyChanged "DeviceToControl"
    'whenever device is changed we have to change
    'a value for volume and mute
    m_Volume = GetVolume(m_DeviceToControl)
    RaiseEvent VolumeChanged(m_Volume)
    m_Mute = GetMute(m_DeviceToControl + 7)
    RaiseEvent MuteChanged(m_Mute)
End If

End Property
