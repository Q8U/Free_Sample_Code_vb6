Attribute VB_Name = "Volume"
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&

Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
''''''''''''''
'   Tipi di componenti per destinazioni e origini

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)

Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
''''''''''''''''''''''


'''''''''''''
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000 '  in decimi
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000 '  in decimi
''''''''''''''''
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)

'''''''''''''''''''''
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
''''''''''''''''''''''''''''''''''''''''

Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Type MIXERCAPS
   wMid As Integer                   '  manufacturer id
   wPid As Integer                   '  product id
   vDriverVersion As Long            '  version of the driver
   szPname As String * MAXPNAMELEN   '  product name
   fdwSupport As Long                '  misc. support bits
   cDestinations As Long             '  count of destinations
End Type

Type MIXERCONTROL
   cbStruct As Long           '  size in Byte of MIXERCONTROL
   dwControlID As Long        '  unique control id for mixer device
   dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
   fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
   cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
   szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
   lMinimum As Long           '  Minimum value
   lMaximum As Long           '  Maximum value
   reserved(10) As Long       '  reserved structure space
   End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

'   MIXER_GETCONTROLDETAILSF_VALUE

Type MIXERCONTROLDETAILS_BOOLEAN
        fValue As Long
End Type

Type MIXERCONTROLDETAILS_SIGNED
        lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
        dwValue As Long
End Type

Type MIXERLINE
   cbStruct As Long               '  size of MIXERLINE structure
   dwDestination As Long          '  zero based destination index
   dwSource As Long               '  zero based source index (if source)
   dwLineID As Long               '  unique line id for mixer device
   fdwLine As Long                '  state/information about line
   dwUser As Long                 '  driver specific information
   dwComponentType As Long        '  component type line connects to
   cChannels As Long              '  number of channels line supports
   cConnections As Long           '  number of connections (possible)
   cControls As Long              '  number of controls at this line
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
   dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                          '  MIXER_GETLINECONTROLSF_ONEBYID or
   dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
   cControls As Long      '  count of controls pmxctrl points to
   cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
   pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

Public Lvu As Long
Public Rvu As Long
Public Lvol As Long
Public Rvol As Long
Public Lrecvu As Long
Public Rrecvu As Long
Public volumeout As Long
Function GetWaveInMute(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = MIXERCONTROL_CONTROLTYPE_MIXER
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)

       If (MMSYSERR_NOERROR = rc) Then
           GetWaveInMute = True
           
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetWaveInMute = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If

   GetWaveInMute = False
   
End Function

Function GetVuControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
   Dim ac As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   'If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       '/////////////////////////////////////////////////////

mxcd.cChannels = mxl.cChannels
   
mxcd.item = mxc.cMultipleItems
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol(1))

   ' Allocate a buffer for the control value buffer
hmem = GlobalAlloc(&H40, Len(vol(1)))
mxcd.paDetails = GlobalLock(hmem)

   ' Copy the data into the control value buffer

'///////////////////////////////////////////////////

       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       
       ac = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      
         ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       
            CopyStructFromPtr vol(1).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
            CopyStructFromPtr vol(0).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
         
       GlobalFree (hmem)
       GetVuControl = vol(0).dwValue
       
       Lvu = vol(0).dwValue
       Rvu = vol(1).dwValue

End Function


Function GetVolumeControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long
                        
' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
   Dim ac As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       '/////////////////////////////////////////////////////
       
       mxcd.cChannels = mxl.cChannels

   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(1))
   
   ' Allocate a buffer for the control value buffer
   hmem = GlobalAlloc(&H40, Len(vol(1)))
   mxcd.paDetails = GlobalLock(hmem)
   
   
   'vol.dwValue = volume
   
   ' Copy the data into the control value buffer
   
   
   '///////////////////////////////////////////////////
   
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       ac = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
       If (MMSYSERR_NOERROR = rc) Then
           GetVolumeControl = True
       
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
            CopyStructFromPtr vol(1).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
         CopyStructFromPtr vol(0).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
       Else
           GetVolumeControl = False
       End If
       GlobalFree (hmem)
       GetVolumeControl = vol(1).dwValue
       Lvol = vol(1).dwValue
       Rvol = vol(0).dwValue
       Exit Function
   End If

   GetVolumeControl = False

End Function

Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
                        
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED

   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(1))
   
   ' Allocate a buffer for the control value buffer
   
   mxcd.cChannels = 2
   
   hmem = GlobalAlloc(&H40, Len(vol(1)))
   mxcd.paDetails = GlobalLock(hmem)
   
   vol(1).dwValue = volR
   vol(0).dwValue = volL
  
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
   CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
   
End Function



Function GetInitVolume(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
            
       If (MMSYSERR_NOERROR = rc) Then
           GetInitVolume = True
           
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetInitVolume = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If

   GetInitVolume = False
   
End Function

Function getrecvucontrol(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
   Dim ac As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   'If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       '/////////////////////////////////////////////////////

mxcd.cChannels = mxl.cChannels
   
mxcd.item = mxc.cMultipleItems
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol(1))

   ' Allocate a buffer for the control value buffer
hmem = GlobalAlloc(&H40, Len(vol(1)))
mxcd.paDetails = GlobalLock(hmem)

   ' Copy the data into the control value buffer

'///////////////////////////////////////////////////
'Stop
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       
       ac = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      
         ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       
            CopyStructFromPtr vol(1).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
            CopyStructFromPtr vol(0).dwValue, mxcd.paDetails, Len(vol(1)) * mxcd.cChannels
         
       GlobalFree (hmem)
       getrecvucontrol = vol(0).dwValue
       
       Lrecvu = vol(0).dwValue
       Rrecvu = vol(1).dwValue

End Function

Function setoutvolumecontrol(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
                        
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED

   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol)
   
   ' Allocate a buffer for the control value buffer
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   mxcd.cChannels = 1
   vol.dwValue = volume
   
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol.dwValue, Len(vol)
  
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       setoutvolumecontrol = True
   Else
       setoutvolumecontrol = False
   End If
   
End Function

Function getinitvolumeout(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
            
       If (MMSYSERR_NOERROR = rc) Then
           getinitvolumeout = True
           
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           getinitvolumeout = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If

   getinitvolumeout = False
   
End Function

Function getoutvolumecontrol(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long

' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
   Dim ac As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = mxl.cControls
       mxlc.cbmxctrl = Len(mxc)
       
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       '/////////////////////////////////////////////////////

   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol)
   
   ' Allocate a buffer for the control value buffer
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   
   mxcd.cChannels = 1
   'vol.dwValue = volume
   
   ' Copy the data into the control value buffer
   
   
   '///////////////////////////////////////////////////
   
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       ac = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
       If (MMSYSERR_NOERROR = rc) Then
           getoutvolumecontrol = True
       
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
            CopyStructFromPtr vol.dwValue, mxcd.paDetails, Len(vol)
         
       Else
           getoutvolumecontrol = False
       End If
       GlobalFree (hmem)
       getoutvolumecontrol = vol.dwValue
       volumeout = vol.dwValue
      
       Exit Function
   End If

   getoutvolumecontrol = False

End Function

Function setmuteon(ByVal hmixer As Long, mxc As MIXERCONTROL) As Boolean
'This function sets the value for a volume control. Returns True if successful

   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol() As MIXERCONTROLDETAILS_UNSIGNED
   Dim u As Long

   mxcd.item = mxc.cMultipleItems
   
   ReDim vol(mxcd.item - 1) As MIXERCONTROLDETAILS_UNSIGNED
   
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(mxcd.item - 1))
   
   ' Allocate a buffer for the control value buffer
   hmem = GlobalAlloc(&H40, Len(vol(mxcd.item - 1)) * mxcd.item)
   mxcd.paDetails = GlobalLock(hmem)
   mxcd.cChannels = 1
  
   u = mxcd.item
   
   ' Copy the data into the control value buffer
   Do
   u = u - 1
   vol(u).dwValue = 1
   CopyPtrFromStruct mxcd.paDetails, vol(u).dwValue, Len(vol(mxcd.item - 1)) * mxcd.item
   Loop Until u < 1
   
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   
   GlobalFree (hmem)
  
   If (MMSYSERR_NOERROR = rc) Then
       setmuteon = True
   Else
       setmuteon = False
   End If
   
End Function
