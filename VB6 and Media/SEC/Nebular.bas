Attribute VB_Name = "Nebular"
Private m_dx As New DirectX7
Private m_dxs As DirectSound 'Then there is the sub object, DirectSound:
Declare Function agGetAddressForObject Lib "apigid32.dll" (object As Any) As Long
Type dxBuffers
  isLoaded As Boolean
  Buffer As DirectSoundBuffer
End Type
Private SoundFolder As String 'Holds Path to Sound folder
Private SB() As dxBuffers 'An Array of BUFFERS,
Private CurrentBuffer As Integer 'Holds last assign Random Buffer Number

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
' ---------
' Mixer errors
' ---------
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13)
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)
Public Const MMSYSERR_NOERROR = 0
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)

Public Const MIDIERR_BASE = 64
Public Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)
Public Const HIGHEST_VOLUME_SETTING = 65535
Public Const AUX_MAPPER = -1&
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2& ' separate left-right volume control
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&

' Mixer control get/set details
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&

Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&

' Mixer get line controls
Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

' Mixer get line info of
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&

'
Public Const MIXER_OBJECTF_AUX = &H50000000
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000

Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)

' Control flag
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&

' Mixer control classes

Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)

Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)

Public Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)

Public Const MIXERLINE_COMPONENTTYPE_SRC_TADVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)

Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                             (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
               
Public Const MIXERLINE_COMPONENTTYPE_src_AUXVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)

Public Const MIXERLINE_COMPONENTTYPE_SRC_PSPKVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)

Public Const MIXERLINE_COMPONENTTYPE_SRC_MBOOST = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)

Public Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)

Public Const CALLBACK_FUNCTION = &H30000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MM_WOM_DONE = &H3BD
Public Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Public Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks
Public Const AUXCAPS_VOLUME = &H1   '  supports volume control
Public Const AUXCAPS_LRVOLUME = &H2 '  separate left-right volume control
' Mixer control types
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)

Public Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" _
(ByVal DllPath As String, bRegister As Boolean) As Boolean

Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Declare Function mixerClose Lib "winmm.dll" _
               (ByVal hmx As Long) As Long
   
Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
   
Declare Function mixerGetDevCaps Lib "winmm.dll" _
               Alias "mixerGetDevCapsA" _
               (ByVal uMxId As Long, _
               ByVal pmxcaps As MIXERCAPS, _
               ByVal cbmxcaps As Long) As Long
   
Declare Function mixerGetID Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pumxID As Long, _
               ByVal fdwId As Long) As Long
               
Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long

Declare Function mixerMessage Lib "winmm.dll" _
               (ByVal hmx As Long, _
               ByVal uMsg As Long, _
               ByVal dwParam1 As Long, _
               ByVal dwParam2 As Long) As Long
               
Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Declare Function GlobalLock Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function GlobalFree Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function waveOutOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As waveFormat, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveOutAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal x As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As waveFormat, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long


Dim rc As Long
Dim msg As String * 200

' variables for managing wave file
Public formatA As waveFormat
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim hWaveOut As Long
Dim bufferIn As Long
Dim hmem As Long
Dim outHdr As WAVEHDR
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean
Public fPlaying As Boolean
               
               
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Type VolumeSetting
    leftVol As Integer
    rightVol As Integer
End Type

Type AUXCAPS
       wMid As Integer
       wPid As Integer
       vDriverVersion As Long
       szPname As String * MAXPNAMELEN
       wTechnology As Integer
       dwSupport As Long
End Type

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
   Reserved(10) As Long       '  reserved structure space
   End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   Item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long        '  value of the control (volume level)
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

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
   End Type
   Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
   End Type
   Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
    End Type

Function GetMixerControl(ByVal hmixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        
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
       mxlc.cControls = 1
       mxlc.cbmxctrl = Len(mxc)
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       ' Get the control
       rc = mixerGetLineControls(hmixer, _
                                 mxlc, _
                                 MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function

Function SetVolumeControl(ByVal hmixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal volume As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
   
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(Vol)
   ' Allocate memory for the control value buffer
   ' Len(Vol) = Number of Bytes
   ' &H40     = GMEM_ZEROINIT = Initialize to zero
   ' hmem     = Handle to the newly created memory object (Buffer)
   hmem = GlobalAlloc(&H40, Len(Vol))
   ' Lock the memory object (buffer) and return a pointer to the first byte.
   mxcd.paDetails = GlobalLock(hmem)
   Vol.dwValue = volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, _
                              mxcd, _
                              MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function

Function unSetMuteControl(ByVal hmixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal unmute As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
   
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(Vol)
   ' Allocate memory for the control value buffer
   ' Len(Vol) = Number of Bytes
   ' &H40     = GMEM_ZEROINIT = Initialize to zero
   ' hmem     = Handle to the newly created memory object (Buffer)
   hmem = GlobalAlloc(&H40, Len(Vol))
   ' Lock the memory object (buffer) and return a pointer to the first byte.
   mxcd.paDetails = GlobalLock(hmem)
   Vol.dwValue = unmute
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       unSetMuteControl = True
   Else
       unSetMuteControl = False
   End If
End Function


Function SetMuteControl(ByVal hmixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal mute As Boolean) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
   
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(Vol)
   ' Allocate memory for the control value buffer
   ' Len(Vol) = Number of Bytes
   ' &H40     = GMEM_ZEROINIT = Initialize to zero
   ' hmem     = Handle to the newly created memory object (Buffer)
   hmem = GlobalAlloc(&H40, Len(Vol))
   ' Lock the memory object (buffer) and return a pointer to the first byte.
   mxcd.paDetails = GlobalLock(hmem)
   Vol.dwValue = volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetMuteControl = True
   Else
       SetMuteControl = False
   End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED

    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.Item = 0
    mxcd.cbDetails = Len(Vol)
    mxcd.paDetails = 0
    ' Allocate memory for the control value buffer
    ' Len(Vol) = Number of Bytes
    ' &H40     = GMEM_ZEROINIT = Initialize to zero
    ' hmem     = Handle to the newly created memory object (Buffer)
    hmem = GlobalAlloc(&H40, Len(Vol))
    ' Lock the memory object (buffer) and return a pointer to the first byte.
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    ' Copy the data from the control value buffer
    CopyStructFromPtr Vol, mxcd.paDetails, Len(Vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = Vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function

Function GetControlType(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

Select Case vValue

   Case (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_FADER"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_BASS"

   Case (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BOOLEAN"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BOOLEANMETER"

   Case (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BUTTON"

   Case (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
      sName = "MIXERCONTROL_CONTROLTYPE_CUSTOM"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
      sName = "MIXERCONTROL_CONTROLTYPE_DECIBELS"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 4)
      sName = "MIXERCONTROL_CONTROLTYPE_EQUALIZER"

   Case (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
      sName = "MIXERCONTROL_CONTROLTYPE_LOUDNESS"

   Case (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_MICROTIME"

   Case (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_MILLITIME"

   Case (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_MIXER"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
      sName = "MIXERCONTROL_CONTROLTYPE_MONO"

   Case (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SLIDER"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
      sName = "MIXERCONTROL_CONTROLTYPE_STEREOENH"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 3)
      sName = "MIXERCONTROL_CONTROLTYPE_TREBLE"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_UNSIGNED"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_VOLUME"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SIGNEDMETER"

   Case (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_SINGLESELECT"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_MUTE"

   Case (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_MUX"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_ONOFF"

   Case (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_PAN"

   Case (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_PEAKMETER"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
      sName = "MIXERCONTROL_CONTROLTYPE_PERCENT"

   Case (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_QSOUNDPAN"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SIGNED"

   Case Else
      sName = "<invalid>"

End Select

GetControlType = sName

End Function


Sub lCrossFader()
Vol1 = 100 - sldPan.Value ' Left
Vol2 = 100 - sldPan.Value ' Right
E = CrossFader.Value
F = 100 - E
If Check4.Value = 1 Then ' Half Fader Check
    LVol = (F * Val(Vol1) / 100) * 2
    RVOL = (E * Val(Vol2) / 100) * 2
    If LVol > (50 * Val(Vol1) / 100) * 2 Then
        LVol = (50 * Val(Vol1) / 100) * 2
    End If
    If RVOL > (50 * Val(Vol2) / 100) * 2 Then
        RVOL = (50 * Val(Vol2) / 100) * 2
    End If
Else
    LVol = (F * Val(Vol1) / 100)
    RVOL = (E * Val(Vol2) / 100)
End If
Label1.Caption = "Fader: " + LTrim$(Str$(LVol)) + " x " + LTrim$(Str$(RVOL))

End Sub


Public Function lSetVolume(ByRef lLeftVol As Long, ByRef lrightVol As Long, lDeviceID As Long) As Long

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API passed value of the Combined Volumes
    
    volume.leftVol = nSigned(lLeftVol * 65535 / HIGHEST_VOLUME_SETTING)
    volume.rightVol = nSigned(lrightVol * 65535 / HIGHEST_VOLUME_SETTING)
    
    lDataLen = Len(volume)
    CopyMemory lBothVolumes, volume.leftVol, lDataLen

    lAPIReturnVal = auxSetVolume(lDeviceID, lBothVolumes)
    lSetVolume = lAPIReturnVal

End Function


Public Function lGetVolume(ByRef lLeftVol As Long, ByRef lrightVol As Long, lDeviceID As Long) As Long

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API Return of the Combined Volumes
    lAPIReturnVal = auxGetVolume(lDeviceID, lBothVolumes)
    lDataLen = Len(volume)
    CopyMemory volume.leftVol, lBothVolumes, lDataLen
    lLeftVol = HIGHEST_VOLUME_SETTING * lUnsigned(volume.leftVol) / 65535
    lrightVol = HIGHEST_VOLUME_SETTING * lUnsigned(volume.rightVol) / 65535
    lGetVolume = lAPIReturnVal
End Function

Public Function nSigned(ByVal lUnsignedInt As Long) As Integer
    Dim nReturnVal As Integer                          ' Return value from Function
    
    If lUnsignedInt > 65535 Or lUnsignedInt < 0 Then
        MsgBox "Error in conversion from Unsigned to nSigned Integer"
        nSignedInt = 0
        Exit Function
    End If

    If lUnsignedInt > 32767 Then
        nReturnVal = lUnsignedInt - 65536
    Else
        nReturnVal = lUnsignedInt
    End If
    
    nSigned = nReturnVal

End Function

Public Function lUnsigned(ByVal nSignedInt As Integer) As Long
    Dim lReturnVal As Long                          ' Return value from Function
    
    If nSignedInt < 0 Then
        lReturnVal = nSignedInt + 65536
    Else
        lReturnVal = nSignedInt
    End If
    
    If lReturnVal > 65535 Or lReturnVal < 0 Then
        MsgBox "Error in conversion from nSigned to Unsigned Integer"
        lReturnVal = 0
    End If
    
    lUnsigned = lReturnVal
End Function


Sub waveOutProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
' Wave IO Callback function
   If (uMsg = MM_WOM_DONE) Then
      fPlaying = False
   End If
End Sub

Sub CloseWaveOut()
' Close the waveout device
    rc = waveOutReset(hWaveOut)
    rc = waveOutUnprepareHeader(hWaveOut, outHdr, Len(outHdr))
    rc = waveOutClose(hWaveOut)
End Sub

Sub LoadFile(inFile As String)
' Load wavefile into memory
   Dim hmmioIn As Long
   Dim mmioinf As mmioinfo
   fFileLoaded = False
   If (inFile = "") Then
       GlobalFree (hmem)
       Exit Sub
   End If
   ' Open the input file
   hmmioIn = mmioOpen(inFile, mmioinf, MMIO_READ)
   If hmmioIn = 0 Then
       MsgBox "Error opening input file, rc = " & mmioinf.wErrorRet
       Exit Sub
   End If
   
   ' Check if this is a wave file
   mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
   rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Not a wave file"
       Exit Sub
   End If
   
   ' Get format info
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Couldn't get format chunk"
       Exit Sub
   End If
   rc = mmioReadFormat(hmmioIn, formatA, Len(formatA))
   If (rc = -1) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Error reading format"
      Exit Sub
   End If
   rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
   
   ' Find the data subchunk
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Couldn't get data chunk"
      Exit Sub
   End If
   
   ' Allocate soundbuffer and read sound data
   GlobalFree hmem
   hmem = GlobalAlloc(&H40, mmckinfoSubchunkIn.ckSize)
   bufferIn = GlobalLock(hmem)
   rc = mmioRead(hmmioIn, bufferIn, mmckinfoSubchunkIn.ckSize)
   
   numSamples = mmckinfoSubchunkIn.ckSize / formatA.nBlockAlign
   
   ' Close file
   rc = mmioClose(hmmioOut, 0)
   
   fFileLoaded = True
    
End Sub

Sub play(ByVal soundcard As Integer)
' Send audio buffer to wave output

    rc = waveOutOpen(hWaveOut, soundcard, formatA, AddressOf waveOutProc, 0, CALLBACK_FUNCTION)
    If (rc <> 0) Then
      GlobalFree (hmem)
      waveOutGetErrorText rc, msg, Len(msg)
      MsgBox msg
      Exit Sub
    End If

    outHdr.lpData = bufferIn + (drawFrom * formatA.nBlockAlign)
    outHdr.dwBufferLength = (drawTo - drawFrom) * formatA.nBlockAlign
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0

    rc = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      waveOutGetErrorText rc, msg, Len(msg)
      MsgBox msg
    End If

    rc = waveOutWrite(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      GlobalFree (hmem)
    Else
      fPlaying = True
    End If
End Sub

Sub StopPlay()
   waveOutReset (hWaveOut)
End Sub


Sub GetStereo16Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)
' These subs obtain a PCM sample and converts it into volume levels from (-1 to 1)
   Dim sample16 As Integer
   Dim ptr As Long
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768
   CopyStructFromPtr sample16, ptr + 2, 2
   rightVol = sample16 / 32768

End Sub

Sub GetStereo8Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128
   CopyStructFromPtr sample8, ptr + 1, 1
   rightVol = (sample8 - 128) / 128

End Sub

Sub GetMono16Sample(ByVal sample As Long, ByRef leftVol As Double)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768

End Sub

Sub GetMono8Sample(ByVal sample As Long, ByRef leftVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128

End Sub

Public Function CheckFileVersion(FilenameAndPath As Variant) As Variant
    On Error GoTo HandelCheckFileVersionError
    Dim lDummy As Long, lsize As Long, rc As Long
    Dim lVerbufferLen As Long, lVerPointer As Long
    Dim sBuffer() As Byte
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    Dim ProdVer As String
    lsize = GetFileVersionInfoSize(FilenameAndPath, lDummy)
    If lsize < 1 Then Exit Function
    ReDim sBuffer(lsize)
    rc = GetFileVersionInfo(FilenameAndPath, 0&, lsize, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    '**** Determine Product Version number *
    '     ***
    ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl)
    CheckFileVersion = ProdVer
    Exit Function
HandelCheckFileVersionError:
    CheckFileVersion = "N/A"
    Exit Function
    MsgBox check
End Function



Public Sub SoundDir(FolderPath As String)
  SoundFolder = FolderPath & "\"
End Sub

Public Sub CreateBuffers(AmountOfBuffer As Integer, DefaultFile As String)
  ReDim SB(AmountOfBuffer)
  For AmountOfBuffer = 0 To AmountOfBuffer
    DX7LoadSound AmountOfBuffer, DefaultFile 'must assign a defualt sound
    VolumeLevel AmountOfBuffer, 75 ' set volume to 50% for default
  Next AmountOfBuffer
End Sub

Public Sub SetupDX7Sound(CurrentForm As Form)
  Set m_dxs = m_dx.DirectSoundCreate("") 'create a DSound object
 'Next you check for any errors, if there are no errors the user has got DX7 and a functional sound card

  If err.Number <> 0 Then
    MsgBox "Unable to start DirectSound. Check to see that your sound card is properly installed"
    End
  End If
  m_dxs.SetCooperativeLevel CurrentForm.hwnd, DSSCL_PRIORITY 'THIS MUST BE SET BEFORE WE CREATE ANY BUFFERS
  
  'associating our DS object with our window is important. This tells windows to stop
  'other sounds from interfering with ours, and ours not to interfere with other apps.
  'The sounds will only be played when the from has got focus.
  'DSSCL_PRIORITY=no cooperation, exclusive access to the sound card, Needed for games
  'DSSCL_NORMAL=cooperates with other apps, shares resources, Good for general windows multimedia apps.
  
End Sub

Public Sub DX7LoadSound(Buffer As Integer, sfile As String)
  Dim filename As String
  Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
  Dim waveFormat As WAVEFORMATEX 'what sort of buffer to create
  
  bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN _
  Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC 'These settings should do for almost any app....
  
  waveFormat.nFormatTag = WAVE_FORMAT_PCM
  waveFormat.nChannels = 2    '2 channels
  waveFormat.lSamplesPerSec = 22050
  waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
  waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
  waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign

  filename = SoundFolder & sfile
  On Error GoTo Continue
  Set SB(Buffer).Buffer = m_dxs.CreateSoundBufferFromFile(filename, bufferDesc, waveFormat)
  SB(Buffer).isLoaded = True
  Exit Sub
Continue:
  MsgBox "Error can't find file: " & filename
End Sub

Public Function PlaySoundAnyBuffer(filename As String, Optional volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte) As Integer
  
  Do While SB(CurrentBuffer).Buffer.GetStatus = DSBSTATUS_PLAYING 'Find an empty buffer
    CurrentBuffer = CurrentBuffer + 1
    If CurrentBuffer > UBound(SB) Then CurrentBuffer = 0
  Loop

  DX7LoadSound CurrentBuffer, filename
  If PanValue <> 50 Then PanSound CurrentBuffer, PanValue
  If volume < 100 Then VolumeLevel CurrentBuffer, volume
  If SB(CurrentBuffer).isLoaded Then SB(CurrentBuffer).Buffer.play LoopIt 'dsb_looping=1, dsb_default=0
End Function

Public Sub PlaySoundWithPan(Buffer As Integer, filename As String, Optional volume As Byte, Optional PanValue As Byte, Optional LoopIt As Byte)
  DX7LoadSound Buffer, filename
  If PanValue <> 50 And PanValue < 100 Then PanSound Buffer, PanValue
  If volume < 100 Then VolumeLevel Buffer, volume
  If SB(Buffer).isLoaded Then SB(Buffer).Buffer.play LoopIt 'dsb_looping=1, dsb_default=0
End Sub

Public Sub PanSound(Buffer As Integer, PanValue As Byte)
  Select Case PanValue
    Case 0
      SB(Buffer).Buffer.SetPan -10000
    Case 100
      SB(Buffer).Buffer.SetPan 10000
    Case Else
      SB(Buffer).Buffer.SetPan (100 * PanValue) - 5000
  End Select
End Sub

Public Sub VolumeLevel(Buffer As Integer, volume As Byte)
  If volume > 0 Then ' stop division by 0
    SB(Buffer).Buffer.SetVolume (60 * volume) - 6000
  Else
    SB(Buffer).Buffer.SetVolume -6000
  End If
End Sub

Public Function IsPlaying(Buffer As Integer) As Long
  IsPlaying = SB(Buffer).Buffer.GetStatus
End Function


