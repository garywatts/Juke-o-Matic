Attribute VB_Name = "Module1"

'Call the OpenMixer function (usually mixer 0).
'Then you can get/set all the primary mixer controls with the SetVolume, GetVolume, SetMute & GetMute functions.
'Example:
'
'OpenMixer(0)
'Slider1.Value=GetVolume(WAVEOUT)
'
Option Explicit

' ***************************************************************************
'                        Mixer API Support (winmm.dll)
' ***************************************************************************
Private Const MMSYSERR_NOERROR = 0
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MAXPNAMELEN = 32                  ' max product name length (including NULL)
Private Const MAXERRORLENGTH = 128              ' max error text length (including final NULL)
Private Const MM_MIXM_LINE_CHANGE = &H3D0       ' mixer line change notify
Private Const MM_MIXM_CONTROL_CHANGE = &H3D1    ' mixer control change notify
'
'   MMRESULT error return values specific to the mixer API
'
Private Const MIXERR_BASE = 1024
Private Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Private Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Private Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Private Const MIXERR_LASTERROR = (MIXERR_BASE + 2)

Private Const MIXER_OBJECTF_HANDLE = &H80000000
Private Const MIXER_OBJECTF_MIXER = &H0&
Private Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Private Const MIXER_OBJECTF_WAVEOUT = &H10000000
Private Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Private Const MIXER_OBJECTF_WAVEIN = &H20000000
Private Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Private Const MIXER_OBJECTF_MIDIOUT = &H30000000
Private Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Private Const MIXER_OBJECTF_MIDIIN = &H40000000
Private Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Private Const MIXER_OBJECTF_AUX = &H50000000
'
'   MIXERLINE.fdwLine
'
Private Const MIXERLINE_LINEF_ACTIVE = &H1&
Private Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Private Const MIXERLINE_LINEF_SOURCE = &H80000000
'
'   MIXERLINE.dwComponentType
'
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Private Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Private Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)

Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
'
'   MIXERLINE.Target.dwType
'
Private Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Private Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Private Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Private Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Private Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Private Const MIXERLINE_TARGETTYPE_AUX = 5

Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
'
'   MIXERCONTROL.fdwControl
'
Private Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Private Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Private Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
'
'   MIXERCONTROL_CONTROLTYPE_xxx building block defines
'
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
'
'   Commonly used control types for specifying MIXERCONTROL.dwControlType
'
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)


Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&

Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&


' mixer API types

Private Type Target    ' for use in MIXERLINE and others (embedded structure)
    dwType          As Long                 '  MIXERLINE_TARGETPrivate Type_xxxx
    dwDeviceID      As Long                 '  target device ID of device Private Type
    wMid            As Integer              '  of target device
    wPid            As Integer              '       "
    vDriverVersion  As Long                 '       "
    szPname         As String * MAXPNAMELEN
End Type

Private Type MIXERLINE
    cbStruct        As Long                 '  size of MIXERLINE structure
    dwDestination   As Long                 '  zero based destination index
    dwSource        As Long                 '  zero based source index (if source)
    dwLineID        As Long                 '  unique line id for mixer device
    fdwLine         As Long                 '  state/information about line
    dwUser          As Long                 '  driver specific information
    dwComponentType As Long                 '  component Private Type line connects to
    cChannels       As Long                 '  number of channels line supports
    cConnections    As Long                 '  number of connections (possible)
    cControls       As Long                 '  number of controls at this line
    szShortName     As String * MIXER_SHORT_NAME_CHARS
    szName          As String * MIXER_LONG_NAME_CHARS
    tTarget         As Target
End Type

Private Type MIXERCONTROL
    cbStruct        As Long                 '  size in Byte of MIXERCONTROL
    dwControlID     As Long                 '  unique control id for mixer device
    dwControlType   As Long                 '  MIXERCONTROL_CONTROLPrivate Type_xxx
    fdwControl      As Long                 '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems  As Long                 '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
    szName(1 To MIXER_LONG_NAME_CHARS) As Byte
    'Bounds(1 To 6)  As Long                 '  Longest member of the Bounds union
    'Metrics(1 To 6) As Long                 '  Longest member of the Metrics union
    ' alternate defs for two items above
    lMinimum As Long                        '  Minimum value
    lMaximum As Long                        '  Maximum value
    RESERVED(10) As Long                    '  reserved structure space
End Type

Private Type MIXERLINECONTROLS
        cbStruct        As Long             '  size in Byte of MIXERLINECONTROLS
        dwLineID        As Long             '  line id (from MIXERLINE.dwLineID)
                                            '  MIXER_GETLINECONTROLSF_ONEBYID or
        dwControl       As Long             '  MIXER_GETLINECONTROLSF_ONEBYPrivate Type
        cControls       As Long             '  count of controls pmxctrl points to
        cbmxctrl        As Long             '  size in Byte of _one_ MIXERCONTROL
        'pamxctrl        As MIXERCONTROL     '  pointer to first MIXERCONTROL array
        pamxctrl        As Long             '  pointer to first MIXERCONTROL array
End Type

Private Type MIXERCONTROLDETAILS
        cbStruct        As Long             '  size in Byte of MIXERCONTROLDETAILS
        dwControlID     As Long             '  control id to get/set details on
        cChannels       As Long             '  number of channels in paDetails array
        item            As Long             '  hwndOwner or cMultipleItems
        cbDetails       As Long             '  size of _one_ details_XX struct
        paDetails       As Long             '  pointer to array of details_XX structs
End Type

Private Type MIXERCONTROLDETAILS_UNSIGNED
        dwValue As Long
End Type

' mixer API prototypes
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

' misc API prototypes
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private uMixerControls(20)      As MIXERCONTROL ' local array to store mixer controls
Private hMixerHandle            As Long         ' handle for mixer

' enums to identify mixer volume controls
Public Enum VOL_CONTROL
    SPEAKER = 0
    LINEIN = 1
    MICROPHONE = 2
    SYNTHESIZER = 3
    COMPACTDISC = 4
    WAVEOUT = 5
    AUXILIARY = 6
End Enum

' enums to identify mixer mute controls
Public Enum MUTE_CONTROL
    SPEAKER_MUTE = 7
    LINEIN_MUTE = 8
    MICROPHONE_MUTE = 9
    SYNTHESIZER_MUTE = 10
    COMPACTDISC_MUTE = 11
    WAVEOUT_MUTE = 12
    AUXILIARY_MUTE = 13
End Enum

Public Function OpenMixer(ByVal MixerNumber As Long) As Long
Dim ret As Long

' is there a mixer available?
If MixerNumber < 0 Or MixerNumber > mixerGetNumDevs - 1 Then Exit Function

' open the mixer
ret = mixerOpen(hMixerHandle, MixerNumber, 0, 0, 0)
If ret <> MMSYSERR_NOERROR Then Exit Function

' get the primary line controls by name, (this does not get all of the controls).

' speaker (master) volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(SPEAKER))
' microphone volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(MICROPHONE))
' Line volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(AUXILIARY))
' CD volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(COMPACTDISC))
' Synthesizer volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(SYNTHESIZER))
' wave volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(WAVEOUT))
' Aux volume
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_LINE, MIXERCONTROL_CONTROLTYPE_VOLUME, uMixerControls(LINEIN))

' speaker (master) mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(SPEAKER_MUTE))
' microphone mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(MICROPHONE_MUTE))
' Line mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(AUXILIARY_MUTE))
' CD mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(COMPACTDISC_MUTE))
' Synthesizer mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(SYNTHESIZER_MUTE))
' wave mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(WAVEOUT_MUTE))
' Aux mute
ret = GetMixerControl(hMixerHandle, MIXERLINE_COMPONENTTYPE_SRC_LINE, MIXERCONTROL_CONTROLTYPE_MUTE, uMixerControls(LINEIN_MUTE))

' return the mixer handle
OpenMixer = True

End Function

Public Function CloseMixer() As Long
CloseMixer = mixerClose(hMixerHandle)
hMixerHandle = 0
    
End Function

Public Function SetVolume(Control As VOL_CONTROL, ByVal NewVolume As Long) As Long
SetVolume = SetControlValue(hMixerHandle, uMixerControls(Control), NewVolume)

End Function

Public Function GetVolume(Control As VOL_CONTROL) As Long
GetVolume = GetControlValue(hMixerHandle, uMixerControls(Control))

End Function

Public Function SetMute(Control As MUTE_CONTROL, ByVal MuteState As Boolean) As Boolean
Dim Mute As Long

Mute = Abs(MuteState)
SetMute = SetControlValue(hMixerHandle, uMixerControls(Control), Mute)
    
End Function

Public Function GetMute(Control As MUTE_CONTROL) As Boolean
GetMute = CBool(-GetControlValue(hMixerHandle, uMixerControls(Control)))

End Function

Private Function GetMixerControl(ByVal hMixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Long
' This function attempts to obtain a mixer control. Returns True if successful.
Dim mxlc As MIXERLINECONTROLS
Dim mxl As MIXERLINE
Dim hMem As Long
Dim ret As Long
         
mxl.cbStruct = Len(mxl)
mxl.dwComponentType = componentType

' Obtain a line corresponding to the component type
ret = mixerGetLineInfo(hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
     
If ret = MMSYSERR_NOERROR Then
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.dwControl = ctrlType
    mxlc.cControls = 1
    mxlc.cbmxctrl = Len(mxc)
         
    ' Allocate a buffer for the control
    hMem = GlobalAlloc(&H40, Len(mxc))
    mxlc.pamxctrl = GlobalLock(hMem)
    mxc.cbStruct = Len(mxc)
         
    ' Get the control
    ret = mixerGetLineControls(hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
              
    If ret = MMSYSERR_NOERROR Then
        GetMixerControl = True
             
        ' Copy the control into the destination structure
        CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
    Else
        GetMixerControl = False
    End If
    GlobalFree (hMem)
    Exit Function
End If
  
GetMixerControl = False

End Function

Private Function SetControlValue(ByVal hMixer As Long, mxc As MIXERCONTROL, ByVal NewVolume As Long) As Boolean
'This function sets the value for a control. Returns True if successful
Dim mxcd As MIXERCONTROLDETAILS
Dim vol As MIXERCONTROLDETAILS_UNSIGNED
Dim hMem As Long
Dim ret As Long

mxcd.item = 0
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol)

' Allocate a buffer for the control value buffer
hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
mxcd.cChannels = 1

' setup value, use percent of range if max is greater than 100
If mxc.lMaximum > 100 Then
    vol.dwValue = NewVolume * (mxc.lMaximum \ 100)
Else
    vol.dwValue = NewVolume
End If
If vol.dwValue > mxc.lMaximum Then vol.dwValue = mxc.lMaximum
If vol.dwValue < mxc.lMinimum Then vol.dwValue = mxc.lMinimum

' Copy the data into the control value buffer
CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
     
' Set the control value
ret = mixerSetControlDetails(hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
GlobalFree (hMem)

If ret = MMSYSERR_NOERROR Then SetControlValue = True

End Function

Private Function GetControlValue(ByVal hMixer As Long, mxc As MIXERCONTROL) As Long
'This function gets the value for a control.
Dim mxcd As MIXERCONTROLDETAILS
Dim vol As MIXERCONTROLDETAILS_UNSIGNED
Dim hMem As Long
Dim ret As Long

mxcd.item = 0
mxcd.dwControlID = mxc.dwControlID
mxcd.cbStruct = Len(mxcd)
mxcd.cbDetails = Len(vol)

hMem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hMem)
mxcd.cChannels = 1

' Get the control value
ret = mixerGetControlDetails(hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)

' Copy the data into the control value buffer
CopyStructFromPtr vol, mxcd.paDetails, Len(vol)

If mxc.lMaximum > 100 Then
    GetControlValue = (vol.dwValue * 100) / mxc.lMaximum - mxc.lMinimum
Else
    GetControlValue = vol.dwValue
End If

GlobalFree (hMem)

End Function
