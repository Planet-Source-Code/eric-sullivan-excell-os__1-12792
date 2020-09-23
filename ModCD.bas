Attribute VB_Name = "ModCD"
Option Explicit
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public hmem As Long

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_PURGE = &H40
Public Const SND_FILENAME = &H20000

Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Public Type MIXERCONTROLDETAILS
    cbStruct    As Long
    dwControlID As Long
    cChannels   As Long
    item        As Long
    cbDetails   As Long
    paDetails   As Long
End Type

Public Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type

Public Type MIXERCONTROL
    cbStruct       As Long
    dwControlID    As Long
    dwControlType  As Long
    fdwControl     As Long
    cMultipleItems As Long
    szShortName    As String * MIXER_SHORT_NAME_CHARS
    szName         As String * MIXER_LONG_NAME_CHARS
    lMinimum       As Long
    lMaximum       As Long
    reserved(10)   As Long
End Type

Public Type MIXERLINECONTROLS
    cbStruct  As Long
    dwLineID  As Long
    dwControl As Long
    cControls As Long
    cbmxctrl  As Long
    pamxctrl  As Long
End Type

Public Type MIXERLINE
    cbStruct        As Long
    dwDestination   As Long
    dwSource        As Long
    dwLineID        As Long
    fdwLine         As Long
    dwUser          As Long
    dwComponentType As Long
    cChannels       As Long
    cConnections    As Long
    cControls       As Long
    szShortName     As String * MIXER_SHORT_NAME_CHARS
    szName          As String * MIXER_LONG_NAME_CHARS
    dwType          As Long
    dwDeviceID      As Long
    wMid            As Integer
    wPid            As Integer
    vDriverVersion  As Long
    szPname         As String * MAXPNAMELEN
End Type

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Function fGetVolumeControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl  As MIXERLINE
    Dim hmem As Long
    Dim rc   As Long
    
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
    rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    If MMSYSERR_NOERROR = rc Then
        With mxlc
            .cbStruct = Len(mxlc)
            .dwLineID = mxl.dwLineID
            .dwControl = ctrlType
            .cControls = 1
            .cbmxctrl = Len(mxc)
        End With

        hmem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hmem)
        mxc.cbStruct = Len(mxc)

        rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If MMSYSERR_NOERROR = rc Then
            fGetVolumeControl = True
            Call CopyStructFromPtr(mxc, mxlc.pamxctrl, Len(mxc))
        Else
            fGetVolumeControl = False
        End If
        
        Call GlobalFree(hmem)
        Exit Function
    End If
    fGetVolumeControl = False
End Function

