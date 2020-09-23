Attribute VB_Name = "SoundBas"

Option Explicit

Public hmixer As Long
Public outputVolCtrl As MIXERCONTROL
Public rc As Long
Public OK As Boolean
Public mxcd As MIXERCONTROLDETAILS
Public volume As Long
Public volHmem As Long


Public Type VULights
    VUOn As Boolean
    InOutLev As Double
    VolLev As Double
    VULev As Double
    VolUnit As Variant
    VUArray As Long
    Freq(0 To 9) As Double
    FreqNum As Integer
    FreqVal As Double
End Type

Public Const GMEM_FIXED = &H0

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

Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
'
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&

Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&

Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)

Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
           (ByVal hmxobj As Long, _
            pmxcd As MIXERCONTROLDETAILS, _
            ByVal fdwDetails As Long) As Long

Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" _
                  (ByVal hmxobj As Long, _
                  pmxlc As MIXERLINECONTROLS, _
                  ByVal fdwControls As Long) As Long
'
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" _
                    (ByVal hmxobj As Long, _
                     pmxl As MIXERLINE, _
                     ByVal fdwInfo As Long) As Long

Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, _
                                             ByVal uMxId As Long, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal fdwOpen As Long) As Long


Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                             ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal HMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal HMem As Long) As Long

Type MIXERCONTROL
   cbStruct As Long
   dwControlID As Long
   dwControlType As Long
   fdwControl As Long
   cMultipleItems As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   lMinimum As Long
   lMaximum As Long
   Reserved(10) As Long
End Type
'
Type MIXERCONTROLDETAILS
   cbStruct As Long
   dwControlID As Long
   cChannels As Long
   item As Long
   cbDetails As Long
   paDetails As Long
End Type

Type MIXERLINE
   cbStruct As Long
   dwDestination As Long
   dwSource As Long
   dwLineID As Long
   fdwLine As Long
   dwUser As Long
   dwComponentType As Long
   cChannels As Long

   cConnections As Long

   cControls As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS


   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type
'
Type MIXERLINECONTROLS
   cbStruct As Long
   dwLineID As Long
   dwControl As Long
   cControls As Long
   cbmxctrl As Long
   pamxctrl As Long
End Type

'Public hmem(NUM_BUFFERS) As Long
Public Const DEVICEID = 0


Function GetControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim HMem As Long
   Dim rc As Long

   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)

   If (MMSYSERR_NOERROR = rc) Then
      mxlc.cbStruct = Len(mxlc)
      mxlc.dwLineID = mxl.dwLineID
      mxlc.dwControl = ctrlType
      mxlc.cControls = 1
      mxlc.cbmxctrl = Len(mxc)
      mxlc.pamxctrl = 9

      HMem = GlobalAlloc(GMEM_FIXED, Len(mxc))
      mxlc.pamxctrl = GlobalLock(HMem)
      mxc.cbStruct = Len(mxc)

      rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)

      If (MMSYSERR_NOERROR = rc) Then
         GetControl = True

         CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
      Else
         GetControl = False
      End If
      GlobalFree (HMem)
      Exit Function
   End If

   GetControl = False
End Function
