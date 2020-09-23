Attribute VB_Name = "Module1"
'All declarations, type definitions and constants
'are from WIN32API.txt coming with VB 5.0 Proffessional

'But, the callback function created by myself,
'after wasting hours on the documents related with
'C++ Windows midi programming.
'The program is not running from VB
'on my computer. So I had to make a shortcut
'to it's EXE, and everytime compiled it to the
'same exe file and run it from desktop.
'I don't know why, but I advise you to do the same.

'Function declarations for midi out
Declare Function midiOutGetNumDevs Lib "winmm" () As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long

'Function declarations for midi in
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Any, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long

Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)

'Type declaration for getting the device capabilities.
'Use with midiInGetDevCaps, midiOutGetdevCaps
Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
End Type
Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type

Public oxMIC As MIDIINCAPS
Public oxMOC As MIDIOUTCAPS
Public tmp As Long, tmp1 As Long, tmp2 As Long

'These are the return values for the second argument
'of our callback function.
Public Const MM_MIM_OPEN = &H3C1  '  Midi in device opened by MidiInOpen
Public Const MM_MIM_CLOSE = &H3C2  '  Midi in device closed by MidiInOpen
Public Const MM_MIM_DATA = &H3C3  '  A non sys-ex midi msg (Note on, note off, pitch ch., bank ch. volume ch. etc..)
Public Const MM_MIM_LONGDATA = &H3C4  'Sys-ex data
Public Const MM_MIM_ERROR = &H3C5  'An error occured inputting non sys-ex data
Public Const MM_MIM_LONGERROR = &H3C6  'An error occured inputting sys-ex data
Public Const MIDIERR_BASE = 64
'These are the callback types.
'We are using the CALLBACK_FUNCTION constant
'Windows calls our callback function each time
'the opened midi in port inputs data
Public Const CALLBACK_TYPEMASK = &H70000      '  callback type mask
Public Const CALLBACK_NULL = &H0        '  no callback
Public Const CALLBACK_WINDOW = &H10000      '  dwCallback is a HWND
Public Const CALLBACK_TASK = &H20000      '  dwCallback is a HTASK
Public Const CALLBACK_FUNCTION = &H30000      '  dwCallback is a FARPROC

' Check these, if a function has returned an error:
' General error return values
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_NOERROR = 0  '  no error
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)  '  unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '  device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)  '  driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)  '  device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '  device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)  '  memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)  '  function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)  '  error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10) '  invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11) '  invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12) '  handle being used simultaneously on another thread (eg callback)
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) '  "Specified alias not found in WIN.INI
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13) '  last error in range


'This is the callback function which windows calls during the recording.
'You SHOULD declare the arguments as I did.
'ByVal is a must, or your app or Visual Basic will crash with an
''PERFORMED AN ILLEGAL OPERATION' message. No turn back. No error handling.

'ARGUMENTS: (Windows calls our function with these arguments)
'MidiInHandle is the same handle for the opened device
'Message is the type of the data (MM_MIM_DATA etc.)
'Instance is the argument that you passed to the dll with MidiInOpen's dwInstance arg.
'     This will also pass to the callback function each time.
'dw1 is the event data in the same format of MidiOutShortMsg's data format.
'dw2 is the time value of the event in milliseconds. It starts from 0 when you called MidiInStart.

'The name of the function and the names of the arguments are up to you.

Public Function Memorize_Event(ByVal MidiInHandle As Long, ByVal Message As Long, _
            ByVal Instance As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'This is useful if there is an error which VB can handle.
'The error will be run over.
On Error Resume Next
'Now, play the note from (or send the same event msg to)
'the internal synth (MidiOut) hMidiOut:tmp2
'(If it is a real event msg)
'First Show it:
If dw1 > 255 Then Form1.Text1.Text = Right("00000000" + Hex(dw1), 8)
'If it is an event, echo it to midi out device
If Message = MM_MIM_DATA Then
   tmp = midiOutShortMsg(tmp2, dw1)
End If
'Now, this program only echoes the midi messages.
'You can do everything with the values. For example,
'(All of this is for that) you can record the time,
'and the events, so this is a complete MIDI INPUT RECORDER!
'You can also convert the total data to a midi file if you know the exact format.
'After this point, IT IS ON YOU!
End Function

