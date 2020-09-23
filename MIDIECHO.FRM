VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Midi Echo"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   1425
      ItemData        =   "Midiecho.frx":0000
      Left            =   3360
      List            =   "Midiecho.frx":0002
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   1425
      ItemData        =   "Midiecho.frx":0004
      Left            =   240
      List            =   "Midiecho.frx":0006
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   6015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Recording"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Recording"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Data passed:"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Select Output Device"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Input Device"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Private Sub Command1_Click()
'Here, tmp1 is the handle for the MidiIn device
'and tmp2 is the handle for the MidiOut device
'First, open the Input device if one is selected
If List2.ListIndex >= 0 Then
   tmp = midiInOpen(tmp1, List2.ListIndex, AddressOf Memorize_Event, 0, CALLBACK_FUNCTION)
Else
   MsgBox "Please select an Input Device"
   Exit Sub
End If
'Then the output device
If List3.ListIndex >= 0 Then
   tmp = midiOutOpen(tmp2, List3.ListIndex, 0, 0, 0)
Else
   MsgBox "Please select an Output Device"
   midiInClose (tmp1)
   Exit Sub
End If
List2.Enabled = False
List3.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
'Start the recording. This resets the timer to 0
tmp = midiInStart(tmp1)
End Sub

Private Sub Command2_Click()
List2.Enabled = True
List3.Enabled = True
Command2.Enabled = False
Command1.Enabled = True
'Stop the recording, and close the midi in device
tmp = midiInReset(tmp1)
tmp = midiInStop(tmp1)
tmp = midiInClose(tmp1)
tmp = midiOutClose(tmp2)
End Sub

Private Sub Command3_Click()
Visible = False
Unload Me
End
End Sub

Private Sub Form_Load()
'Disable the 'stop recording' button
Command2.Enabled = False
'Here, first getting the number of midi in devices
tmp = midiInGetNumDevs
tmp1 = tmp
'Then examining the capabilities of the devices
For i = 0 To tmp1 - 1
   tmp = midiInGetDevCaps(i, oxMIC, Len(oxMIC))
   If tmp <> 0 Then End
List2.AddItem oxMIC.szPname
Next i
tmp = midiOutGetNumDevs
tmp1 = tmp
'Then examining the capabilities of the devices
For i = 0 To tmp1 - 1
   tmp = midiOutGetDevCaps(i, oxMOC, Len(oxMOC))
   If tmp <> 0 Then End
List3.AddItem oxMOC.szPname
Next i
End Sub

