VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VisualEQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "EQ"
   ClientHeight    =   3030
   ClientLeft      =   105
   ClientTop       =   2190
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   6000
   Begin VB.Timer Timer2 
      Interval        =   4
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   504
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar6 
      Height          =   2196
      Left            =   3096
      TabIndex        =   5
      Top             =   216
      Width           =   276
      _ExtentX        =   503
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar5 
      Height          =   2196
      Left            =   2520
      TabIndex        =   4
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   2196
      Left            =   1944
      TabIndex        =   3
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   2196
      Left            =   1368
      TabIndex        =   2
      Top             =   216
      Width           =   276
      _ExtentX        =   503
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   2196
      Left            =   792
      TabIndex        =   1
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2196
      Left            =   216
      TabIndex        =   0
      Top             =   216
      Width           =   276
      _ExtentX        =   503
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar7 
      Height          =   2196
      Left            =   3672
      TabIndex        =   6
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar8 
      Height          =   2196
      Left            =   4248
      TabIndex        =   7
      Top             =   216
      Width           =   276
      _ExtentX        =   503
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar9 
      Height          =   2196
      Left            =   4824
      TabIndex        =   8
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar10 
      Height          =   2196
      Left            =   5400
      TabIndex        =   9
      Top             =   216
      Width           =   276
      _ExtentX        =   476
      _ExtentY        =   3889
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "16kHz"
      Height          =   228
      Index           =   9
      Left            =   5256
      TabIndex        =   19
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "8kHz"
      Height          =   228
      Index           =   8
      Left            =   4680
      TabIndex        =   18
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "4kHz"
      Height          =   228
      Index           =   7
      Left            =   4104
      TabIndex        =   17
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "2kHz"
      Height          =   228
      Index           =   6
      Left            =   3528
      TabIndex        =   16
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1kHz"
      Height          =   228
      Index           =   5
      Left            =   2952
      TabIndex        =   15
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "31Hz"
      Height          =   228
      Index           =   4
      Left            =   2376
      TabIndex        =   14
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "62Hz"
      Height          =   228
      Index           =   3
      Left            =   1800
      TabIndex        =   13
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "125Hz"
      Height          =   228
      Index           =   2
      Left            =   1224
      TabIndex        =   12
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "250Hz"
      Height          =   228
      Index           =   1
      Left            =   648
      TabIndex        =   11
      Top             =   2592
      Width           =   516
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "500Hz"
      Height          =   228
      Index           =   0
      Left            =   72
      TabIndex        =   10
      Top             =   2592
      Width           =   516
   End
End
Attribute VB_Name = "VisualEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hmixer As Long                  ' mixer handle
Dim inputVolCtrl As MIXERCONTROL    ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL   ' microphone volume control
Dim rc As Long                      ' return code
Dim OK As Boolean                   ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long
Private LVL As LevLights
Private Sub LightsA()
' ProgressBar1
For LVL.InOutLev = CDbl(LVL.VolLev * 5) To Frequency.Freq500Hz  '100
Next
ProgressBar1.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsB()
' ProgressBar2
For LVL.InOutLev = CDbl(LVL.VolLev * 2.5) To Frequency.Freq250Hz  '95
Next LVL.InOutLev
ProgressBar2.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsC()
' ProgressBar3
For LVL.InOutLev = CDbl(LVL.VolLev * 1.25) To Frequency.Freq125Hz '90
Next LVL.InOutLev
ProgressBar3.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsD()
' ProgressBar4
For LVL.InOutLev = CDbl(LVL.VolLev * 0.62) To Frequency.Freq62Hz '85
Next LVL.InOutLev
ProgressBar4.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsE()
' ProgressBar5
For LVL.InOutLev = CDbl(LVL.VolLev * 0.31) To Frequency.Freq31Hz '80
Next LVL.InOutLev
ProgressBar5.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsF()
' ProgressBar6
For LVL.InOutLev = CDbl(LVL.VolLev * 0.01) To Frequency.Freq1kHz '75
Next LVL.InOutLev
ProgressBar6.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsG()
' ProgressBar7
For LVL.InOutLev = CDbl(LVL.VolLev * 0.02) To Frequency.Freq2kHz '70
Next LVL.InOutLev
ProgressBar7.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsH()
' ProgressBar8
For LVL.InOutLev = CDbl(LVL.VolLev * 0.04) To Frequency.Freq4kHz '65
Next LVL.InOutLev
ProgressBar8.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsI()
' ProgressBar9
For LVL.InOutLev = CDbl(LVL.VolLev * 0.08) To Frequency.Freq8kHz '60
Next LVL.InOutLev
ProgressBar9.Value = CDbl(LVL.InOutLev)
End Sub
Private Sub LightsJ()
' ProgressBar10
For LVL.InOutLev = CDbl(LVL.VolLev * 0.16) To Frequency.Freq16kHz '55
Next LVL.InOutLev
ProgressBar10.Value = CDbl(LVL.InOutLev)
End Sub

Private Sub Form_Load()
Timer1.Interval = CCur(5)
Timer2.Interval = CCur(5)
' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   ProgressBar1.Max = Frequency.Freq500Hz + 1 '101
   ProgressBar1.min = Frequency.Freq500Hz
   ProgressBar2.Max = Frequency.Freq250Hz + 1
   ProgressBar2.min = Frequency.Freq250Hz
   ProgressBar3.Max = Frequency.Freq125Hz + 1 '101
   ProgressBar3.min = Frequency.Freq125Hz '90
   ProgressBar4.Max = Frequency.Freq62Hz + 1 '101
   ProgressBar4.min = Frequency.Freq62Hz '85
   ProgressBar5.Max = Frequency.Freq31Hz + 1 '101
   ProgressBar5.min = Frequency.Freq31Hz '80
   ProgressBar6.Max = Frequency.Freq1kHz + 1 '101
   ProgressBar6.min = Frequency.Freq1kHz '75
   ProgressBar7.Max = Frequency.Freq2kHz + 1 '101
   ProgressBar7.min = Frequency.Freq2kHz '70
   ProgressBar8.Max = Frequency.Freq4kHz + 1 '101
   ProgressBar8.min = Frequency.Freq4kHz '65
   ProgressBar9.Max = Frequency.Freq8kHz + 1 '101
   ProgressBar9.min = Frequency.Freq8kHz '60
   ProgressBar10.Max = Frequency.Freq16kHz + 1 '101
   ProgressBar10.min = Frequency.Freq16kHz '55
   Else
      MsgBox "Couldn't get waveout meter"
   End If
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1



End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
End Sub

Private Sub Timer1_Timer()
LVL.VolLev = volume / 327.67
If (volume < 0) Then
volume = -volume
End If
' Get the current output level
  If (1 = 1) Then
  mxcd.dwControlID = outputVolCtrl.dwControlID
  mxcd.Item = outputVolCtrl.cMultipleItems
  rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
  CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
  If (volume < 0) Then volume = -volume
End If
LightsA
LightsB
LightsC
LightsD
LightsE
End Sub

Private Sub Timer10_Timer()

End Sub

Private Sub Timer2_Timer()
LVL.VolLev = volume / 327.67
If (volume < 0) Then
volume = -volume
End If
' Get the current output level
  If (1 = 1) Then
  mxcd.dwControlID = outputVolCtrl.dwControlID
  mxcd.Item = outputVolCtrl.cMultipleItems
  rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
  CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
  If (volume < 0) Then volume = -volume
End If
LightsF
LightsG
LightsH
LightsI
LightsJ
End Sub
