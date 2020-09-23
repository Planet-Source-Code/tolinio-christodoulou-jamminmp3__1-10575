VERSION 5.00
Begin VB.Form mp3Info 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2550
   ClientLeft      =   7170
   ClientTop       =   5895
   ClientWidth     =   3045
   BeginProperty Font 
      Name            =   "Wingdings 2"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2550
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "File Info"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "mp3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bitrate_lookup(7, 15) As Integer
Public actual_bitrate As Long
   Public Function Getmp3data(MP3File As String)
   On Error GoTo theHell
     Dim dIN As String
     cr = Chr(10)
     Open MP3File For Binary As #1
     ' read in 1st 4k of .mp3 file to find a frame header
     dIN = Input(4096, #1)
     filesize = LOF(1) ' needed to calculate track duration
     Close #1
     
     ' frame header starts with 12 set bits [sync]
     ' NB this ignores MPEG-2.5 which is 11 set bits, 1 zero bit.
     
     ' my search for the sync bits only works on nibble boundaries,
     ' I'm not sure if it is necessary to search on bit boundaries -
     ' if so then this search will be 4* slower and require a rewrite
     ' of this search section and shift_those_bits.
     Do Until i = 4095
       i = i + 1
       d1 = Asc(Mid(dIN, i, 1))
       d2 = Asc(Mid(dIN, i + 1, 1))
       If d1 = &HFF And (d2 And &HF0) = &HF0 Then
         'Debug.Print "Found at"; i
         ' get 20 hdr bits - they are last 20 bits of next 3 bytes
         temp_string = Mid(dIN, i + 1, 3)
         mp3bits_string = shift_those_bits(Mid(dIN, i + 1, 3))
         Exit Do
       End If
       ' if we haven't found the sync yet then shift left by 4 bits
       dSHIFT = shift_those_bits(Mid(dIN, i, 3))
       dd1 = Asc(Left(dSHIFT, 1))
       dd2 = Asc(Right(dSHIFT, 1))
       If dd1 = &HFF And (dd2 And &HF0) = &HF0 Then
         'Debug.Print "Found at"; i; "& a nibble"
         ' get 20 hdr bits - they are first 20 bits of next 3 bytes
         mp3bits_string = Mid(dIN, i + 2, 3)
         Exit Do
       End If
     Loop
     
     ' 1st 20 bits of mp3bits_string are hdr info for this frame
     ' 1st bit is ID - 0=MPG-2, 1=MPG-1
     mp3_id = (&H80 And Asc(Left(mp3bits_string, 1))) / 128
     ' next 2 bits are Layer
     mp3_layer = (&H60 And Asc(Left(mp3bits_string, 1))) / 32
     ' next bit is Protection
     mp3_prot = &H10 And Asc(Left(mp3bits_string, 1))
     ' next 4 bits are bitrate
     mp3_bitrate = &HF And Asc(Left(mp3bits_string, 1))
     'next 2 bits are frequency
     mp3_freq = &HC0 And Asc(Mid(mp3bits_string, 2, 1))
     ' next bit is Padding
     mp3_pad = (&H20 And Asc(Mid(mp3bits_string, 2, 1))) / 2
     actual_bitrate = 1000 * CLng((bitrate_lookup((mp3_id * 4) Or mp3_layer, mp3_bitrate)))
     
     dat = "ID: "
     If mp3_id = 0 Then
       dat = dat + "MPEG-2"
       whatLayer = "MPEG-2"
     Else
       dat = dat + "MPEG-1"
       whatLayer = "MPEG-1"
     End If
     
     dat = dat + cr + "Layer: "
      Select Case mp3_layer
        Case 1
          dat = dat + "Layer III"
          whatLayer = "MPEG-III"
        Case 2
          dat = dat + "Layer II"
          whatLayer = "MPEG-II"
        Case 3
          dat = dat + "Layer I"
          whatLayer = "MPEG-I"
      End Select
      dat = dat + cr + "Bitrate: " + Str(actual_bitrate)
      bitr4text = Str(actual_bitrate) / 1000
      Select Case (mp3_id * 4) Or mp3_freq
        Case 0
          sample_rate = 22050
        Case 1
          sample_rate = 24000
        Case 2
          sample_rate = 16000
        Case 4
          sample_rate = 44100
        Case 5
          sample_rate = 48000
        Case 6
          sample_rate = 32000
      End Select
      dat = dat + cr + "Sample rate: " + Str(sample_rate)
      frmMain.Text4.Text = Str(sample_rate)
     
      'display all the info
      lblInfo.Caption = dat
theHell:
      lblInfo.Caption = dat
      
   End Function
   Public Function shift_those_bits(dIN As String) As String
     ' need to left shift 4 bits losing most significant 4 bits
     Dim sd1, sd2, sd3, do1, do2 As Integer
     duff = Left(dIN, 1)
     duff2 = Asc(duff)
     sd1 = Asc(Left(dIN, 1))
     sd2 = Asc(Mid(dIN, 2, 1))
     sd3 = Asc(Right(dIN, 1))
     
     do1 = ((sd1 And &HF) * 16) Or ((sd2 And &HF0) / 16)
     do2 = ((sd2 And &HF) * 16) Or ((sd3 And &HF0) / 16)
     shift_those_bits = Chr(do1) + Chr(do2)
   End Function


Private Sub cmdGo_Click()
On Error GoTo error
File1.Filename = frmList.List1.Text
  fname = File1.Path + "\" + File1.Filename
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (File1.Path + "\" + File1.Filename)
  Else
    lblInfo.Caption = " "
  End If
error:
  Exit Sub
  End Sub





Private Sub Form_Load()
Me.Left = frmList.Left
Me.Top = frmList.Top
  ' setup array for mpeg bitrate info
  bitrate_data = "032,032,032,032,008,008,"
  bitrate_data = bitrate_data + "064,048,040,048,016,016,"
  bitrate_data = bitrate_data + "096,056,048,056,024,024,"
  bitrate_data = bitrate_data + "128,064,056,064,032,032,"
  bitrate_data = bitrate_data + "160,080,064,080,040,040,"
  bitrate_data = bitrate_data + "192,096,080,096,048,048,"
  bitrate_data = bitrate_data + "224,112,096,112,056,056,"
  bitrate_data = bitrate_data + "256,128,112,128,064,064,"
  bitrate_data = bitrate_data + "288,160,128,144,080,080,"
  bitrate_data = bitrate_data + "320,192,160,160,096,096,"
  bitrate_data = bitrate_data + "352,224,192,176,112,112,"
  bitrate_data = bitrate_data + "384,256,224,192,128,128,"
  bitrate_data = bitrate_data + "416,320,256,224,144,144,"
  bitrate_data = bitrate_data + "448,384,320,256,160,160,"
    
  For Y = 1 To 14
    For X = 7 To 5 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
    For X = 3 To 1 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
  Next
End Sub


Private Sub Form_Resize()
cmdGo_Click
End Sub
Public Sub getTheInfo()
cmdGo_Click
End Sub
Public Sub GetTheInfoCMD()
File1.Filename = frmMain.CommonDialog1.FileTitle
  fname = File1.Path + "\" + File1.Filename
  If UCase(Right(fname, 4)) = ".MP3" Then
    Getmp3data (File1.Path + "\" + File1.Filename)
  Else
    lblInfo.Caption = "not an .mp3 file"
  End If
  End Sub


Private Sub Label1_Click()
Unload Me
End Sub


Private Sub lblInfo_Click()
Unload Me
End Sub
