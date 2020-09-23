VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Jammin Mp3"
   ClientHeight    =   3810
   ClientLeft      =   8955
   ClientTop       =   3585
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":044A
   ScaleHeight     =   3810
   ScaleWidth      =   5565
   Begin VB.TextBox txtLoad 
      Height          =   285
      Left            =   720
      TabIndex        =   28
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   600
      Top             =   2040
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2040
   End
   Begin VB.CheckBox OptRes 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   375
   End
   Begin VB.CheckBox OptRND 
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H000000FF&
      Caption         =   "EQ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000FF&
      Caption         =   "PL"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   2880
      Max             =   5000
      Min             =   -5000
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   10
      Left            =   1560
      Max             =   2500
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Text            =   "49"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Text            =   "128"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   8.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Text            =   "------"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   480
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Text            =   "0:00"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1560
      ScaleHeight     =   255
      ScaleWidth      =   2535
      TabIndex        =   20
      Top             =   360
      Width           =   2535
      Begin VB.TextBox text3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   0
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Text6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   0
         Width           =   3015
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   135
      Left            =   600
      TabIndex        =   22
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393216
      LargeChange     =   10
      Max             =   2500
      TickStyle       =   3
      TickFrequency   =   10
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   135
      Left            =   3960
      TabIndex        =   25
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   135
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   1440
      X2              =   1440
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   1440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   1440
      Y1              =   240
      Y2              =   240
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -480
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "khz"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "kbps"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyKey As String      'for ini
Dim KeyValue As String     'for ini

Dim listenabled As Boolean


Private Sub Check1_Click()
If Check1.Value = 1 Then
frmList.Show
minList = True
ElseIf Check1.Value = 0 Then
frmList.endTheList
minList = True
End If
End Sub

Private Sub Check3_Click()

End Sub

Private Sub Command2_Click()

End Sub


Private Sub Check2_Click()
If Check2.Value = 1 Then
Equalizer.Show
minEq = True
ElseIf Check2.Value = 0 Then
Equalizer.endMe
End If
End Sub

Private Sub cmdplay_Click()

Text1 = frmList.List1.Text
On Error Resume Next
MediaPlayer1.Filename = Text1.Text
text3.Text = MediaPlayer1.Filename
If Text1.Text <> "" Then
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration


Else
Exit Sub
End If

End Sub

Private Sub Command1_Click()
If MediaPlayer1.Filename <> "" Then

MediaPlayer1.Pause

Else
Exit Sub
End If

End Sub

Private Sub Command3_Click()
On Error Resume Next
If OptRND.Value = 1 Then GoTo Random
frmList.List1.ListIndex = frmList.List1.ListIndex + 1
frmList.List2.ListIndex = frmList.List2.ListIndex + 1

MediaPlayer1.Filename = frmList.List1.Text
Exit Sub
Random:
Randomize
frmList.List1.ListIndex = Int(Rnd * frmList.List1.ListCount)
MediaPlayer1.Filename = frmList.List1.Text

End Sub

Private Sub Command4_Click()

On Error GoTo theHell
If frmList.List1.ListCount = 0 Then
        Exit Sub
    Else
        If frmList.List1.ListIndex - 1 > -1 Then
           frmList.List2.ListIndex = frmList.List2.ListIndex - 1
           frmList.List1.ListIndex = frmList.List1.ListIndex - 1
            MediaPlayer1.Filename = frmList.List1.Text
        Else
            frmList.List2.ListIndex = frmList.List2.ListCount - 1
            frmList.List1.ListIndex = frmList.List1.ListCount - 1
            MediaPlayer1.Filename = frmList.List1.Text
        End If
    End If
theHell:
    Exit Sub
End Sub

Private Sub Command5_Click()
frmList.List1.Clear
frmList.List2.Clear

On Error GoTo Fin
 
    frmMain.CommonDialog1.CancelError = True
    frmMain.CommonDialog1.Filename = ""
    frmMain.CommonDialog1.Filter = "MP3 Files|*.mp3"
    
    frmMain.CommonDialog1.MaxFileSize = 32000
    frmMain.CommonDialog1.ShowOpen
    

   
    frmList.List1.AddItem frmMain.CommonDialog1.FileTitle
    frmList.List2.AddItem frmMain.CommonDialog1.Filename
    frmList.Show
Exit Sub
Fin:

Exit Sub


End Sub

Private Sub Command7_Click()

Slider1.Value = 0
lblTime.Caption = "--:--:--"
TxtTime.Text = "--:--:--"
MediaPlayer1.Stop

End Sub

Private Sub Form_Load()
HScroll1.Value = 1700
listenabled = False

Me.Height = 1815
Me.Width = 4455
Text6.Caption = "Artist -Song"

'==============ini file

KeySection = "Load Default"
KeyKey = "load"
loadini
txtLoad.Text = KeyValue
If txtLoad.Text = "default" Then
Me.Picture = LoadPicture(App.Path & "\" & "main.bmp")
ElseIf txtLoad.Text = "blue" Then
Me.Picture = LoadPicture(App.Path & "\" & "mainblue.bmp")
ElseIf txtLoad.Text = "L" Then
Me.Picture = LoadPicture(App.Path & "\" & "mainL.bmp")

End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
Timer2.Interval = 1
Text6.Caption = MediaPlayer1.Filename
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If

End Sub







Private Sub Form_Resize()
If Me.WindowState = 0 And minList = True Then
frmList.Show
End If
If Me.WindowState = 0 And minEq = True Then
Equalizer.Show
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmList.endTheList
Equalizer.endMe
End Sub

Private Sub HScroll1_Change()
Timer2.Interval = 0
Text6.Left = 0
Dim pim, sha
sha = HScroll1.Value - 2500
MediaPlayer1.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.min
foo = HScroll1.Value
 Text6.Caption = "Volume " & foo \ 25 & " %"
hell:
Exit Sub
End Sub
Private Sub HScroll1_GotFocus()
text3.Left = 0
Timer2.Interval = 0
Text6.Alignment = 0
End Sub

Private Sub HScroll1_LostFocus()
Text6.Alignment = 2
Timer2.Interval = 1
Text6.Caption = frmList.List1.Text

End Sub



Private Sub HScroll1_Scroll()
Dim pim, sha
sha = HScroll1.Value - 2500
MediaPlayer1.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.min
foo = HScroll1.Value
 Text6.Caption = "Volume " & foo \ 25 & " %"
hell:
Exit Sub

End Sub

Private Sub HScroll2_Change()
On Error GoTo hell
Timer2.Interval = 0
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Text6.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Text6.Caption = "Balance:" & -(MediaPlayer1.Balance / 50) & " % Left "
End If
If HScroll2.Value > 500 Then
Text6.Left = 0
Text6.Alignment = 0
Text6.Caption = "Balance :" & MediaPlayer1.Balance / 50 & " % Right "
End If
MediaPlayer1.Balance = HScroll2.Value
hell:
Exit Sub

End Sub

Private Sub HScroll2_GotFocus()
Text6.Left = 0
Timer2.Interval = 0
text3.Alignment = 0
End Sub

Private Sub HScroll2_LostFocus()
Text6.Alignment = 2
Timer2.Interval = 1
Text6.Caption = MediaPlayer1.Filename

End Sub

Private Sub HScroll2_Scroll()
text3.Left = 0

If HScroll2.Value > -2500 And HScroll2.Value < 2500 Then
Text6.Caption = "Center"

End If
If HScroll2.Value < -2500 Then
Text6.Caption = "Balance:" & -(MediaPlayer1.Balance / 50) & " % Left "
End If
If HScroll2.Value > 2500 Then
Text6.Caption = "Balance:" & MediaPlayer1.Balance / 50 & " % Right "
End If
MediaPlayer1.Balance = HScroll2.Value
End Sub

Private Sub imgMin_Click()
WindowState = 1
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If
 
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
 PopupMenu frmList.mnuFile3
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub

Private Sub Label3_Click()
WindowState = 1
If Check1.Value = 0 Then
frmList.endTheList
ElseIf Check1.Value = 1 Then
frmList.Hide
End If
If Check2.Value = 0 Then
Equalizer.endMe
ElseIf Check2.Value = 1 Then
Equalizer.Hide
End If
End Sub

Private Sub Label4_Click()
Equalizer.endMe
frmList.endTheList
End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
     
   
End Sub










Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
PopupMenu frmList.mnuFile2

End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub



Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If

End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
If OptRND.Value = 1 Then
Randomize Timer
 MyValue = Int((frmList.List1.ListCount * Rnd))
   
    frmList.List1.ListIndex = MyValue
    MediaPlayer1.Filename = frmList.List1.Text
    If frmList.List1.Text <> "" Then
        MediaPlayer1.Play
        text3.Text = MediaPlayer1.Filename
        Slider1.Max = MediaPlayer1.Duration
        
        Exit Sub
    Else
        MsgBox "No file to play", vbOKOnly, "Error"
    End If
Else
    If OptRes.Value = 0 And OptRND.Value = 0 Then
        If frmList.List1.ListIndex <> frmList.List1.ListCount - 1 Then
        frmList.List1.ListIndex = frmList.List1.ListIndex + 1
        
        
        Text1.Text = frmList.List1.Text
        MediaPlayer1.Filename = Text1.Text
       
        text3.Text = MediaPlayer1.Filename
       End If
    End If
    
     
End If
 If OptRes.Value = 1 And OptRND.Value = 0 Then
 Dim k
    k = frmList.List1.ListCount
    If frmList.List1.ListIndex = k - 1 Then
    
    frmList.List1.ListIndex = 0
    End If
 End If
replaylist:


End Sub

Private Sub OptRND_Click()
If OptRND.Value = 1 Then
OptRND.BackColor = &H8000&

ElseIf OptRND.Value = 0 Then
OptRND.BackColor = &H8000000F

End If
End Sub

Private Sub Picture1_Click()
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If

End Sub

Private Sub Slider1_Click()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub
Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If

End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub


Private Sub text3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If
End Sub

Private Sub text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
End Sub









Private Sub Text6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu frmList.mnuFile3
End If

End Sub

Private Sub text6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If EQ = True Then
   Equalizer.Left = Me.Left
    Equalizer.Top = Me.Top + Me.Height
       frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height + 1830
    ElseIf EQ = False Then
    frmList.Left = Me.Left
    frmList.Top = Me.Top + Me.Height
    End If
End Sub


Private Sub Timer1_Timer()
On Error GoTo error
Slider1.Value = MediaPlayer1.CurrentPosition
tinseconden = MediaPlayer1.CurrentPosition
Dim min As Integer
Dim sec As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
If sec = "-1" Then sec = "0"
If sec < 10 Then
TxtTime.Text = min & ":0" & sec
lblTime.Caption = min & ":0" & sec
ElseIf sec >= 10 Then
TxtTime.Text = min & ":" & sec
lblTime.Caption = min & ":" & sec
lblTime.Width = MediaPlayer1.Filename
End If

error:
Exit Sub

End Sub

Private Sub Timer2_Timer()
If Text6.Left < Picture1.Width - Picture1.Width - Text6.Width Then
    Text6.Left = Picture1.Width - 1
    
    Text6.Left = Text6.Left - 5
    
Else
    Text6.Left = Text6.Left - 10
    
End If

End Sub




Private Sub loadini()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub


Private Sub saveini()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub

