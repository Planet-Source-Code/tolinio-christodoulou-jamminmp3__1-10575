VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmList 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3450
   ClientLeft      =   3900
   ClientTop       =   2115
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmList.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3000
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLoad 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inv"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton remTwo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton remOne 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sel"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton addDirSub 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dir"
      DownPicture     =   "frmList.frx":2FB62
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton addFileSub 
      BackColor       =   &H00E0E0E0&
      Caption         =   "File"
      DownPicture     =   "frmList.frx":47B40
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sel"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "REM"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   375
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2940
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4095
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuFilePlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu dia1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuFileClearPlaylist 
         Caption         =   "Clear Playlist"
      End
      Begin VB.Menu dia2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileFileSub 
         Caption         =   "File"
         Begin VB.Menu mnuFileFileInfo 
            Caption         =   "File Info"
         End
         Begin VB.Menu mnuFileEdit 
            Caption         =   "Edit File"
         End
         Begin VB.Menu mnuFilePlaylistInfo 
            Caption         =   "Playlist Entry"
         End
      End
      Begin VB.Menu mnuFileGetMp3s 
         Caption         =   "Get Mp3s"
      End
      Begin VB.Menu fggf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSkins1 
         Caption         =   "Skins"
      End
   End
   Begin VB.Menu mnuFile2 
      Caption         =   "file2"
      Visible         =   0   'False
      Begin VB.Menu mnuFile2About 
         Caption         =   "Jammin Mp3 Player"
      End
      Begin VB.Menu diaFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile2Contact 
         Caption         =   "Contact Us"
      End
      Begin VB.Menu mnuFile2Chk4Upds 
         Caption         =   "Check For Updates"
      End
      Begin VB.Menu mnuFileWeb 
         Caption         =   "Web Page"
      End
      Begin VB.Menu dia0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile2Exit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuFile3 
      Caption         =   "file3"
      Visible         =   0   'False
      Begin VB.Menu nothing 
         Caption         =   "Jammin Mp3 Player"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile3Play 
         Caption         =   "Play"
         Begin VB.Menu mnuFile3PlayFile 
            Caption         =   "File"
         End
         Begin VB.Menu mnuFile3PlayDir 
            Caption         =   "Dir"
         End
         Begin VB.Menu mnuFile3PlayLocation 
            Caption         =   "Location"
         End
         Begin VB.Menu mnuFile3PlayCd 
            Caption         =   "Audio CD:"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainWindow 
         Caption         =   "Main Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "Playlist"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEqualizer 
         Caption         =   "Equalizer"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile3Options 
         Caption         =   "Options"
         Begin VB.Menu mnuOptionsSkins 
            Caption         =   "Change Skin"
         End
         Begin VB.Menu mnuSettings 
            Caption         =   "Settings"
         End
         Begin VB.Menu mnuGetMp3s 
            Caption         =   "Get Mp3s"
         End
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetOut 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyKey As String      'for ini
Dim KeyValue As String     'for ini

Private Sub addDirSub_Click()
Form1.Show
addFileSub.Visible = False
addDirSub.Visible = False

End Sub

Private Sub addDirSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addDirSub.BackColor = &HFF&
addFileSub.BackColor = &HE0E0E0
End Sub

Private Sub addFileSub_Click()

addFileSub.Visible = False
addDirSub.Visible = False

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

Private Sub addFileSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addFileSub.BackColor = &HFF&
addDirSub.BackColor = &HE0E0E0

End Sub

Private Sub cmdRem_Click()
remOne.Visible = True
remTwo.Visible = True
addFileSub.Visible = False
addDirSub.Visible = False
Command7.Visible = False
Command8.Visible = False

End Sub

Private Sub cmdRem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
remOne.BackColor = &HE0E0E0
End Sub

Private Sub Command2_Click()
addFileSub.Visible = False
addDirSub.Visible = False
remOne.Visible = False
remTwo.Visible = False
Command7.Visible = True
Command8.Visible = True

End Sub

Private Sub Command5_Click()
addFileSub.Visible = True
addDirSub.Visible = True
remOne.Visible = False
remTwo.Visible = False
Command7.Visible = False
Command8.Visible = False
addFileSub.BackColor = &HE0E0E0
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addFileSub.BackColor = &HE0E0E0
addDirSub.BackColor = &HE0E0E0

End Sub


Private Sub Command7_Click()
Command7.Visible = False
Command8.Visible = False


Dim file As String
CommonDialog2.DialogTitle = "Load PlayList."
   CommonDialog2.MaxFileSize = 16384
   CommonDialog2.Filename = ""
   CommonDialog2.Filter = "list Files|*.lis"
   CommonDialog2.ShowOpen     ' = 1
If CommonDialog2.Filename = "" Then Exit Sub
file = CommonDialog2.Filename
Dim A As String
Dim X As String
On Error GoTo error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
List1.AddItem A$
List2.AddItem A$
Loop
Close 1
Exit Sub
error:
X = MsgBox("File Not Found", vbOKOnly, "Error")

End Sub

Private Sub Command8_Click()
Command7.Visible = False
Command8.Visible = False

Dim i
On Error GoTo error
CommonDialog2.Filename = "mylist.lis"
CommonDialog2.Filter = "lis Files (*.lis)|*.lis"
CommonDialog2.ShowSave
If CommonDialog2.Filename <> "" Then
    Open CommonDialog2.Filename For Output As #1
    For i = 0 To 100
    List2.ListIndex = i
    Write #1, List1.Text
    
Next i
error:
Close #1

End If


End Sub

Private Sub Form_Load()
'==============ini file

KeySection = "Load Default"
KeyKey = "load"
loadini
txtLoad.Text = KeyValue
If txtLoad.Text = "default" Then
Me.Picture = LoadPicture(App.Path & "\" & "main3.bmp")
ElseIf txtLoad.Text = "blue" Then
Me.Picture = LoadPicture(App.Path & "\" & "main3blue.bmp")
ElseIf txtLoad.Text = "L" Then
Me.Picture = LoadPicture(App.Path & "\" & "main3L.bmp")
End If
'================end ini code
minList = True
'Here we load the list that the user had when it end the program... @!#.lst is the name of the txt file where the list was saved
List_Load List1, "@!#2.lst"
List_Load List2, "@!#.lst"
lstindex = List1.ListIndex
 Me.Left = frmMain.Left
 If EQ = True Then
    Me.Top = frmMain.Top + frmMain.Height + 1830
    Me.Width = frmMain.Width
  ElseIf EQ = False Then
   Me.Top = frmMain.Top + frmMain.Height
    Me.Width = frmMain.Width
    End If
    Dim lngNum As Long

On Error GoTo Fin
If sPlayList(0) <> "" Then
    For lngNum = 0 To UBound(sPlayList)
        List1.AddItem sPlayList(lngNum)
        List2.AddItem sPlayList(lngNum)
    Next
    List1.ListIndex = lngCurrentClip
    List2.ListIndex = lngCurrentClip
End If
Fin:

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
remOne.Visible = False
remTwo.Visible = False
addFileSub.Visible = False
addDirSub.Visible = False
Command7.Visible = False
Command8.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    ' Detect the left mouse button
    If Button = 1 Then
    ' Release the capture of the mouse
    Call ReleaseCapture
    ' Move the form with the mouse
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    addFileSub.BackColor = &HE0E0E0
    addDirSub.BackColor = &HE0E0E0
remOne.BackColor = &HE0E0E0
remTwo.BackColor = &HE0E0E0

End Sub


Private Sub Form_Resize()
Me.Left = frmMain.Left
Me.Width = frmMain.Width
 Me.Height = 3555
End Sub



Private Sub Form_Unload(Cancel As Integer)
minList = False
'Save the current list in the text file so the next time that
'the program will open it will load it
List_Save List2, "@!#.lst"
List_Save List1, "@!#2.lst"

End Sub


Private Sub List1_Click()
If List1.Selected(List1.ListIndex) = True Then
List2.Selected(List1.ListIndex) = True
ElseIf List1.Selected(List1.ListIndex) = False Then
List2.Selected(List1.ListIndex) = False
End If
End Sub

Private Sub List1_DblClick()
List2_DblClick
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then

PopupMenu mnuFile
End If

remOne.Visible = False
remTwo.Visible = False
addFileSub.Visible = False
addDirSub.Visible = False
Command7.Visible = False
Command8.Visible = False
End Sub




Private Sub mnuexit_Click()
frmMain.Check1.Value = 0
End Sub



Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addFileSub.BackColor = &HE0E0E0
addDirSub.BackColor = &HE0E0E0
remOne.BackColor = &HE0E0E0
remTwo.BackColor = &HE0E0E0

End Sub



Private Sub List2_DblClick()
mp3Info.getTheInfo
frmMain.Text6.Caption = List1.Text
frmMain.Text1 = List2.Text
On Error Resume Next
frmMain.MediaPlayer1.Filename = frmMain.Text1.Text
If frmMain.Text1.Text <> "" Then
frmMain.MediaPlayer1.Play
frmMain.Slider1.Max = frmMain.MediaPlayer1.Duration
frmMain.Text5.Text = bitr4text
frmMain.Text2.Text = whatLayer
Else
Exit Sub
End If

End Sub

Private Sub mnuFile3PlayFile_Click()
Dim sFiles() As String, sRoot As String
Dim i As Long

On Error GoTo Fin
With frmMain.CommonDialog1
    .CancelError = True
    .Filename = ""
    .Filter = "MP3 Files|*.mp3"
    .Flags = CommonDialog1OFNAllowMultiselect Or CommonDialog1OFNExplorer
    .MaxFileSize = 32000
    .ShowOpen
    

    sFiles = Split(frmMain.CommonDialog1.Filename, Chr(0))
End With
If UBound(sFiles) = 0 Then
    frmList.List1.AddItem frmMain.CommonDialog1.FileTitle
    frmList.List2.AddItem frmMain.CommonDialog1.FileTitle
    frmList.Show
Else
    sRoot = sFiles(0)
    For i = 1 To UBound(sFiles)
   
        frmList.List1.AddItem sFiles(i)
        frmList.List2.AddItem sFiles(i)
         
    Next i
End If
Exit Sub
Fin:

Exit Sub

End Sub

Private Sub mnuFileFileInfo_Click()
mp3Info.Show
mp3Info.Label2.Caption = List1.Text
End Sub

Private Sub mnuFilePlay_Click()
List1_DblClick


End Sub

Private Sub mnuFileRemove_Click()
If List1.ListIndex = -1 Then
Exit Sub
Else
List1.RemoveItem List1.ListIndex

End If
End Sub

Private Sub mnuGetOut_Click()
Unload frmList
Unload Equalizer
End
End Sub

Private Sub mnuMainWindow_Click()
End
End Sub



Private Sub mnuOptionsSkins_Click()
frmSkin.Show
End Sub

Private Sub mnuSkins1_Click()
frmSkin.Show
End Sub

Private Sub nothing_Click()
frmAbout.Show
End Sub

Private Sub remOne_Click()
remOne.Visible = False
remTwo.Visible = False

If List1.ListIndex = -1 Then
MsgBox "     No file selected    ", vbSystemModal, "Jammin Mp3 Player -No file Selected"
Else
List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex

End If
End Sub

Private Sub remOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
remOne.BackColor = &HFF&
remTwo.BackColor = &HE0E0E0


End Sub

Private Sub remTwo_Click()

remOne.Visible = False
remTwo.Visible = False
List1.Clear
List2.Clear
End Sub
Public Sub List_Save(TheList As ListBox, Filename As String)
    'Save a listbox as FileName
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Output As fFile

    
    For Save = 0 To TheList.ListCount - 1
        Print #fFile, TheList.List(Save)
    Next Save
    Close fFile
    
End Sub

Public Sub List_Add(List As ListBox, txt As String)
    List.AddItem txt
End Sub
Public Sub List_Load(TheList As ListBox, Filename As String)
    'Loads a file to a list box
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Input As fFile


    Do
        Line Input #fFile, TheContents$
        Call List_Add(TheList, TheContents$)
    Loop Until EOF(fFile)
    Close fFile
End Sub
Public Sub endTheList()
List_Save List1, "@!#.lst"
Unload Me
End Sub


Private Sub remTwo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
remTwo.BackColor = &HFF&
remOne.BackColor = &HE0E0E0

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


