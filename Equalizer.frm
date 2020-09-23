VERSION 5.00
Begin VB.Form Equalizer 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Equalizer.frx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLoad 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Presets"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auto"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.VScrollBar VScroll8 
      Height          =   855
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll7 
      Height          =   855
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll6 
      Height          =   855
      Left            =   1200
      TabIndex        =   10
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll5 
      Height          =   855
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   855
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   855
      Left            =   2280
      TabIndex        =   7
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll11 
      Height          =   855
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll10 
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll9 
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1095
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preamp"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUALIZER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuSkins2 
         Caption         =   "Skins"
      End
   End
End
Attribute VB_Name = "Equalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I am still trying to figure out how to create an equalizer
'if you know how to do it please submit it
Dim KeyKey As String      'for ini
Dim KeyValue As String     'for ini


Private Sub Form_Load()
'==============ini file

KeySection = "Load Default"
KeyKey = "load"
loadini
txtLoad.Text = KeyValue
If txtLoad.Text = "default" Then
Me.Picture = LoadPicture(App.Path & "\" & "main2.bmp")
ElseIf txtLoad.Text = "blue" Then
Me.Picture = LoadPicture(App.Path & "\" & "main2blue.bmp")
ElseIf txtLoad.Text = "L" Then
Me.Picture = LoadPicture(App.Path & "\" & "main2L.bmp")
End If
'================end ini code
EQ = True
minEq = True
Me.Left = frmMain.Left
    Me.Top = frmMain.Top + frmMain.Height
    Me.Width = frmMain.Width
    Me.Height = 1830
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Button = vbRightButton Then
PopupMenu mnuFile
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Left = frmMain.Left
    Me.Top = frmMain.Top + frmMain.Height
    Me.Width = frmMain.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
EQ = False
minEq = False

End Sub
Public Sub endMe()
EQ = False

Unload Me
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



Private Sub mnuSkins2_Click()
frmSkin.Show
End Sub
