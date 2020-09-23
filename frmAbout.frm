VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3270
   ClientLeft      =   7170
   ClientTop       =   4095
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      BorderWidth     =   12
      Height          =   3135
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2055
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Mp3 Player.. Created By Chris Christodoulou"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub




