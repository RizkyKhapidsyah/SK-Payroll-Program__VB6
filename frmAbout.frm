VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404000&
   Caption         =   "Credits"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6240
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      Caption         =   "(c) 2000 Evan Silich"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "Written By..."
      BeginProperty Font 
         Name            =   "Lucida Blackletter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Payroll Program"
      BeginProperty Font 
         Name            =   "GothicE"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
    'call module
    TextEffect Me, "", 12, 12, , 200, 0, RGB(&H80, 0, 0)
End Sub


Private Sub Label4_Click()
    Do
          Me.Height = Me.Height - 20
          DoEvents
     Loop Until Me.Height = 420
     Do
          Me.Top = Me.Top + 20
          Me.Move Me.Left, Me.Top
          DoEvents
     Loop Until Me.Top > Screen.Height - 2000
     Do
          Me.Left = Me.Left + 20
          Me.Move Me.Left, Me.Top
          DoEvents
     Loop Until Me.Left > Screen.Width
     Unload Me
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set label's font size and color
    Label3.Font.Size = 22
    Label3.ForeColor = &HFFFFFF
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set form's font size and color
    Label3.Font.Size = 20
    Label3.ForeColor = &HC0C0C0
    Label4.Font.Size = 20
    Label4.ForeColor = &HC0C0C0
End Sub

Private Sub Label3_Click()
'call module
    TextEffect Me, "", 12, 12, , 200, 0, RGB(&H80, 0, 0)
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set label's font size and color
    Label4.Font.Size = 22
    Label4.ForeColor = &HFFFFFF
End Sub

