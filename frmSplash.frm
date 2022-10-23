VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmPayroll"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "(c) 2000 Evan Silich"
      BeginProperty Font 
         Name            =   "CityBlueprint"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   1560
      Picture         =   "frmSplash.frx":000C
      Top             =   840
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   3660
      Left            =   -120
      Picture         =   "frmSplash.frx":B676
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'internal call play wav
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub Form_Load()
'create splash screen
frmSplash.Show
frmSplash.Refresh
frmSplash.Timer1.Interval = 1
Load frmPayroll
End Sub

Private Sub Timer1_Timer()
'set wav path
frmPayroll.Show
Unload frmSplash
Dim playsound As Long
    playsound = sndPlaySound("c:\windows\desktop\as snd thunder clap.wav", 0)
End Sub

