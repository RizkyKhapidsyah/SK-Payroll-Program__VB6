VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPayroll 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Money"
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9315
   Icon            =   "Evan's Payroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9315
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Left            =   3960
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Evan's Payroll.frx":000C
   End
   Begin VB.CommandButton cmdOvertime 
      Caption         =   "Calculate Overtime"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cboOvertime 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox cboHours 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cboPercent 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlPayroll 
      Left            =   7920
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   6000
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdTax 
      Caption         =   "Taxes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPayRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdAfterTax 
      Caption         =   "After Taxes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Hours"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   7560
      TabIndex        =   24
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lblOvertime 
      BackColor       =   &H00400000&
      Caption         =   "Over Time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3480
      TabIndex        =   22
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblMod 
      BackColor       =   &H00400000&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblTaxPercent 
      BackColor       =   &H00400000&
      Caption         =   "Tax Percent:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   20
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   7560
      TabIndex        =   19
      Top             =   120
      Width           =   60
   End
   Begin VB.Label lblAfterTaxes 
      BackColor       =   &H00400000&
      Caption         =   "After Taxes:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblTaxDeduction 
      BackColor       =   &H00400000&
      Caption         =   "Tax Deduction:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lblBeforeTax 
      BackColor       =   &H00400000&
      Caption         =   "Before Taxes:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblHourlyWage 
      BackColor       =   &H00400000&
      Caption         =   "Hourly Wage:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblHours 
      BackColor       =   &H00400000&
      Caption         =   "Hours:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblAfterTax 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label lblTax 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&SaveAs"
      End
      Begin VB.Menu Hyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditColor 
         Caption         =   "&Color"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditHyphen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAbout 
         Caption         =   "&About Author"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************************************************************************************************
'***Author: Evan Silich
'***Program: PayRoll Calculation
'***Created: 08/17/00
'*****************************************************************************************************

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Dim blnChange As Boolean

Private Sub cmdAfterTax_Click()
'calculate net pay
On Error GoTo Errhandler
    lblAfterTax.Caption = Format(lblMoney - lblTax, "Currency")
    Exit Sub
Errhandler:
    MsgBox "Missing Data!", vbCritical, "Error"
End Sub

Private Sub cmdCalculate_Click()
'calculate gross pay
   On Error GoTo Errhandler
   
   lblMoney.Caption = Format(Val(cboHours.Text) * (Val(txtPayRate.Text)), "Currency")
    
    If cboHours.Text = "" Then
    MsgBox "You must select hours worked!", vbExclamation, "Error"
     
    End If

    Exit Sub
Errhandler:
    MsgBox "You must select hours worked!", vbExclamation, "Error"
End Sub

Private Sub cmdClear_Click()
   'clear text
    txtPayRate.Text = ""
    lblMoney.Caption = ""
    lblTax.Caption = ""
    lblAfterTax.Caption = ""
    
End Sub

Private Sub cmdEnd_Click()
'create save prompt
Const conbtns As Integer = vbYesNoCancel + vbExclamation
    
    Const message  As String = "Do you want to save the following changes?"
    
        Dim intUserResponse As Integer
        If blnChange = True Then
            intUserResponse = MsgBox(message, conbtns)
            Select Case intUserResponse
        Case vbYes
            Call mnuFileSave_Click
        Case vbCancel
            frmPayroll.SetFocus
        Case vbNo
        Me.Caption = "Good Bye"
        Do
          Me.Top = Me.Top + 20
          Me.Move Me.Left, Me.Top
          DoEvents
     Loop Until Me.Top > Screen.Height - 100
        End
        End Select
        End If

End Sub



Private Sub cmdPrint_Click()
PrintForm
End Sub

Private Sub cmdTax_Click()
'calculate tax
    On Error GoTo Errhandler
    lblTax.Caption = Format((lblMoney.Caption * cboPercent), "Currency")
   Exit Sub
Errhandler:
   MsgBox "You must select a percent!", vbExclamation, "Error"
End Sub

Private Sub cmdovertime_Click()
'set error prompt
If cboHours.Text < "40" Then
        
        MsgBox "You cannot have overtime if didn't work 40 hours!", vbExclamation, "Duh"
        cboOvertime.Text = "0"
        Me.Refresh
        
        Else
'calculate time and a half...not double time
On Error GoTo Errhandler
   Dim OverTime As Currency
   Dim HalfPayRate As Currency
   Dim OverTimePayRate As Currency
   Dim Total As Currency
   HalfPayRate = txtPayRate.Text / 2
   
   OverTimePayRate = txtPayRate.Text + HalfPayRate
   
   OverTime = cboOvertime.Text * OverTimePayRate
Total = OverTime
lblMoney.Caption = Format(Val(Total) + (lblMoney.Caption), "Currency")
End If

Exit Sub
Errhandler:
MsgBox "Missing Data!", vbCritical, "Error"

End Sub

Private Sub Form_Load()
   
    
    'fill combobox with percent
    cboPercent.AddItem ".15"
    cboPercent.AddItem ".16"
    cboPercent.AddItem ".17"
    cboPercent.AddItem ".18"
    cboPercent.AddItem ".19"
    cboPercent.AddItem ".20"
    cboPercent.AddItem ".21"
    cboPercent.AddItem ".22"
    cboPercent.AddItem ".23"
    cboPercent.AddItem ".24"
    cboPercent.AddItem ".25"
    cboPercent.AddItem ".26"
    cboPercent.AddItem ".27"
    cboPercent.AddItem ".28"
    cboPercent.AddItem ".29"
    cboPercent.AddItem ".30"
    
    'fill combobox with hours
    cboHours.AddItem "1"
    cboHours.AddItem "2"
    cboHours.AddItem "3"
    cboHours.AddItem "4"
    cboHours.AddItem "5"
    cboHours.AddItem "6"
    cboHours.AddItem "7"
    cboHours.AddItem "8"
    cboHours.AddItem "9"
    cboHours.AddItem "10"
    cboHours.AddItem "11"
    cboHours.AddItem "12"
    cboHours.AddItem "13"
    cboHours.AddItem "14"
    cboHours.AddItem "15"
    cboHours.AddItem "16"
    cboHours.AddItem "17"
    cboHours.AddItem "18"
    cboHours.AddItem "19"
    cboHours.AddItem "20"
    cboHours.AddItem "21"
    cboHours.AddItem "22"
    cboHours.AddItem "23"
    cboHours.AddItem "24"
    cboHours.AddItem "25"
    cboHours.AddItem "26"
    cboHours.AddItem "27"
    cboHours.AddItem "28"
    cboHours.AddItem "29"
    cboHours.AddItem "30"
    cboHours.AddItem "31"
    cboHours.AddItem "32"
    cboHours.AddItem "33"
    cboHours.AddItem "34"
    cboHours.AddItem "35"
    cboHours.AddItem "36"
    cboHours.AddItem "37"
    cboHours.AddItem "38"
    cboHours.AddItem "39"
    cboHours.AddItem "40"
    
    'fill combobox with over time hours
    cboOvertime.AddItem "0"
    cboOvertime.AddItem "1"
    cboOvertime.AddItem "2"
    cboOvertime.AddItem "3"
    cboOvertime.AddItem "4"
    cboOvertime.AddItem "5"
    cboOvertime.AddItem "6"
    cboOvertime.AddItem "7"
    cboOvertime.AddItem "8"
    cboOvertime.AddItem "9"
    cboOvertime.AddItem "10"
    cboOvertime.AddItem "11"
    cboOvertime.AddItem "12"
    cboOvertime.AddItem "13"
    cboOvertime.AddItem "14"
    cboOvertime.AddItem "15"
    cboOvertime.AddItem "16"
    cboOvertime.AddItem "17"
    cboOvertime.AddItem "18"
    cboOvertime.AddItem "19"
    cboOvertime.AddItem "20"
    
    'scroll forms caption
    Timer1.Interval = 5
    
    frmPayroll.Caption = ""
        NewCaption = "Payroll Program       "
        counter = 1
        direction = 1
        totalcount = Len(NewCaption)
   
   'time/date
   lblTime.Caption = Time
   lblDate.Caption = Date
   

    'center screen
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
    
    'Change value depending On the speed of flahing.
    Timer1.Interval = 300
End Sub

Private Sub lblTime_Change()
'set true/false for save prompt
blnChange = True
End Sub


Private Sub mnFileOpen_Click()

'set properties for opening a data file

RichTextBox1.Visible = True
Dim strOpenFile As String
cdlPayroll.Filter = "Data Files(*.dat)|*.dat|All Files(*.*)|*.*"
cdlPayroll.FileName = ""
cdlPayroll.ShowOpen
strOpenFile = cdlPayroll.FileName
RichTextBox1.LoadFile strOpenFile

End Sub

Private Sub mnuEditAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuEditColor_Click()
'cancelError set in property pages
cdlPayroll.CancelError = True
On Error GoTo Errhandler
'display color dialog box
cdlPayroll.ShowColor
'set the form's background color
'selected color
frmPayroll.BackColor = cdlPayroll.Color
lblTaxPercent.BackColor = cdlPayroll.Color
lblMoney.BackColor = cdlPayroll.Color
lblTax.BackColor = cdlPayroll.Color
lblAfterTax.BackColor = cdlPayroll.Color
lblHours.BackColor = cdlPayroll.Color
lblBeforeTax.BackColor = cdlPayroll.Color
lblHourlyWage.BackColor = cdlPayroll.Color
lblTaxDeduction.BackColor = cdlPayroll.Color
lblAfterTaxes.BackColor = cdlPayroll.Color
lblMod.BackColor = cdlPayroll.Color
lblTime.BackColor = cdlPayroll.Color
lblDate.BackColor = cdlPayroll.Color
lblOvertime.BackColor = cdlPayroll.Color

'fix background/forecolor
If cdlPayroll.Color = vbWhite Then
    
    lblTime.ForeColor = vbBlack
    lblDate.ForeColor = vbBlack
    lblTaxPercent.ForeColor = vbBlack
    lblHours.ForeColor = vbBlack
    lblOvertime.ForeColor = vbBlack
    lblHourlyWage.ForeColor = vbBlack
    lblBeforeTax.ForeColor = vbBlack
    lblTaxDeduction.ForeColor = vbBlack
    lblAfterTaxes.ForeColor = vbBlack
    lblMoney.ForeColor = vbBlack
    lblTax.ForeColor = vbBlack
    lblAfterTax.ForeColor = vbBlack
    lblMod.ForeColor = vbBlack
        
        Else
        
        lblTime.ForeColor = vbWhite
        lblDate.ForeColor = vbWhite
        lblTaxPercent.ForeColor = vbWhite
        lblHours.ForeColor = vbWhite
        lblOvertime.ForeColor = vbWhite
        lblHourlyWage.ForeColor = vbWhite
        lblBeforeTax.ForeColor = vbWhite
        lblTaxDeduction.ForeColor = vbWhite
        lblAfterTaxes.ForeColor = vbWhite
        lblMoney.ForeColor = vbWhite
        lblTax.ForeColor = vbWhite
        lblAfterTax.ForeColor = vbWhite
        lblMod.ForeColor = vbWhite

End If

Exit Sub
Errhandler:
    MsgBox "No color chosen.", vbInformation
End Sub

Private Sub mnuFileExit_Click()
'create save prompt
Const conbtns As Integer = vbYesNoCancel + vbExclamation
    
    Const message  As String = "Do you want to save the following changes?"
    
        Dim intUserResponse As Integer
        If blnChange = True Then
            intUserResponse = MsgBox(message, conbtns)
            Select Case intUserResponse
        Case vbYes
            Call mnuFileSave_Click
        Case vbCancel
            frmPayroll.SetFocus
        Case vbNo
        
            DrawWidth = 4
            For i = 1 To 16000
            Down = Down + 1
            Across = Across + 1
            PSet (Rnd * Across, Rnd * Down), QBColor(Rnd * 15)
            Next i
            End
        End Select
        End If

End Sub

Private Sub mnuFileNew_Click()
'cleat text
    txtPayRate.Text = ""
    lblMoney.Caption = ""
    lblTax.Caption = ""
    lblAfterTax.Caption = ""
End Sub



Private Sub mnuFilePrint_Click()
'print form
PrintForm

End Sub



Private Sub mnuFileSave_Click()
'set properties for saving a data file
cdlPayroll.Filter = "Data Files(*.dat)|*.dat|All Files(*.*)|*.*"

Call mnuFileSaveAs_Click

End Sub

Private Sub mnuFileSaveAs_Click()
'set properties for saving a data file

cdlPayroll.Filter = "Data Files(*.dat)|*.dat|All Files(*.*)|*.*"
cdlPayroll.FileName = ""
cdlPayroll.ShowSave
Open cdlPayroll.FileName For Output As #1
Write #1, lblTime.Caption _
 ; lblTaxPercent.Caption, Val(cboPercent.Text) _
 ; lblHours.Caption, Val(cboHours.Text) _
 ; lblOvertime.Caption, Val(cboOvertime.Text) _
 ; lblHourlyWage.Caption, Val(txtPayRate.Text) _
 ; lblBeforeTax.Caption, lblMoney.Caption _
 ; lblTaxDeduction.Caption, lblTax.Caption _
 ; lblAfterTaxes.Caption, lblAfterTax.Caption
Close #1



End Sub

Private Sub RichTextBox1_click()
RichTextBox1.Visible = False
    
End Sub

Private Sub Timer1_Timer()
    'scrolls forms caption
    frmPayroll.Caption = Left(NewCaption, counter) & directionNull
        If counter = totalcount Then direction = 2
        If counter = 0 Then direction = 1

        Select Case direction
            Case 1
            counter = counter + 1
            directionNull = ""
            Case 2
            counter = counter - 1
            directionNull = ""
        End Select
    
    'flash forms title bar
    FlashWindow hwnd, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'create save prompt
    Const conbtns As Integer = vbYesNo + vbExclamation
    
    Const message  As String = "Do you want to save the following changes?"
    
        Dim intUserResponse As Integer
        If blnChange = True Then
            intUserResponse = MsgBox(message, conbtns)
            Select Case intUserResponse
        Case vbYes
            Call mnuFileSave_Click
            
        Case vbNo
         Do
          Me.Height = Me.Height - 20
          DoEvents
     Loop Until Me.Height = 420
     Do
          Me.Top = Me.Top + 20
          Me.Move Me.Left, Me.Top
          DoEvents
     Loop Until Me.Top > Screen.Height - 100
    
            End
        End Select
        End If
                
        
End Sub


