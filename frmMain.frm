VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Box Hook Demo"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkModal 
      Caption         =   "Modal (I'll leave that for you to finish :))"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox chkDontBeep 
      Caption         =   "Don't Beep (TOFIX: Won't display icon)"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   3135
   End
   Begin VB.OptionButton optIconNone 
      Caption         =   "(None)"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton optCritical 
      Caption         =   "Critical"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton optExclamation 
      Caption         =   "Exclamation"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optInformation 
      Caption         =   "Information"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optQuestion 
      Caption         =   "Question"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtThirdButton 
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Text            =   "Poor!"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtSecondButton 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "OK"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtFirstButton 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Excellent!"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtMsgBoxMsg 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "How do you rate this demo?"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtMsgBoxTitle 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "MsgBox Hook Demo"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblMsgBoxIcon 
      AutoSize        =   -1  'True
      Caption         =   "Message Box Icon:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label lblMsgBoxButtons 
      AutoSize        =   -1  'True
      Caption         =   "Message Box Buttons:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Label lblMsgBoxMsg 
      AutoSize        =   -1  'True
      Caption         =   "Message Box Prompt:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label lblMsgBoxTitle 
      AutoSize        =   -1  'True
      Caption         =   "Message Box Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExitMsg As Boolean

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
  Select Case MessageBoxH(Me.hwnd, GetDesktopWindow())
    Case IDABORT: If Not ExitMsg Then MsgBox txtFirstButton.Text + " selected"
    Case IDRETRY: If Not ExitMsg Then MsgBox txtSecondButton.Text + " selected"
    Case IDIGNORE: If Not ExitMsg Then MsgBox txtThirdButton.Text + " selected"
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  txtMsgBoxMsg.Text = "If you want to find out how to make the message box modal and a bunch of other stuff, check it out for yourself, or wait for my upcoming new version of this demo VERY soon." & vbCrLf & vbCrLf & "P.S.: The upcoming version will also allow you to have less than three buttons on the box."
  txtMsgBoxTitle.Text = "Bye!"

  optInformation.Value = True

  txtFirstButton.Text = "E&xit"
  txtSecondButton.Text = "E&nd"
  txtThirdButton.Text = "&Quit"

  ExitMsg = True
  cmdGenerate_Click

  End
End Sub
