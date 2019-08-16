VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6840
      Top             =   3720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©  2018. All rights reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SoftSkill Developer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   3795
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   1
      Top             =   3240
      Width           =   585
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
Dim a As Long
Dim playsound As String

ProgressBar1.Max = 101
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value = 20 Then
Label2.Caption = "Loading System Files..."
'playsound = sndPlaySound("SoundsCrate-SciFi-PowerUp3.wav", 1)
ElseIf ProgressBar1.Value = 40 Then
Label2.Caption = "Loading Database..."
ElseIf ProgressBar1.Value = 60 Then
Label2.Caption = "Loading Visual Basic Project..."
ElseIf ProgressBar1.Value = 80 Then
Label2.Caption = "Loading Components..."
ElseIf ProgressBar1.Value = 90 Then
Label2.Caption = "Loading Complete..."
ElseIf ProgressBar1.Value = 101 Then
  ProgressBar1.Value = 0
  Me.Hide
  Load MDIMenu
  MDIMenu.Show
  'Load frmLogin
  'frmLogin.Show
  Timer1.Enabled = False
End If
Label4.Caption = ProgressBar1.Value & "%"
End Sub
