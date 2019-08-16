VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login System"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmd_login 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Txt_Password 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Txt_User 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
End
End Sub

Private Sub cmd_login_Click()
If Txt_Password.Text = "" Then
    MsgBox "Isikan User dan Password Anda !", vbCritical, "Kosong..."
    Exit Sub
End If
LoginAja
'MenuUtama.Enabled = True
'Login.Hide
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Me.Caption = "Login System"
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Load()
Bersih
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 3
Txt_Password.Text = "admins"
Txt_User.Text = "admins"
End Sub
Sub Bersih()
Txt_Password.Text = ""
Txt_User.Text = ""
End Sub
Private Sub LoginAja()
Set RsLogin = New ADODB.Recordset
Csql = "Select * from tbuser WHERE username='" & Trim(Txt_User.Text) & "' And password='" & Trim(Txt_Password.Text) & "'"
RsLogin.Open Csql, Conn, adOpenKeyset, adLockReadOnly
If RsLogin.EOF Then
    Dicoba = Dicoba + 1
    MsgBox "Kesempatan " & Dicoba & "Salah", vbCritical, "Error Login..."
    Txt_Password.SetFocus
    SendKeys "{Home}+{End}"
            If Dicoba = 3 Then
                MsgBox "Silahkan Keluar...", vbCritical, "Keluar"
                End
            End If
Else
    frmSplash.Show
    frmLogin.Hide
    
End If
    End Sub

