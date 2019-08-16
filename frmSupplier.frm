VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSupplier 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   5715
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Width           =   5625
      Begin VB.Label lblJlhData 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data Jenis Ketenagaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   2820
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Proses Data Supplier"
      ForeColor       =   &H80000005&
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   5655
      Begin VB.CommandButton btnbatal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "&Batal"
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
         Left            =   4320
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btnTambah 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "&Tambah"
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
         Left            =   120
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btnSimpan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "&Simpan"
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
         Left            =   1560
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btnHapus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "&Hapus"
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
         Left            =   3000
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Insert Data Supplier"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txttelphone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   8
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtalamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2160
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtnama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtkode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Supplier"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Supplier"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   5475
      _cx             =   9657
      _cy             =   5953
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483629
      ForeColorFixed  =   0
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483633
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSupplier.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnbatal_Click()
Call Form_Load
End Sub

Private Sub btnHapus_Click()
 On Error GoTo SalahHapus
    Set RsCek = New ADODB.Recordset
    sql = "Select * From tbsupplier where id_supplier ='" & txtkode.Text & "'"
    RsCek.Open sql, Conn, 3, 4
    If RsCek.RecordCount = 0 Then Exit Sub
        If MsgBox("Anda yakin akan menghapus data?", vbQuestion + vbYesNo, "Hapus Data") = vbNo Then Exit Sub
        Set RsHapus = New ADODB.Recordset
        sql = "DELETE FROM tbsupplier WHERE id_supplier ='" & txtkode.Text & "'"
        RsHapus.Open sql, Conn, adOpenDynamic, adLockOptimistic
        Call Form_Load
        Exit Sub
SalahHapus:
    MsgBox Err.Description, vbCritical, "Kesalahan hapus"
    Exit Sub
End Sub

Private Sub btnSimpan_Click()
If txtkode = "" Or txtnama = "" Or txtalamat = "" Or txttelphone = "" Then
    MsgBox "Objek isian harus diisi.", vbExclamation, "Kesalahan"
    Exit Sub
End If
    On Error GoTo SalahSimpan
    Set RsCek = New ADODB.Recordset
    sql = "Select * From tbsupplier Where id_supplier ='" & txtkode.Text & "'"
    RsCek.Open sql, Conn, adOpenKeyset, adLockReadOnly
    If RsCek.RecordCount = 0 Then
    sql = "INsert into tbsupplier values('" & txtkode & "','" & txtnama & "','" & txtalamat & "','" & txttelphone & "')"
Else
    sql = "Update tbsupplier set nm_supplier = '" & txtnama & "', almt_supplier ='" & txtalamat & "' , tlp_supplier ='" & txttelphone & "'where id_supplier = '" & txtkode.Text & "'"
End If
    Set RsSimpan = New ADODB.Recordset
    RsSimpan.Open sql, Conn, adOpenDynamic, adLockBatchOptimistic
    MsgBox "Data tersimpan ke database!", vbInformation, "Simpan Data"
    Call Form_Load
    
    Exit Sub
SalahSimpan:
    MsgBox Err.Description, vbCritical, "Kesalahan SImpan"
    Exit Sub
End Sub

Private Sub btnTambah_Click()
Call BersihObjek
Set RsCek = New ADODB.Recordset
Call Kode_auto
txtnama.SetFocus
btnbatal.Enabled = True
btnSimpan.Enabled = True
End Sub

Private Sub Form_Load()
Me.Caption = "Data Supplier"
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 8
Call BersihObjek
TampilDataGrid
btnSimpan.Enabled = False
btnHapus.Enabled = False
btnbatal.Enabled = False
btnTambah.Enabled = True
End Sub

Private Sub Kode_auto()
Dim Urutan As String * 6
Dim Hitung As Long
With Grid
Set RsCek = New ADODB.Recordset
sql = "Select *From   tbsupplier   order by id_supplier desc"
RsCek.Open sql, Conn, 3, 4
If RsCek.RecordCount = 0 Then
    Urutan = "SPP" + "001"
    txtkode.Text = Urutan
Else
    Hitung = Right(RsCek!id_supplier, 3) + 1
    Urutan = "SPP" + Right("000" & Hitung, 3)
      
End If
txtkode.Text = Urutan
End With
End Sub

Private Sub BersihObjek()
txtkode.Enabled = False
txtkode.Text = ""
txtnama.Text = ""
txtalamat.Text = ""
txttelphone.Text = ""
End Sub
Private Sub TampilDataGrid()
Set rsGrid = New ADODB.Recordset
sql = "select * from tbsupplier"
rsGrid.Open sql, Conn, 3, 4
Set Grid.DataSource = rsGrid
Dim X As Integer
For X = 0 To Grid.Cols - 1
Grid.FixedAlignment(X) = flexAlignCenterCenter
Next X
lblJlhData.Caption = "Jumlah = " & rsGrid.RecordCount & " Data Supplier"
End Sub



Private Sub Grid_Click()
With Grid
If .Rows = 0 Then Exit Sub
txtkode.Text = .TextMatrix(.Row, 0)
txtnama.Text = .TextMatrix(.Row, 1)
txtalamat.Text = .TextMatrix(.Row, 2)
txttelphone.Text = .TextMatrix(.Row, 3)
End With
btnTambah.Enabled = False
btnSimpan.Enabled = True
btnHapus.Enabled = True
btnbatal.Enabled = True
End Sub
