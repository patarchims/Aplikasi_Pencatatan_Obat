VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmPembelian 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10980
   Begin VSFlex7Ctl.VSFlexGrid GrdSupplier 
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   7875
      _cx             =   13891
      _cy             =   2143
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPembelian.frx":0000
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
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   10695
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
         Height          =   615
         Left            =   3480
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1335
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
         Height          =   615
         Left            =   120
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1575
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
         Height          =   615
         Left            =   1920
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   7440
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton btnOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtSubTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   7920
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtJumlah 
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
      Left            =   7200
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtkdObat 
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
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtNmObat 
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtharga 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   5280
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtpemasok 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox txtnota 
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtkdSupp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DptTanggal 
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97845251
      CurrentDate     =   43093
   End
   Begin VSFlex7Ctl.VSFlexGrid GrdObat 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   7035
      _cx             =   12409
      _cy             =   2143
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPembelian.frx":0074
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
   Begin VSFlex7Ctl.VSFlexGrid Grid 
      Height          =   2535
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   10755
      _cx             =   18971
      _cy             =   4471
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPembelian.frx":00E1
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
   Begin VB.Label lblNomor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OBAT"
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
      TabIndex        =   26
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "@"
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
      Left            =   7320
      TabIndex        =   19
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUB TOTAL"
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
      Left            =   7320
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HARGA"
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
      Left            =   4680
      TabIndex        =   17
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA OBAT"
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
      Left            =   2280
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OBAT"
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
      Left            =   -840
      TabIndex        =   15
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PEMASOK"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "KODE"
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
      Left            =   -480
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
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
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHapus_Click()
On Error GoTo SalahHapus
    Set RsCek = New ADODB.Recordset
    sql = "Select * From tbpembelian_detail where nomor ='" & lblNomor.Caption & "'"
    RsCek.Open sql, Conn, 3, 4
    If RsCek.RecordCount = 0 Then Exit Sub
        If MsgBox("Anda yakin akan menghapus data?", vbQuestion + vbYesNo, "Hapus Data") = vbNo Then Exit Sub
        Set RsHapus = New ADODB.Recordset
        sql = "DELETE FROM tbpembelian_detail WHERE nomor ='" & lblNomor.Caption & "'update tbobat set stok=stok - " & txtJumlah.Text & " where  id_obat = '" & txtkdObat.Text & "'"
        RsHapus.Open sql, Conn, adOpenDynamic, adLockOptimistic
        Call Form_Load
        BErsihBELi2
        lblNomor.Caption = ""
        TampilBeli
        btnHapus.Enabled = False
        btnSimpan.Enabled = True
        TampilJumlah
        Exit Sub
SalahHapus:
    MsgBox Err.Description, vbCritical, "Kesalahan hapus"
    Exit Sub
End Sub

Private Sub btnOK_Click()
If txtkdObat = "" Or txtNmObat = "" Or txtJumlah = "" Or txtharga.Text = "" Or txtSubTotal.Text = "" Then
    MsgBox "Objek isian harus diisi.", vbExclamation, "Kesalahan"
    Exit Sub
End If
On Error GoTo SalahSimpan
    Set RsCek = New ADODB.Recordset
    sql = "Select * From tbpenjualan_detail "
    RsCek.Open sql, Conn, adOpenKeyset, adLockReadOnly
    sql = "Insert Into tbpembelian_detail  values ('" & txtnota & "','" & txtkdObat & "','" & txtJumlah & "','" & txtharga & "','" & txtSubTotal & "') update tbobat set stok=stok + " & txtJumlah.Text & " where  id_obat = '" & txtkdObat.Text & "'"
    Set RsSimpan = New ADODB.Recordset
    RsSimpan.Open sql, Conn, adOpenDynamic, adLockBatchOptimistic
         MsgBox "Data tersimpan ke database!", vbInformation, "Simpan Data"
   
    BErsihBELi2
    TampilBeli
    TampilJumlah
    Exit Sub
SalahSimpan:
    MsgBox Err.Description, vbCritical, "Kesalahan SImpan"
    Exit Sub
End Sub

Sub TampilBeliKOsong()
Set rsCari = New ADODB.Recordset
sql = "select tbpembelian_detail.nomor, tbobat.nm_obat, tbpembelian_detail.jml_masuk, tbpembelian_detail.hrg_beli,tbpembelian_detail.sub_total from tbpembelian_detail inner join tbobat on tbobat.id_obat=tbpembelian_detail.id_obat where id_beli = '1' "
rsCari.Open sql, Conn, adOpenKeyset, adLockReadOnly
Grid.Refresh
Set Grid.DataSource = rsCari
Dim X As Integer
For X = 0 To Grid.Cols - 1
Grid.FixedAlignment(X) = flexAlignCenterCenter
Next X
End Sub
Sub BErsihBELi2()
txtkdObat.Text = ""
txtNmObat.Text = ""
txtJumlah.Text = ""
txtharga.Text = ""
txtSubTotal.Text = ""
End Sub

Sub TampilBeli()
Set rsCari = New ADODB.Recordset
sql = "select tbpembelian_detail.nomor, tbobat.id_obat, tbobat.nm_obat, tbpembelian_detail.jml_masuk, tbpembelian_detail.hrg_beli,tbpembelian_detail.sub_total from tbpembelian_detail inner join tbobat on tbobat.id_obat=tbpembelian_detail.id_obat where id_beli Like '%" & txtnota.Text & "%'  "
rsCari.Open sql, Conn, adOpenKeyset, adLockReadOnly
Grid.Refresh
Set Grid.DataSource = rsCari
Dim X As Integer
For X = 0 To Grid.Cols - 1
Grid.FixedAlignment(X) = flexAlignCenterCenter
Next X
End Sub

Private Sub btnSimpan_Click()
'If txtnama = "" Or txttmptlahir = "" Or cbokelompok = "" Or cbojk = "" Or txt_notelp = "" Or txt_kecamatan = "" Or txt_desa = "" Then
If txtnota = "" Or txtkdSupp = "" Or txtpemasok = "" Then
    MsgBox "Objek isian harus diisi.", vbExclamation, "Kesalahan"
    Exit Sub
End If
    On Error GoTo SalahSimpan
    Set RsCek = New ADODB.Recordset
    sql = "Select * From tbpembelian Where id_beli ='" & txtnota.Text & "'"
    RsCek.Open sql, Conn, adOpenKeyset, adLockReadOnly
    If RsCek.RecordCount = 0 Then
    sql = "INsert into tbpembelian values('" & txtnota & "','" & Format(DptTanggal.Value, "dd MM yyyy") & "','" & txtkdSupp & "','" & Label14.Text & "')"

End If
    Set RsSimpan = New ADODB.Recordset
    RsSimpan.Open sql, Conn, adOpenDynamic, adLockBatchOptimistic
    MsgBox "Data tersimpan ke database!", vbInformation, "Simpan Data"
    BersiPembelian
    TampilBeliKOsong
    'Bersih_Beli
    'TampilBeliKOsong
    'Call Form_Load
    'TampilDataGrid
    Exit Sub
SalahSimpan:
    MsgBox Err.Description, vbCritical, "Kesalahan SImpan"
    Exit Sub
End Sub
Sub BersiPembelian()
txtnota.Text = ""
txtkdSupp.Text = ""
txtpemasok.Text = ""
Label14.Text = ""
End Sub
Sub Bersih_Beli()
txtnota.Text = ""
txtnama.Text = ""
txtalamat.Text = ""
Label14.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub
Private Sub btnTambah_Click()
NOMOR
btnSimpan.Enabled = True
End Sub
Sub NOMOR()
Set RsCek = New ADODB.Recordset
sql = "select distinct (id_beli) from tbpembelian order by id_beli desc"
RsCek.Open sql, Conn, adOpenKeyset, adLockReadOnly
If RsCek.RecordCount = 0 Then
txtnota.Text = "B01"
Else
'Call BersihObjek
BErsihBELi
txtnota.Text = "B0" & Format(Trim(Right(RsCek!id_beli, 2)) + 1, "0")
End If
End Sub
Sub BErsihBELi()
txtnota.Text = ""
txtkdSupp.Text = ""
txtpemasok.Text = ""
End Sub
Private Sub Form_Load()
Me.Caption = "Pembelian Obat"
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 8
GrdObat.Visible = False
GrdSupplier.Visible = False
btnSimpan.Enabled = False
btnTambah.Enabled = True
btnHapus.Enabled = False
End Sub

Private Sub GrdObat_Click()
With GrdObat
If .Rows = 0 Then Exit Sub
txtkdObat.Text = .TextMatrix(.Row, 0)
txtNmObat.Text = .TextMatrix(.Row, 1)
txtharga.Text = .TextMatrix(.Row, 2)
txtJumlah.SetFocus
End With
GrdObat.Visible = False
End Sub

Private Sub ShowBUku()
Set rsGrid = New ADODB.Recordset
sql = "select tbobat.id_obat, tbobat.nm_obat, tbobat.hrg_jual from tbobat where  nm_obat  Like '%" & txtNmObat.Text & "%' order by id_obat asc"
rsGrid.Open sql, Conn, 3, 4
Set GrdObat.DataSource = rsGrid
Dim X As Integer
For X = 0 To GrdObat.Cols - 1
'Grid.FixedAlignment(x) = flexAlignCenterCenter
Next X
End Sub

Private Sub GrdSupplier_Click()
With GrdSupplier
If .Rows = 0 Then Exit Sub
txtkdSupp.Text = .TextMatrix(.Row, 0)
txtpemasok.Text = .TextMatrix(.Row, 1)
'txtharga.Text = .TextMatrix(.Row, 2)
'txtJumlah.SetFocus
End With
GrdSupplier.Visible = False
End Sub

Private Sub Grid_Click()
btnHapus.Enabled = True
With Grid
If .Rows = 0 Then Exit Sub
lblNomor.Caption = .TextMatrix(.Row, 0)
txtkdObat.Text = .TextMatrix(.Row, 1)
txtNmObat.Text = .TextMatrix(.Row, 2)
txtJumlah.Text = .TextMatrix(.Row, 3)
txtharga.Text = .TextMatrix(.Row, 4)
txtSubTotal.Text = .TextMatrix(.Row, 5)
'txtJumlah.SetFocus
End With
GrdObat.Visible = False
End Sub

Private Sub txtJumlah_Change()
txtSubTotal.Text = Val(txtJumlah.Text) * Val(txtharga.Text)
End Sub
Private Sub TampilJumlah()
Set RsCombo = New ADODB.Recordset
sql = "select sum(sub_total) as total from tbpembelian_detail where id_beli='" & txtnota.Text & "'"
RsCombo.Open sql, Conn, adOpenKeyset, adLockReadOnly
Label14.Text = ""
'cbokelompok.Clear
Do Until RsCombo.EOF
Label14.Text = RsCombo!total
RsCombo.MoveNext
Loop
End Sub
Private Sub txtNmObat_Change()
    If txtNmObat.Text = "" Then
        GrdObat.Visible = False
    Else
        GrdObat.Visible = True
        ShowBUku
    End If
End Sub

Private Sub txtpemasok_Change()
    If txtpemasok.Text = "" Then
        GrdSupplier.Visible = False
    Else
        GrdSupplier.Visible = True
        ShowSupp
    End If
End Sub

Private Sub ShowSupp()
Set rsGrid = New ADODB.Recordset
sql = "select tbsupplier.id_supplier, tbsupplier.nm_supplier, tbsupplier.almt_supplier from tbsupplier where  nm_supplier  Like '%" & txtpemasok.Text & "%' order by id_supplier asc"
rsGrid.Open sql, Conn, 3, 4
Set GrdSupplier.DataSource = rsGrid
Dim X As Integer
For X = 0 To GrdObat.Cols - 1
'Grid.FixedAlignment(x) = flexAlignCenterCenter
Next X
End Sub
