VERSION 5.00
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "APLIKASI PENCATATAN OBAT PADA KLINIK CINTA KASIH"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9930
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMenu.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnObat 
         Caption         =   "&Obat"
      End
      Begin VB.Menu mnsupplier 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mnJenis 
         Caption         =   "&Jenis"
      End
   End
   Begin VB.Menu mnTransaksi 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnPenjualan 
         Caption         =   "&Penjualan Obat"
      End
      Begin VB.Menu mnPembelian 
         Caption         =   "Pem&belian Obat"
      End
   End
   Begin VB.Menu mnLaporan3 
      Caption         =   "&Laporan"
      Begin VB.Menu mnLaporan 
         Caption         =   "Laporan &Stok Obat"
      End
      Begin VB.Menu mnPenjualan2 
         Caption         =   "Laporan Penjualan"
      End
      Begin VB.Menu mnPembelian2 
         Caption         =   "Laporan Pembelian"
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Apakah Anda Yakin Ingin Keluar??", vbYesNo, "Keluar?") = vbYes Then
End
End If
End Sub

Private Sub mnExit_Click()
If MsgBox("Apakah Anda Yakin Ingin Keluar??", vbYesNo, "Keluar?") = vbYes Then
End
End If
End Sub

Private Sub mnJenis_Click()
frmJenis.Show
End Sub

Private Sub mnLaporan_Click()
frmLaporanStok.Show
End Sub

Private Sub mnObat_Click()
frmObat.Show
End Sub

Private Sub mnPembelian_Click()
frmPembelian.Show
End Sub

Private Sub mnPembelian2_Click()
frmLapPembelian.Show
End Sub

Private Sub mnpenjualan_Click()
frmJual.Show
End Sub

Private Sub mnPenjualan2_Click()
frmLaPenjualan.Show
End Sub

Private Sub mnsupplier_Click()
frmSupplier.Show
End Sub

