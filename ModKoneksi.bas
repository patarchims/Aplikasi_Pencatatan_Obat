Attribute VB_Name = "ModKoneksi"
Public Conn As ADODB.Connection
Public Conn1 As ADODB.Connection
Public sql As String
Public SQL2, SQL1 As String
Public Csql As String
Public Pemakai As String
'Public sqlcommand As NewADODB.Recordset
Dim RsLogin As New ADODB.Recordset
Public rsGrid As ADODB.Recordset
Public rsTampil As ADODB.Recordset
Public RsCek As ADODB.Recordset
Public RsCek1, RsCek2 As ADODB.Recordset
Public RsCombo As ADODB.Recordset
Public RsSimpan, RsSimpan1, RsSimpan2 As ADODB.Recordset
Public RsHapus As ADODB.Recordset
Public db As New ADODB.Connection
Public bpjs As New ADODB.Recordset
Dim constr As String
Public Rs As New ADODB.Recordset
Public Strconn As String
Public strsql As String
Public no As Integer
Sub KONEKSI()
Strconn = "Provider = SQLOLEDB.1;Integrated Security = SSPI;Persist Security Info=False;Initial Catalog=dbklinik"
Conn.CursorLocation = adUseClient
If Conn.State = adStateClosed Then
Conn.Open Strconn
End If
End Sub

Public Sub main()
Set Conn = New ADODB.Connection
KONEKSI
frmLogin.Show
End Sub


