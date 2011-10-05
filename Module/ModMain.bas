Attribute VB_Name = "Module1"

'---------------------------------------------------------------------------------------
' Module    : Module1
' DateTime  : 2/25/2011 09:25
' Author    : Agoes Said
' Purpose   :
'---------------------------------------------------------------------------------------

Private Const MAX_PATH = 260

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim strWindowsSystemDirectory           As String


Type usr
   UserId            As String
End Type

Type Login
     LogNama         As String
     LogPass         As String
     LogDB           As String
End Type

Type ProfileUser
     gAwalThnFiskal  As Date
     gAkhirThnFiskal As Date
     gAwalBlnFiskal  As Date
     gAkhirBlnFiskal As Date
     gNamaProfil     As String
     AlamatProfil1   As String
     AlamatProfil2   As String
     KotaProfil      As String
     gFiskalBi       As Integer
     TeleponProfil   As String
     FaxProfil       As String
     Npwp            As String
End Type


Public DbConP               As New ADODB.Connection
Public DbCon                As New ADODB.Connection
Public RSFind               As New ADODB.Recordset
Public StrCon               As New ADODB.Connection
Public RsTmp                As New ADODB.Recordset
Public Log                  As Login
Public FileName             As String
Public Trans                As New Convert
Public Profile              As ProfileUser
Public User                 As usr
Public ConDB, Dbs           As String
Public passuser             As String
Public KatAsuransi          As String
Dim strBuffer               As String
Dim lngReturn               As Long
Dim SQL                     As String
Public Connect1             As Boolean
Public Connect2             As Boolean

Function Nul(kode As Variant, Optional data)
      If IsMissing(data) Then _
         data = ""
         Nul = IIf(IsNull(kode) Or kode = "", data, kode)
End Function

Function UpCase(Key As Integer) As Integer
  UpCase = Asc(UCase(Chr(Key)))
End Function

Function k2t(ByVal Value As String) As String
  k2t = Replace(Value, ",", ".")
End Function

Function FormatTgl(ddate As Date) As String
    FormatTgl = Format(ddate, "mm/dd/yyyy")
End Function

Function TglAkhirBulan(Period As Integer, Anydate As Variant) As Variant
   Dim TglAwalBDepan As Variant
   On Error GoTo vb_error
    TglAwalBDepan = DateSerial(Year(Anydate), Month(Anydate) + Period + 1, 1)
    TglAkhirBulan = DateAdd("d", -1, TglAwalBDepan)
   Exit Function
vb_error:
    MsgBox ErrMessage(Erl, Err.Number, "Procedure : ModMain.TglAkhirBulan")
End Function

Sub Main()
10        HariSvr = Format(DateSvr, "dd")
20        Dbs = "SIP-Toko"
30        strBuffer = Space$(MAX_PATH)
40        lngReturn = GetSystemDirectory(strBuffer, MAX_PATH)
50        strWindowsSystemDirectory = Left$(strBuffer, Len(strBuffer) - 1)
60        PathWindows = Trim(strWindowsSystemDirectory)
70        FileName = Left(PathWindows, Len(PathWindows) - 1) & "\" & "Agus.Said"

80        Call GetLogin
          Ulogin$ = Trim(Trans.decryp_pass(21, Log.LogNama))  'Ambil UserLogin SQL
          UPass$ = Trim(Trans.decryp_pass(21, Log.LogPass))  'Ambil PassLogin SQL
90        'Ulogin$ = "sa"
100       'UPass$ = "matahari"
110       If Get_Connection(Dbs, Ulogin$, UPass$) Then
             CekTabel
120          FrmUtama.Show
             Connect1 = True
130       Else
140          MsgBox "Koneksi ke database Gagal, Silahkan Hubungi Administrator/IT", vbCritical + vbMsgBoxRight
    End If

End Sub

Public Sub GetLogin()
  On Error GoTo vb_error
  'Dim nama As String
  Open FileName For Input As #1
  Do Until EOF(1)
  Line Input #1, nama
     a = a + 1
     'nama = Names
     Select Case a
        Case 1: Log.LogNama = nama
        Case 2: Log.LogPass = nama
        Case 3: Log.LogDB = nama
     End Select
     If a = 3 Then Exit Do
  Loop
  Close #1
vb_error:
End Sub

Public Function Get_User(UsrName As String, _
                         pass As String) As Boolean
         Dim RsUser As New ADODB.Recordset
1        On Error GoTo vb_error
2        SQL = "SELECT user_password from [User]" & _
                     "WHERE user_id ='" & UsrName & "' "
3        Set RsUser = DbCon.Execute(SQL)
4        If Not RsUser.BOF Then XX = Trim(Trans.decryp_pass(21, RsUser!user_password))
5        If Not RsUser.BOF And XX = pass Then
6           passuser = Trim(Trans.decryp_pass(21, RsUser!user_password))
7           Get_User = True
8           Ulogin$ = Trim(Trans.decryp_pass(21, Log.LogNama)) 'Ambil UserLogin SQL
9           UPass$ = Trim(Trans.decryp_pass(21, Log.LogPass)) 'Ambil PassLogin SQL
10          User.UserId = UsrName
11       Else
12          Get_User = False
13       End If
14       On Error GoTo 0
15       Exit Function
vb_error:
16        MsgBox ErrMessage(Erl, Err.Number, "Procedure : ModMain.Get_User"), vbExclamation, "Err Number : " & Erl
End Function

Public Function ErrMessage(ByVal Errline As Long, _
                           ByVal ErrNumber As Long, _
                           FunctionName As String) As String
    ErrMessage = "Error line Number = " & Errline & vbCrLf & _
                 "Error Number      = " & ErrNumber & vbCrLf & _
                 "Error Description = " & Error$(ErrNumber) & vbCrLf & _
                 "Location Error    = " & FunctionName
End Function

Public Function Get_Connection(DbName As String, _
                               UsrName As String, _
                               pass As String) As Boolean
1      On Error GoTo vb_error
2      App.Title = "SIP-Toko"
3      Get_Connection = True
4     With DbCon
5      .CursorLocation = adUseClient
6      .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & pass & ";Initial Catalog=" & DbName & ";server=" & GetSetting(App.Title, "startup", "server", "(local)")
7      adocon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & pass & ";Initial Catalog=[" & DbName & "];server=" & GetSetting(App.Title, "startup", "server", "(local)")
8      ConDB = .ConnectionString
9      .Open
10    End With

11     Exit Function
vb_error:
12    Get_Connection = False
13    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_Connection"), vbCritical + vbMsgBoxRight, "Open Connection"
End Function

Sub Set_Tombol(frm As Form, ByRef obj As Object, _
              Jarak As Integer, Lbr As Integer, Tags As String, _
              Kapsion As String, Visib As Boolean, Aktif As Boolean)
With obj
      Set .Container = frm.FrmTombol
        .Caption = Kapsion
        .ColorScheme = Custom
        .Width = Lbr
        .Height = 675
        If Lbr < 1000 Then
           If Jarak = 1 Then
             .Left = 60 + (Lbr / 2)
           Else
             .Left = frm.FrmTombol.Width - (frm.FrmTombol.Width - (Jarak * Lbr)) + 550
           End If
        Else
          .Left = 250 * (Jarak + 1)
        End If
        On Error Resume Next
        .PictureNormal = FrmSu_Profil.ImageList1.ListImages(k2t(LCase(.Name))).Picture
        If Len(Mid(.Name, 4, Len(.Name))) <= 8 And .Name <> "CmdNo" _
           And .Name <> "CmdFirst" And .Name <> "CmdPrevious" And .Name <> "CmdNext" And .Name <> "CmdLast" And Kapsion = "" Then
            .Caption = "&" & Mid(.Name, 4, Len(.Name))
        Else
            .Caption = Kapsion
        End If
        .ToolTipText = Tags
        .MaskColor = &HFFFFFF
        .Top = 50
        .Tag = Tags
        .Visible = Visib
        .Enabled = Aktif
        .ZOrder 0
End With
End Sub

Public Sub SetIcon(frm As Form, _
                   Optional navigator As Boolean = True)
1     On Error Resume Next
2      With Menu_Utama.ImageList1
3       frm.CmdAdd.PictureNormal = .ListImages("tambah").Picture
4       frm.CmdAdd.ToolTipText = "Anda ingin melakukan tambah data baru Klik Disini"
5       frm.CmdAdd.Tag = "Tambah Record"
6       frm.CmdAdd.Caption = "&Tambah"
7       frm.CmdAdd.Height = 810: frm.CmdAdd.Top = 200
8       frm.CmdEdit.PictureNormal = .ListImages("edit").Picture
9       frm.CmdEdit.ToolTipText = "Anda ingin melakukan edit data >> Klik Disini"
10      frm.CmdEdit.Tag = "Edit"
11      frm.CmdEdit.Caption = "&Ubah"
12      frm.CmdEdit.Height = 810: frm.CmdEdit.Top = 200
13      frm.CmdDelete.PictureNormal = .ListImages("delete").Picture
14      frm.CmdDelete.ToolTipText = "Anda ingin menghapus data >> Klik Disini"
15      frm.CmdDelete.Tag = "Hapus"
16      frm.CmdDelete.Caption = "&Hapus"
17      frm.CmdDelete.Height = 810: frm.CmdDelete.Top = 200
18      frm.CmdSave.PictureNormal = .ListImages("save").Picture
19      frm.CmdSave.ToolTipText = "Anda ingin menyimpan data >> Klik Disini"
20      frm.CmdSave.Tag = "Simpan"
21      frm.CmdSave.Caption = "&Simpan"
22      frm.CmdSave.Height = 810: frm.CmdSave.Top = 200
23      frm.CmdCancel.PictureNormal = .ListImages("cancel").Picture
24      frm.CmdCancel.ToolTipText = "Anda ingin membatalkan proses data >> Klik Disini"
25      frm.CmdCancel.Tag = "Batal"
26      frm.CmdCancel.Caption = "&Batal"
27      frm.CmdCancel.Height = 810: frm.CmdCancel.Top = 200
28      frm.CmdFind.PictureNormal = .ListImages("find").Picture
29      frm.CmdFind.ToolTipText = "Anda ingin mencari data >> Klik Disini"
30      frm.CmdFind.Tag = "Cari"
31      frm.CmdFind.Caption = "&Cari"
32      frm.CmdQuit.PictureNormal = .ListImages("cmdquit").Picture
33      frm.CmdQuit.ToolTipText = "Anda ingin keluar dari proses >> Klik Disini"
34      frm.CmdQuit.Tag = "Keluar"
35      frm.CmdQuit.Caption = "K&eluar"
36      frm.CmdQuit.Height = 810: frm.CmdQuit.Top = 200
38      If navigator = True Then
39         frm.CmdLast.PictureNormal = .ListImages("cmdlast").Picture
40         frm.CmdLast.ToolTipText = ""
41         frm.CmdLast.Tag = "Mundur Ke Akhir Record"
42         frm.CmdFirst.PictureNormal = .ListImages("cmdfirst").Picture
43         frm.CmdFirst.ToolTipText = ""
44         frm.CmdFirst.Tag = "Maju Ke Awal Record"
45         frm.CmdPrevious.PictureNormal = .ListImages("cmdprevious").Picture
46         frm.CmdPrevious.ToolTipText = ""
47         frm.CmdPrevious.Tag = "Mundur Satu Record"
48         frm.CmdNext.PictureNormal = .ListImages("cmdnext").Picture
49         frm.CmdNext.ToolTipText = ""
50         frm.CmdNext.Tag = "Maju Satu Record"
51      End If
52         frm.CmdPrint.ToolTipText = ""
53         frm.CmdPrint.PictureNormal = .ListImages("cmdprint").Picture
54         frm.CmdPrint.Tag = "Cetak Ke Bentuk Laporan"
55         frm.CmdPrint.Caption = "&Print"
56         frmcmdno.Tag = "Non Aktif No. Urut"
57    End With
58    Exit Sub
vb_error:
59      MsgBox ErrMessage(Erl, Err.Number, "Procedure : SetIcon"), vbExclamation, "Err Number : " & Erl
End Sub

Public Sub Enter(ByVal Key As Integer, Optional ByRef XX As Object)
     If Key = 13 Then SendKeys "{tab}"
     If Key = 38 Then
         If XX Is Nothing Then
            SendKeys "+{tab}"
            Exit Sub
         End If
     ElseIf Key = 40 Then
         If XX Is Nothing Then
            SendKeys "{tab}"
            Exit Sub
         End If
     End If
End Sub
Function TglNull(kode As Variant) As Variant
If IsNull(kode) Or kode = "" Then
    TglNull = "Null"
Else
    TglNull = "'" & FormatTgl(CDate(kode)) & "'"
End If
End Function





Public Sub BukaDB()

Dim strString As String

    strString = "provider = Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & "/database/tmp.mdb;" & _
    "Persist Security Info=False; "

    Set StrCon = New ADODB.Connection
    StrCon.Open strString
    StrCon.CursorLocation = adUseClient


   'membuka koneksi
    RsTmp.LockType = adLockOptimistic
    RsTmp.Open "tmp_spk", StrCon, adOpenDynamic, adLockOptimistic

End Sub

Sub LogSistem(NamaForm As String, Tombol As String, Keterangan As String)
SQL = "insert into LogSistem values('" & User.UserId & "',getdate(),'" & NamaForm & _
    "','" & Tombol & "','" & Keterangan & "')"
DbCon.Execute (SQL)
End Sub

Sub MenuTombol(UserName As String, NamaForm As String, Tombol As String)
SQL = "insert into MenuTombol values('" & UserName & "','" & NamaForm & "','" & Tombol & "')"
DbCon.Execute (SQL)
End Sub

Function CekKonek(Form As Form)
If Not Connect1 Then
    MsgBox "Anda Belum Konek Ke SQL Server"
    FrmLoginServer.Show
    Unload Form
ElseIf Not Connect2 Then
    MsgBox "Anda Belum Login Aplikasi"
    FrmLogin.Show
    Unload Form
Else
    Form.Show , FrmUtama
End If
End Function

