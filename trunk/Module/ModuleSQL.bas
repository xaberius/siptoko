Attribute VB_Name = "ModuleSQL"
Sub CekTabel()
TabelBarang 'ok
TabelJabatan 'ok
TabelJenis 'ok
TabelKonsumen 'ok
TabelPegawai 'ok
TabelSales 'ok
TabelSupplier 'ok
TabelUserX 'ok
TabelMutasi 'ok
TabelDtlMutasi 'ok
TabelBeliBarang 'ok
TabelDtlBeliBarang 'ok
TabelReturBeli 'ok
TabelDtlReturBeli 'ok
TabelJualGrosir 'ok
TabelDtlJualGrosir 'ok
TabelJualEcer 'ok
TabelDtlJualEcer 'ok
TabelReturJual 'ok
TabelDtlReturJual 'ok
TabelPerubahanHarga 'ok
End Sub

Sub TabelBarang()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Barang') "
SQL = SQL + " CREATE TABLE [dbo].[Barang] ( "
SQL = SQL + " [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + " [NamaBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + " [Jenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + " [Satuan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + " [HargaBeli] [money] NOT NULL ,"
SQL = SQL + " [BiayaKirim] [money] NOT NULL ,"
SQL = SQL + " [HargaPokok] [money] NOT NULL ,"
SQL = SQL + " [HargaGrosir] [money] NOT NULL ,"
SQL = SQL + " [HargaEcer] [money] NOT NULL ,"
SQL = SQL + " [StockMin] [int] NOT NULL ,"
SQL = SQL + " [StockMax] [int] NOT NULL ,"
SQL = SQL + " [Stock] [Int] NOT NULL"
SQL = SQL + " Primary Key(KodeBarang)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelJabatan()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Jabatan') "
SQL = SQL + " CREATE TABLE [dbo].[Jabatan] ("
SQL = SQL + "     [KodeJabatan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaJabatan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodeJabatan)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelJenis()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Jenis') "
SQL = SQL + " CREATE TABLE [dbo].[Jenis] ("
SQL = SQL + "     [KodeJenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaJenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodeJenis)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelKonsumen()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Konsumen') "
SQL = SQL + " CREATE TABLE [dbo].[Konsumen] ("
SQL = SQL + "     [KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodeKonsumen)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub


Sub TabelPegawai()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Pegawai') "
SQL = SQL + " CREATE TABLE [dbo].[Pegawai] ("
SQL = SQL + "     [KodePegawai] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaPegawai] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Handphone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [TglMasuk] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [TglKeluar] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodePegawai)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelSales()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Sales') "
SQL = SQL + " CREATE TABLE [dbo].[Sales] ("
SQL = SQL + "     [KodeSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Telepon]  [VarChar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodeSales)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelSupplier()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Supplier') "
SQL = SQL + " CREATE TABLE [dbo].[Supplier] ("
SQL = SQL + "     [KodeSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [NamaSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL"
SQL = SQL + " primary key (KodeSupplier)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelUserX()
SQL = "if not exists(select * from dbo.sysobjects where name = 'UserX') "
SQL = SQL + " CREATE TABLE [dbo].[UserX] ("
SQL = SQL + "     [UserID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [UserName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [Password] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [TglEntry] [datetime] NULL,"
SQL = SQL + "     [UserEntry] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
SQL = SQL + "     [TglExp] [datetime] NULL"
SQL = SQL + " primary key (UserID)"
SQL = SQL + " ) ON [PRIMARY]"
DbCon.Execute SQL

SQL = "select * from UserX where userID='admin'"
Set RSFind = DbCon.Execute(SQL)
If RSFind.BOF Or RSFind.RecordCount = 0 Then
    SQL = "insert into UserX values('Admin','Admin','" & Trans.encryp_pass(25, "admin") & _
        "','" & FormatTgl(Date) & "', 'Admin', '" & FormatTgl(Date + 60) & "') "
    DbCon.Execute SQL
End If
End Sub

Sub TabelMutasi()
SQL = "if not exists(select * from dbo.sysobjects where name = 'Mutasi') "
SQL = SQL + "CREATE TABLE [dbo].[Mutasi] ("
SQL = SQL + "    [TglMutasi] [datetime] NULL ,"
SQL = SQL + "    [Asal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Tujuan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelDtlMutasi()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlMutasi') "
SQL = SQL + "CREATE TABLE [dbo].[DtlMutasi] ("
SQL = SQL + "    [TglMutasi] [datetime] NULL ,"
SQL = SQL + "    [NoKet] [int] NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [int] NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelBeliBarang()
SQL = "if not exists(select * from dbo.sysobjects where name = 'BeliBarang') "
SQL = SQL + "CREATE TABLE [dbo].[BeliBarang] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [KodeSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglKirim] [datetime] NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelDtlBeliBarang()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlBeliBarang') "
SQL = SQL + "CREATE TABLE [dbo].[DtlBeliBarang] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [HargaBeli] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [BiayaKirim] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelDtlReturBeli()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlReturBeli') "
SQL = SQL + "CREATE TABLE [dbo].[DtlReturBeli] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglRetur] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [int] NULL ,"
SQL = SQL + "    [Alasan] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub

Sub TabelReturBeli()
SQL = "if not exists(select * from dbo.sysobjects where name = 'ReturBeli') "
SQL = SQL + "CREATE TABLE [dbo].[ReturBeli] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglRetur] [datetime] NULL ,"
SQL = SQL + "    [Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
DbCon.Execute SQL
End Sub


Sub TabelDtlJualEcer()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlJualEcer') "
SQL = SQL + "CREATE TABLE [dbo].[DtlJualEcer] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [int] NULL ,"
SQL = SQL + "    [HargaJual] [money] NULL ,"
SQL = SQL + "    [BiayaKirim] [money] NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelDtlJualGrosir()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlJualGrosir') "
SQL = SQL + "CREATE TABLE [dbo].[DtlJualGrosir] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [int] NULL ,"
SQL = SQL + "    [HargaJual] [money] NULL ,"
SQL = SQL + "    [BiayaKirim] [money] NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelDtlReturJual()
SQL = "if not exists(select * from dbo.sysobjects where name = 'DtlReturJual') "
SQL = SQL + "CREATE TABLE [dbo].[DtlReturJual] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Jumlah] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Alasan] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelJualEcer()
SQL = "if not exists(select * from dbo.sysobjects where name = 'JualEcer') "
SQL = SQL + "CREATE TABLE [dbo].[JualEcer] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglKirim] [datetime] NULL ,"
SQL = SQL + "    [Discount] [money] NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelJualGrosir()
SQL = "if not exists(select * from dbo.sysobjects where name = 'JualGrosir') "
SQL = SQL + "CREATE TABLE [dbo].[JualGrosir] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglTransaksi] [datetime] NULL ,"
SQL = SQL + "    [KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglKirim] [datetime] NULL ,"
SQL = SQL + "    [Discount] [money] NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelReturJual()
SQL = "if not exists(select * from dbo.sysobjects where name = 'ReturJual') "
SQL = SQL + "CREATE TABLE [dbo].[ReturJual] ("
SQL = SQL + "    [KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglRetur] [datetime] NULL ,"
SQL = SQL + "    [Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub

Sub TabelPerubahanHarga()
SQL = "if not exists(select * from dbo.sysobjects where name = 'PerubahanHarga') "
SQL = SQL + "CREATE TABLE [dbo].[PerubahanHarga] ("
SQL = SQL + "    [NoTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [TglPerubahan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
SQL = SQL + "    [HargaLama] [money] NULL ,"
SQL = SQL + "    [HargaBaru] [money] NULL ,"
SQL = SQL + "    [Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
SQL = SQL + ") ON [PRIMARY]"
End Sub
