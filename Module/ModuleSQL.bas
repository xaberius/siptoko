Attribute VB_Name = "ModuleSQL"
Sub CekTabel()
TabelBarang 'ok
TabelJabatan 'ok
TabelJenis
TabelKonsumen
TabelPegawai
TabelSales
TabelSupplier
TabelUserX
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

