if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Barang]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Barang]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BeliBarang]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BeliBarang]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlBeliBarang]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlBeliBarang]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlJualEcer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlJualEcer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlJualGrosir]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlJualGrosir]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlMutasi]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlMutasi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlReturBeli]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlReturBeli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlReturJual]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlReturJual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Jabatan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Jabatan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Jenis]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Jenis]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[JualEcer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[JualEcer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[JualGrosir]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[JualGrosir]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Konsumen]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Konsumen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mutasi]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mutasi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pegawai]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Pegawai]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PerubahanHarga]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PerubahanHarga]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ReturBeli]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ReturBeli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ReturJual]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ReturJual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sales]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Supplier]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Supplier]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserX]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserX]
GO

CREATE TABLE [dbo].[Barang] (
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Jenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Satuan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[HargaBeli] [money] NOT NULL ,
	[BiayaKirim] [money] NOT NULL ,
	[HargaPokok] [money] NOT NULL ,
	[HargaGrosir] [money] NOT NULL ,
	[HargaEcer] [money] NOT NULL ,
	[StockMax] [int] NOT NULL ,
	[StockMin] [int] NOT NULL ,
	[Stock] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BeliBarang] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[KodeSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglKirim] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlBeliBarang] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL ,
	[HargaBeli] [money] NULL ,
	[Discount] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlJualEcer] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL ,
	[HargaJual] [money] NULL ,
	[BiayaKirim] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlJualGrosir] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL ,
	[HargaJual] [money] NULL ,
	[BiayaKirim] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlMutasi] (
	[TglMutasi] [datetime] NULL ,
	[NoKet] [int] NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlReturBeli] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglRetur] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL ,
	[Alasan] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlReturJual] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Alasan] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Jabatan] (
	[KodeJabatan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaJabatan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Jenis] (
	[KodeJenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaJenis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[JualEcer] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglKirim] [datetime] NULL ,
	[Discount] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[JualGrosir] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglKirim] [datetime] NULL ,
	[Discount] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Konsumen] (
	[KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Mutasi] (
	[TglMutasi] [datetime] NULL ,
	[Asal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tujuan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Pegawai] (
	[KodePegawai] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaPegawai] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Handphone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TglMasuk] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TglKeluar] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PerubahanHarga] (
	[NoTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglPerubahan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HargaLama] [money] NULL ,
	[HargaBaru] [money] NULL ,
	[Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ReturBeli] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglRetur] [datetime] NULL ,
	[Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ReturJual] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglRetur] [datetime] NULL ,
	[Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sales] (
	[KodeSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Supplier] (
	[KodeSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaSupplier] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserX] (
	[UserID] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UserName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Password] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TglEntry] [datetime] NULL ,
	[UserEntry] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglExp] [datetime] NULL 
) ON [PRIMARY]
GO

