if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlJualEcer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlJualEcer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlJualGrosir]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlJualGrosir]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlReturJual]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlReturJual]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[JualEcer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[JualEcer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[JualGrosir]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[JualGrosir]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ReturJual]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ReturJual]
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

CREATE TABLE [dbo].[DtlReturJual] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglTransaksi] [datetime] NULL ,
	[NoKet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Alasan] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
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

CREATE TABLE [dbo].[ReturJual] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglRetur] [datetime] NULL ,
	[Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

