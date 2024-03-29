if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BeliBarang]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BeliBarang]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlBeliBarang]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlBeliBarang]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlMutasi]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlMutasi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mutasi]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mutasi]
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
	[Jumlah] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HargaBeli] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BiayaKirim] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DtlMutasi] (
	[TglMutasi] [datetime] NULL ,
	[NoKet] [int] NULL ,
	[KodeBarang] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Jumlah] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Mutasi] (
	[TglMutasi] [datetime] NULL ,
	[Asal] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tujuan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Keterangan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

