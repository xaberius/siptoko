if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DtlReturBeli]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DtlReturBeli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ReturBeli]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ReturBeli]
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

CREATE TABLE [dbo].[ReturBeli] (
	[KodeTransaksi] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TglRetur] [datetime] NULL ,
	[Penerima] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

