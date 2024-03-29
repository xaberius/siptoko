if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PerubahanHarga]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PerubahanHarga]
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

