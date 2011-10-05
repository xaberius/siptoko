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
	[StockMin] [int] NOT NULL ,
	[StockMax] [int] NOT NULL ,
	[Stock] [int] NOT NULL 
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

CREATE TABLE [dbo].[Konsumen] (
	[KodeKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaKonsumen] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepom] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
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

CREATE TABLE [dbo].[Sales] (
	[KodeSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NamaSales] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Alamat] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Kota] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telepon] [varbinary] (50) NOT NULL 
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
	[TglEntry] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Barang] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[KodeBarang]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Jabatan] WITH NOCHECK ADD 
	CONSTRAINT [PK_Jabatan] PRIMARY KEY  CLUSTERED 
	(
		[KodeJabatan]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Jenis] WITH NOCHECK ADD 
	CONSTRAINT [PK_Jenis] PRIMARY KEY  CLUSTERED 
	(
		[KodeJenis]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Konsumen] WITH NOCHECK ADD 
	CONSTRAINT [PK_Konsumen] PRIMARY KEY  CLUSTERED 
	(
		[KodeKonsumen]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Pegawai] WITH NOCHECK ADD 
	CONSTRAINT [PK_Pegawai] PRIMARY KEY  CLUSTERED 
	(
		[KodePegawai]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Sales] WITH NOCHECK ADD 
	CONSTRAINT [PK_Sales] PRIMARY KEY  CLUSTERED 
	(
		[KodeSales]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Supplier] WITH NOCHECK ADD 
	CONSTRAINT [PK_Supplier] PRIMARY KEY  CLUSTERED 
	(
		[KodeSupplier]
	)  ON [PRIMARY] 
GO

