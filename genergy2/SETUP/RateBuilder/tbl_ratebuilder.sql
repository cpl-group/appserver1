USE [dbCore]
GO

ALTER TABLE [dbo].[RateBuilder] DROP CONSTRAINT [DF_RateBuilder_rowguid]
GO

/****** Object:  Table [dbo].[RateBuilder]    Script Date: 10/5/2016 1:40:08 PM ******/
DROP TABLE [dbo].[RateBuilder]
GO

/****** Object:  Table [dbo].[RateBuilder]    Script Date: 10/5/2016 1:40:08 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[RateBuilder](
	[rbid] [int] IDENTITY(1,1) NOT NULL,
	[rbcid] [int] NULL,
	[rbrid] [int] NULL,
	[rateperiod] [date] NULL,
	[createdBy] [nchar](10) NULL,
	[createdOn] [datetime] NULL,
	[modifiedBy] [nchar](10) NULL,
	[modifiedOn] [datetime] NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_RateBuilder] PRIMARY KEY CLUSTERED 
(
	[rbid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[RateBuilder] ADD  CONSTRAINT [DF_RateBuilder_rowguid]  DEFAULT (newid()) FOR [rowguid]
GO

