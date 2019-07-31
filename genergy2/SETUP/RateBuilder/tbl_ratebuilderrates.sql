USE [dbCore]
GO

ALTER TABLE [dbo].[RateBuilderRates] DROP CONSTRAINT [DF_RateBuilderRates_rowguid]
GO

/****** Object:  Table [dbo].[RateBuilderRates]    Script Date: 10/5/2016 1:40:36 PM ******/
DROP TABLE [dbo].[RateBuilderRates]
GO

/****** Object:  Table [dbo].[RateBuilderRates]    Script Date: 10/5/2016 1:40:36 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[RateBuilderRates](
	[rbrid] [int] IDENTITY(1,1) NOT NULL,
	[sc9r1_e_er] [nchar](10) NULL,
	[sc9r1_e_macadj] [nchar](10) NULL,
	[sc9r1_d_mscadj] [nchar](10) NULL,
	[sc9r1_d_dr_5] [nchar](10) NULL,
	[sc9r1_d_dr_100] [nchar](10) NULL,
	[sc9r1_d_dr_999] [nchar](10) NULL,
	[conedsc9r1_s_bppc] [nchar](10) NULL,
	[conedsc9r1_s_cmc] [nchar](10) NULL,
	[conedsc9r1_e_er] [nchar](10) NULL,
	[conedsc9r1_e_macadj] [nchar](10) NULL,
	[conedsc9r1_e_mfc] [nchar](10) NULL,
	[conedsc9r1_d_mscadj] [nchar](10) NULL,
	[conedsc9r1_d_dr_5] [nchar](10) NULL,
	[conedsc9r1_d_dr_999] [nchar](10) NULL,
	[sc9r2_e_er] [nchar](10) NULL,
	[sc9r2_e_macadj] [nchar](10) NULL,
	[sc9r2_d_mscadj] [nchar](10) NULL,
	[sc9r2_d_dr_1800] [nchar](10) NULL,
	[sc9r2_d_dr_2200] [nchar](10) NULL,
	[sc9r2_d_dr_2359] [nchar](10) NULL,
	[sc9ra1_e_er] [nchar](10) NULL,
	[sc9ra1_e_macadj] [nchar](10) NULL,
	[sc9ra1_d_dr_5] [nchar](10) NULL,
	[sc9ra1_d_dr_100] [nchar](10) NULL,
	[sc9ra1_d_dr_999] [nchar](10) NULL,
	[sc9ra2_e_er] [nchar](10) NULL,
	[sc9ra2_e_macadj] [nchar](10) NULL,
	[sc9ra2_d_dr_1800] [nchar](10) NULL,
	[sc9ra2_d_dr_2200] [nchar](10) NULL,
	[sc9ra2_d_dr_2359] [nchar](10) NULL,
	[sc9ra3_e_er] [nchar](10) NULL,
	[sc9ra3_e_macadj] [nchar](10) NULL,
	[sc9ra3_d_dr_1800] [nchar](10) NULL,
	[sc9ra3_d_dr_2200] [nchar](10) NULL,
	[sc9ra3_d_dr_2359] [nchar](10) NULL,
	[sc12ra2_e_er] [nchar](10) NULL,
	[sc12ra2_e_macadj] [nchar](10) NULL,
	[sc12ra2_d_dr_1800] [nchar](10) NULL,
	[sc12ra2_d_dr_2200] [nchar](10) NULL,
	[sc12ra2_d_dr_2359] [nchar](10) NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_RateBuilderRates] PRIMARY KEY CLUSTERED 
(
	[rbrid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[RateBuilderRates] ADD  CONSTRAINT [DF_RateBuilderRates_rowguid]  DEFAULT (newid()) FOR [rowguid]
GO

