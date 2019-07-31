USE [dbCore]
GO

ALTER TABLE [dbo].[RateBuilderRates] DROP CONSTRAINT [DF_RateBuilderRates_rowguid]
GO

/****** Object:  Table [dbo].[RateBuilderRates]    Script Date: 11/15/2016 4:27:20 PM ******/
DROP TABLE [dbo].[RateBuilderRates]
GO

/****** Object:  Table [dbo].[RateBuilderRates]    Script Date: 11/14/2016 3:07:17 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[RateBuilderRates](
	[rbrid] [int] IDENTITY(1,1) NOT NULL,
	[sc9r1_e_er] [nchar](10) NULL,
	[sc9r1_e_er_id] [int] NULL,
	[sc9r1_e_macadj] [nchar](10) NULL,
	[sc9r1_e_macadj_id] [int] NULL,
	[sc9r1_d_mscadj] [nchar](10) NULL,
	[sc9r1_d_mscadj_id] [int] NULL,
	[sc9r1_d_dr_l5] [nchar](10) NULL,
	[sc9r1_d_dr_l5_id] [int] NULL,
	[sc9r1_d_dr_l100] [nchar](10) NULL,
	[sc9r1_d_dr_l100_id] [int] NULL,
	[sc9r1_d_dr_l999] [nchar](10) NULL,
	[sc9r1_d_dr_l999_id] [int] NULL,
	[conedsc9r1_s_bppc] [nchar](10) NULL,
	[conedsc9r1_s_bppc_id] [int] NULL,
	[conedsc9r1_s_cmc] [nchar](10) NULL,
	[conedsc9r1_s_cmc_id] [int] NULL,
	[conedsc9r1_e_er] [nchar](10) NULL,
	[conedsc9r1_e_er_id] [int] NULL,
	[conedsc9r1_e_macadj] [nchar](10) NULL,
	[conedsc9r1_e_macadj_id] [int] NULL,
	[conedsc9r1_e_mfc] [nchar](10) NULL,
	[conedsc9r1_e_mfc_id] [int] NULL,
	[conedsc9r1_d_mscadj] [nchar](10) NULL,
	[conedsc9r1_d_mscadj_id] [int] NULL,
	[conedsc9r1_d_dr_l5] [nchar](10) NULL,
	[conedsc9r1_d_dr_l5_id] [int] NULL,
	[conedsc9r1_d_dr_l999] [nchar](10) NULL,
	[conedsc9r1_d_dr_l999_id] [int] NULL,
	[sc9r2_e_er_p759]  [nchar](10) NULL,
	[sc9r2_e_er_p759_id]  [nchar](10) NULL,
	[sc9r2_e_er_p1800] [nchar](10) NULL,
	[sc9r2_e_er_p1800_id] [int] NULL,
	[sc9r2_e_er_p2200] [nchar](10) NULL,
	[sc9r2_e_er_p2200_id] [int] NULL,
	[sc9r2_e_er_p2359] [nchar](10) NULL,
	[sc9r2_e_er_p2359_id] [int] NULL,
	[sc9r2_e_macadj_p759]  [nchar](10) NULL,
	[sc9r2_e_macadj_p759_id]  [nchar](10) NULL,
	[sc9r2_e_macadj_p1800] [nchar](10) NULL,
	[sc9r2_e_macadj_p1800_id] [int] NULL,
	[sc9r2_e_macadj_p2200] [nchar](10) NULL,
	[sc9r2_e_macadj_p2200_id] [int] NULL,
	[sc9r2_e_macadj_p2359] [nchar](10) NULL,
	[sc9r2_e_macadj_p2359_id] [int] NULL,
	[sc9r2_d_mscadj_p2200] [nchar](10) NULL,
	[sc9r2_d_mscadj_p2200_id] [int] NULL,
	[sc9r2_d_dr_p1800] [nchar](10) NULL,
	[sc9r2_d_dr_p1800_id] [int] NULL,
	[sc9r2_d_dr_p2200] [nchar](10) NULL,
	[sc9r2_d_dr_p2200_id] [int] NULL,
	[sc9r2_d_dr_p2359] [nchar](10) NULL,
	[sc9r2_d_dr_p2359_id] [int] NULL,
	[sc9ra1_e_er] [nchar](10) NULL,
	[sc9ra1_e_er_id] [int] NULL,
	[sc9ra1_e_macadj] [nchar](10) NULL,
	[sc9ra1_e_macadj_id] [int] NULL,
	[sc9ra1_d_dr_l5] [nchar](10) NULL,
	[sc9ra1_d_dr_l5_id] [int] NULL,
	[sc9ra1_d_dr_l100] [nchar](10) NULL,
	[sc9ra1_d_dr_l100_id] [int] NULL,
	[sc9ra1_d_dr_l999] [nchar](10) NULL,
	[sc9ra1_d_dr_l999_id] [int] NULL,
	[sc9ra2_e_er_p759] [nchar](10) NULL,
	[sc9ra2_e_er_p759_id] [int] NULL,
	[sc9ra2_e_er_p1800] [nchar](10) NULL,
	[sc9ra2_e_er_p1800_id] [int] NULL,
	[sc9ra2_e_er_p2200] [nchar](10) NULL,
	[sc9ra2_e_er_p2200_id] [int] NULL,
	[sc9ra2_e_er_p2359] [nchar](10) NULL,
	[sc9ra2_e_er_p2359_id] [int] NULL,
	[sc9ra2_e_macadj_p759]  [nchar](10) NULL,
	[sc9ra2_e_macadj_p759_id]  [nchar](10) NULL,
	[sc9ra2_e_macadj_p1800] [nchar](10) NULL,
	[sc9ra2_e_macadj_p1800_id] [int] NULL,
	[sc9ra2_e_macadj_p2200] [nchar](10) NULL,
	[sc9ra2_e_macadj_p2200_id] [int] NULL,
	[sc9ra2_e_macadj_p2359] [nchar](10) NULL,
	[sc9ra2_e_macadj_p2359_id] [int] NULL,
	[sc9ra2_d_dr_p1800] [nchar](10) NULL,
	[sc9ra2_d_dr_p1800_id] [int] NULL,
	[sc9ra2_d_dr_p2200] [nchar](10) NULL,
	[sc9ra2_d_dr_p2200_id] [int] NULL,
	[sc9ra2_d_dr_p2359] [nchar](10) NULL,
	[sc9ra2_d_dr_p2359_id] [int] NULL,
	[sc9ra3_e_er_p759] [nchar](10) NULL,
	[sc9ra3_e_er_p759_id] [int] NULL,
	[sc9ra3_e_er_p1800] [nchar](10) NULL,
	[sc9ra3_e_er_p1800_id] [int] NULL,
	[sc9ra3_e_er_p2200] [nchar](10) NULL,
	[sc9ra3_e_er_p2200_id] [int] NULL,
	[sc9ra3_e_er_p2359] [nchar](10) NULL,
	[sc9ra3_e_er_p2359_id] [int] NULL,
	[sc9ra3_e_macadj_p759]  [nchar](10) NULL,
	[sc9ra3_e_macadj_p759_id]  [nchar](10) NULL,
	[sc9ra3_e_macadj_p1800] [nchar](10) NULL,
	[sc9ra3_e_macadj_p1800_id] [int] NULL,
	[sc9ra3_e_macadj_p2200] [nchar](10) NULL,
	[sc9ra3_e_macadj_p2200_id] [int] NULL,
	[sc9ra3_e_macadj_p2359] [nchar](10) NULL,
	[sc9ra3_e_macadj_p2359_id] [int] NULL,
	[sc9ra3_d_dr_p1800] [nchar](10) NULL,
	[sc9ra3_d_dr_p1800_id] [int] NULL,
	[sc9ra3_d_dr_p2200] [nchar](10) NULL,
	[sc9ra3_d_dr_p2200_id] [int] NULL,
	[sc9ra3_d_dr_p2359] [nchar](10) NULL,
	[sc9ra3_d_dr_p2359_id] [int] NULL,
	[sc12ra2_e_er_p759] [nchar](10) NULL,
	[sc12ra2_e_er_p759_id] [int] NULL,
	[sc12ra2_e_er_p1800] [nchar](10) NULL,
	[sc12ra2_e_er_p1800_id] [int] NULL,
	[sc12ra2_e_er_p2200] [nchar](10) NULL,
	[sc12ra2_e_er_p2200_id] [int] NULL,
	[sc12ra2_e_er_p2359] [nchar](10) NULL,
	[sc12ra2_e_er_p2359_id] [int] NULL,
	[sc12ra2_e_macadj_p759]  [nchar](10) NULL,
	[sc12ra2_e_macadj_p759_id]  [nchar](10) NULL,
	[sc12ra2_e_macadj_p1800] [nchar](10) NULL,
	[sc12ra2_e_macadj_p1800_id] [int] NULL,
	[sc12ra2_e_macadj_p2200] [nchar](10) NULL,
	[sc12ra2_e_macadj_p2200_id] [int] NULL,
	[sc12ra2_e_macadj_p2359] [nchar](10) NULL,
	[sc12ra2_e_macadj_p2359_id] [int] NULL,
	[sc12ra2_d_dr_p1800] [nchar](10) NULL,
	[sc12ra2_d_dr_p1800_id] [int] NULL,
	[sc12ra2_d_dr_p2200] [nchar](10) NULL,
	[sc12ra2_d_dr_p2200_id] [int] NULL,
	[sc12ra2_d_dr_p2359] [nchar](10) NULL,
	[sc12ra2_d_dr_p2359_id] [int] NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_RateBuilderRates] PRIMARY KEY CLUSTERED 
(
	[rbrid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[RateBuilderRates] ADD  CONSTRAINT [DF_RateBuilderRates_rowguid]  DEFAULT (newid()) FOR [rowguid]
GO




USE [dbCore]
GO

ALTER TABLE [dbo].[RateBuilder] DROP CONSTRAINT [DF_RateBuilder_rowguid]
GO

/****** Object:  Table [dbo].[RateBuilder]    Script Date: 11/15/2016 4:27:20 PM ******/
DROP TABLE [dbo].[RateBuilder]
GO

/****** Object:  Table [dbo].[RateBuilder]    Script Date: 11/14/2016 2:48:05 PM ******/
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


USE [dbCore]
GO

ALTER TABLE [dbo].[RateBuilderComponents] DROP CONSTRAINT [DF_RateBuilderComponents_rowguid]
GO

/****** Object:  Table [dbo].[RateBuilderComponents]    Script Date: 11/15/2016 4:27:20 PM ******/
DROP TABLE [dbo].[RateBuilderComponents]
GO

/****** Object:  Table [dbo].[RateBuilderComponents]    Script Date: 11/14/2016 2:49:31 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[RateBuilderComponents](
	[rbcid] [int] IDENTITY(1,1) NOT NULL,
	[e_edc_r] [nchar](10) NULL,
	[e_edc_sc9_1] [float] NULL,
	[e_edc_sc9_2] [float] NULL,
	[e_edc_sc9_3] [float] NULL,
	[e_edc_sc12_2] [float] NULL,
	[e_tra_r] [nchar](10) NULL,
	[e_tra_sc9_1] [float] NULL,
	[e_tra_sc9_2] [float] NULL,
	[e_tra_sc9_3] [float] NULL,
	[e_tra_sc12_2] [float] NULL,
	[e_mac_r] [nchar](10) NULL,
	[e_mac_1] [float] NULL,
	[e_sbc_r] [float] NULL,
	[e_sbc_1] [float] NULL,
	[e_rpsp_r] [nchar](10) NULL,
	[e_rpsp_1] [float] NULL,
	[e_psls_r] [nchar](10) NULL,
	[e_psls_sc9_1] [float] NULL,
	[e_psls_sc12_2] [float] NULL,
	[e_rdm_r] [nchar](10) NULL,
	[e_rdm_sc9_1] [float] NULL,
	[e_rdm_sc12_2] [float] NULL,
	[e_drs_r] [nchar](10) NULL,
	[e_drs_sc9_1] [float] NULL,
	[e_drs_sc12_2] [float] NULL,
	[e_mfc_r] [nchar](10) NULL,
	[e_mfc_1] [float] NULL,
	[d_mc_r] [nchar](10) NULL,
	[d_mc_1] [float] NULL,
	[d_o5_r] [nchar](10) NULL,
	[d_o5_1] [float] NULL,
	[d_mf86_r] [float] NULL,
	[d_mf86_sc9_2] [float] NULL,
	[d_mf86_sc9_3] [float] NULL,
	[d_mf86_sc12_2] [float] NULL,
	[d_mf810_r] [nchar](10) NULL,
	[d_mf810_sc9_2] [float] NULL,
	[d_mf810_sc9_3] [float] NULL,
	[d_mf810_sc12_2] [float] NULL,
	[d_all_r] [nchar](10) NULL,
	[d_all_sc9_2] [float] NULL,
	[d_all_sc9_3] [float] NULL,
	[d_all_sc12_2] [float] NULL,
	[d_tra_mc_r] [nchar](10) NULL,
	[d_tra_mc_1] [float] NULL,
	[d_tra_o5_r] [nchar](10) NULL,
	[d_tra_o5_1] [float] NULL,
	[d_tra_mf86_r] [float] NULL,
	[d_tra_mf86_sc9_2] [float] NULL,
	[d_tra_mf86_sc9_3] [float] NULL,
	[d_tra_mf86_sc12_2] [float] NULL,
	[d_tra_mf810_r] [nchar](10) NULL,
	[d_tra_mf810_sc9_2] [float] NULL,
	[d_tra_mf810_sc9_3] [float] NULL,
	[d_tra_mf810_sc12_2] [float] NULL,
	[d_tra_all_r] [nchar](10) NULL,
	[d_tra_all_sc9_2] [float] NULL,
	[d_tra_all_sc9_3] [float] NULL,
	[d_tra_all_sc12_2] [float] NULL,
	[d_msccap_r] [nchar](10) NULL,
	[d_msccap_sc9r1_1] [float] NULL,
	[d_msccap_sc9r2_1] [float] NULL,
	[s_cms_r] [nchar](10) NULL,
	[s_cms_sc9_2] [float] NULL,
	[s_bppc_r] [nchar](10) NULL,
	[s_bppc_1] [float] NULL,
	[rowguid] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
 CONSTRAINT [PK_RateBuilderComponents] PRIMARY KEY CLUSTERED 
(
	[rbcid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[RateBuilderComponents] ADD  CONSTRAINT [DF_RateBuilderComponents_rowguid]  DEFAULT (newid()) FOR [rowguid]
GO


