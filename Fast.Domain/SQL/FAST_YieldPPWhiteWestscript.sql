USE [FAST_QA]
GO
/****** Object:  Table [dbo].[NpssReportBatch]    Script Date: 3/1/2020 10:36:40 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Drop TABLE [dbo].[PPReportYieldWhiteWest];
CREATE TABLE [dbo].[PPReportYieldWhiteWest](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[EndTime] [datetime] NULL,
	[SapID] [nvarchar](255) NULL,
	[BatchIdent] [nvarchar](255) NULL,
	[BlendCode] [nvarchar](255) NULL,

	[ProducedQTY] [float] NULL,
        [Tobacco] [float] NULL,
        [TotalStems] [float] NULL,
        [TotalExpandedTobacco] [float] NULL,
        [TotalSmallLamina] [float] NULL,
        [TotalRipperShorts] [float] NULL,

        [WetYield] [float] NULL,
        [WetTarget] [float] NULL,

        [DryTobacco] [float] NULL,
        [DryISCRES] [float] NULL,
        [DryET] [float] NULL,
        [DrySL] [float] NULL,
        [DryRS] [float] NULL,
        [FinalOV] [float] NULL,
        [InvoiceOV] [float] NULL,
        [DMBS] [float] NULL,
        [DMBT] [float] NULL,
        [DMBC] [float] NULL,
        [DMAC] [float] NULL,
        [BS] [float] NULL,
        [BT] [float] NULL,
        [BC] [float] NULL,
        [AC] [float] NULL,
        [Packing] [float] NULL,
        [DryYield] [float] NULL,
 CONSTRAINT [PK_PPReportYieldWhiteWests] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
