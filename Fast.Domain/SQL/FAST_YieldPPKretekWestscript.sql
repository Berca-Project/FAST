USE [FAST_QA]
GO
/****** Object:  Table [dbo].[NpssReportBatch]    Script Date: 3/1/2020 10:36:40 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Drop TABLE [dbo].[PPReportYieldKretekWest];
CREATE TABLE [dbo].[PPReportYieldKretekWest](
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
        [TotalCloves] [float] NULL,
        [TotalOffspec] [float] NULL,
        [CSF] [float] NULL,
        [WetYield] [float] NULL,
        [WetTarget] [float] NULL,

        [DryTobacco] [float] NULL,
        [DryISCRES] [float] NULL,
        [DryRTC] [float] NULL,
        [DryET] [float] NULL,
        [DryCLOVE] [float] NULL,
        [DryCSF] [float] NULL,
        [DryOfSpec] [float] NULL,
        [FinalOV] [float] NULL,
        [InvoiceOV] [float] NULL,
        [DMBC] [float] NULL,
        [TotalBrightCasing] [float] NULL,
        [TotalBurleySpray] [float] NULL,
        [TotalAfterCut] [float] NULL,
        [Packing] [float] NULL,
        [DryYield] [float] NULL,
        [DMAC] [float] NULL,
        [DryCasing] [float] NULL,
        [DryAC] [float] NULL,
 CONSTRAINT [PK_PPReportYieldKretekWests] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
