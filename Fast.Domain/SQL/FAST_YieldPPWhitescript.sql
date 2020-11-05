use [FAST_QA]
go
Drop Table [dbo].[PPReportYieldWhites];
CREATE TABLE [dbo].[PPReportYieldWhites](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Year] [int] NULL,
	[Week] [int] NULL,
	[LocationID] [bigint] NULL,
	[Location] [varchar](500) NULL,

	[Blend] [varchar](500) NULL,
	[InfeedMaterialWet] [float] NULL,
	[InfeedMaterialDry] [float] NULL,
	[SumInputMaterialDry] [float] NULL,
	[CutFiller] [float] NULL,
	[RS_Addback] [float] NULL,
	[CFDry] [float] NULL,
	[RS_AddDry] [float] NULL,
	[CFDryExclude] [float] NULL,
	[RS_AddbackWet] [float] NULL,
	[AvgFinalOV] [float] NULL,
	[RS_AddbackPercen] [float] NULL,
	[DryYield] [float] NULL,
	[WetYield] [float] NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL
 CONSTRAINT [PK_PPReportYieldWhites] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO