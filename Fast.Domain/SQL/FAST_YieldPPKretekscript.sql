use [FAST_QA]
go
Drop Table [dbo].[PPReportYieldKreteks];
CREATE TABLE [dbo].[PPReportYieldKreteks](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Year] [int] NULL,
	[Week] [int] NULL,
	[LocationID] [bigint] NULL,
	[Location] [varchar](500) NULL,

	[Blend] [varchar](500) NULL,
	
	[Leaf] [float] NULL,
	[LeafOV] [float] NULL,
	[Clove] [float] NULL,
	[CloveOV] [float] NULL,
	[CRES] [float] NULL,
	[CRESOV] [float] NULL,
	[DIET] [float] NULL,
	[DIETOV] [float] NULL,
	[RTC] [float] NULL,
	[RTCOV] [float] NULL,
	[SmallLamina] [float] NULL,
	[SmallLaminaOV] [float] NULL,
	[CloveSteamFlake] [float] NULL,
	[CloveSteamFlakeOV] [float] NULL,

	[InfeedMaterialWet] [float] NULL,
	[InfeedMaterialDry] [float] NULL,

	[BCWet] [float] NULL,
	[BCDryMatter] [float] NULL,
	[ACWet] [float] NULL,
	[ACDryMatter] [float] NULL,
	
	[CutFiller] [float] NULL,
	[RipperShort] [float] NULL,
	[Addback] [float] NULL,

	[CFDry] [float] NULL,
	[RSDry] [float] NULL,
	[AddbackDry] [float] NULL,

	[CFDryExclude] [float] NULL,
	[AvgFinalOV] [float] NULL,
	[RipperMC] [float] NULL,
	[AddbackMC] [float] NULL,

	[DryYield] [float] NULL,
	[WetYield] [float] NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL
 CONSTRAINT [PK_PPReportYieldKreteks] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO