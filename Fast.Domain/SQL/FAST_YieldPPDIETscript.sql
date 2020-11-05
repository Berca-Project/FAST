use [FAST_QA]
go
Drop Table [dbo].[PPReportYieldMCDiets];
CREATE TABLE [dbo].[PPReportYieldMCDiets](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Year] [int] NULL,
	[Week] [int] NULL,
	[LocationID] [bigint] NULL,
	[Location] [varchar](500) NULL,

	[MCFlake] [float] NULL,
	[MCKrosok] [float] NULL,
	[CVIB0069] [float] NULL,
	[CSFR0022] [float] NULL,
	[DSCL0034] [float] NULL,
	[CVIB0070] [float] NULL,
	[RV0054] [float] NULL,
	[DM] [float] NULL,
	[Flake] [float] NULL,
	[MCPacking] [float] NULL,


	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL
 CONSTRAINT [PK_PPReportYieldMCDiets] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO