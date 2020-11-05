USE [FAST_TRAINING]
GO

/****** Object:  Table [dbo].[InputDailies]    Script Date: 2/10/2020 1:29:16 AM ******/
DROP TABLE [dbo].[InputDailies]
GO

/****** Object:  Table [dbo].[InputDailies]    Script Date: 2/10/2020 1:29:16 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[InputDailies](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[ProdCenterID] [bigint] NULL,
	[Shift] [varchar](50) NULL,
	[LinkUp] [varchar](50) NULL,
	[Date] [datetime] NULL,
	[Week] [varchar](20) NULL,
	[MTBFValue] [decimal](18, 8) NULL,
	[MTBFFocus] [varchar](200) NULL,
	[MTBFActPlan] [varchar](200) NULL,
	[CPQIValue] [decimal](18, 8) NULL,
	[CPQIFocus] [varchar](200) NULL,
	[CPQIActPlan] [varchar](200) NULL,
	[VQIValue] [decimal](18, 8) NULL,
	[VQIFocus] [varchar](200) NULL,
	[VQIActPlan] [varchar](200) NULL,
	[WorkingValue] [decimal](18, 8) NULL,
	[WorkingFocus] [varchar](200) NULL,
	[WorkingActPlan] [varchar](200) NULL,
	[UptimeValue] [decimal](18, 8) NULL,
	[UptimeFocus] [varchar](200) NULL,
	[UptimeActPlan] [varchar](200) NULL,
	[STRSValue] [decimal](18, 8) NULL,
	[STRSFocus] [varchar](200) NULL,
	[STRSActPlan] [varchar](200) NULL,
	[ProdVolumeValue] [decimal](18, 8) NULL,
	[ProdVolumeFocus] [varchar](200) NULL,
	[ProdVolumeActPlan] [varchar](200) NULL,
	[CRRValue] [decimal](18, 8) NULL,
	[CRRFocus] [varchar](200) NULL,
	[CRRActPlan] [varchar](200) NULL,
	[LocationID] [bigint] NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_InputDaily] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


