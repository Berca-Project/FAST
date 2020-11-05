USE [FAST_USE]
GO

/****** Object:  Table [dbo].[InputDailies]    Script Date: 1/27/2020 12:21:47 AM ******/
DROP TABLE [dbo].[InputDailies]
GO

/****** Object:  Table [dbo].[InputDailies]    Script Date: 1/27/2020 12:21:47 AM ******/
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
	[KPI] [varchar](50) NULL,
	[LinkUp] [varchar](50) NULL,
	[Date] [datetime] NULL,
	[Week] [varchar](20) NULL,
	[LocationID] [bigint] NULL,
	[Value] [decimal](18, 8) NULL,
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


