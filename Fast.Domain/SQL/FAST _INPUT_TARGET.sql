USE [FAST_USE]
GO

/****** Object:  Table [dbo].[InputTargets]    Script Date: 1/25/2020 12:15:02 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[InputTargets](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[ProdCenterID] [bigint] NULL,
	[KPI] [varchar](50) NULL,
	[Version] [varchar](50) NULL,
	[Month] [varchar](20) NULL,
	[Value] [decimal](18, 8) NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_InputTargets] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


