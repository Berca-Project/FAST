USE [FAST_QA]
GO

/****** Object:  Table [dbo].[InputOV]    Script Date: 15/01/2020 14:21:50 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[InputOVs](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[WasteCategory] [varchar](100) NULL,
	[Week] [varchar](50) NULL,
	[OV] [varchar](50) NULL,
	[LocationID] [bigint] NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,	
 CONSTRAINT [PK_InputOVs] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


