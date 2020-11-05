USE [FAST_QA]
GO

/****** Object:  Table [dbo].[ManPowers]    Script Date: 24/01/2020 10.31.29 ******/
DROP TABLE [dbo].[ManPowers]
GO

/****** Object:  Table [dbo].[ManPowers]    Script Date: 24/01/2020 10.31.29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ManPowers](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[JobTitleID] [bigint] NULL,
	[RoleName] [varchar](50) NULL,
	[MachineTypeID] [bigint] NULL,
	[Value] [decimal](4, 2) NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_ManPowers] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


