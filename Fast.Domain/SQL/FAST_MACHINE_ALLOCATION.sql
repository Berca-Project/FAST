USE [FAST_QA]
GO

/****** Object:  Table [dbo].[MachineAllocations]    Script Date: 19/01/2020 09:26:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MachineAllocations](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[MachineID] [bigint] NOT NULL,
	[MachineCode] [varchar](20) NULL,
	[MachineCategory] [varchar](500) NULL,
	[Value] [float] NULL DEFAULT 1,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_MachineAllocations] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


