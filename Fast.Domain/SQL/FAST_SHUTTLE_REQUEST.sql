USE [FAST_QA]
GO

/****** Object:  Table [dbo].[ShuttleRequests]    Script Date: 12/01/2020 22:21:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ShuttleRequests](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[CostCenter] [varchar](50) NOT NULL,
	[StartDate] [date] NOT NULL,
	[EndDate] [date] NOT NULL,
	[Time] [time](7) NOT NULL,
	[ProductionCenter] [varchar](100) NOT NULL,
	[EmployeeID] [char](8) NOT NULL,
	[TotalPassengers] [int] NOT NULL,
	[GuestType] [varchar](100) NULL,
	[LocationFrom] [varchar](500) NULL,
	[LocationTo] [varchar](500) NULL,
	[Purpose] [varchar](500) NULL,
	[Department] [varchar](50) NULL,
	[Phone] [varchar](50) NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_ShuttleRequest] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


