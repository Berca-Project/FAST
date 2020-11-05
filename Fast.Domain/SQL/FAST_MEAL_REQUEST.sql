USE [FAST_QA]
GO

/****** Object:  Table [dbo].[MealRequests]    Script Date: 12/01/2020 22:21:53 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MealRequests](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[CostCenter] [varchar](50) NOT NULL,
	[StartDate] [date] NOT NULL,
	[EndDate] [date] NOT NULL,
	[ProductionCenter] [varchar](100) NOT NULL,
	[Canteen] [varchar](100) NOT NULL,
	[EmployeeID] [char](8) NOT NULL,
	[TotalGuest] [int] NOT NULL,
	[GuestType] [varchar](100) NULL,
	[Company] [varchar](50) NULL,
	[Guest] [varchar](50) NULL,
	[Purpose] [varchar](500) NULL,
	[Department] [varchar](50) NULL,
	[Shift] [varchar](50) NULL,
	[Phone] [varchar](50) NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_MealRequest] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


