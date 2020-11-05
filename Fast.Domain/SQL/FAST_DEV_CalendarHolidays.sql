USE [FAST_DEV]
GO

/****** Object:  Table [dbo].[CalendarHolidays]    Script Date: 13/11/2019 6:16:43 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CalendarHolidays](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Date] [date] NOT NULL,
	[Description] [varchar](500) NOT NULL,
	[Color] [varchar](7) NULL,
	[HolidayTypeID] [bigint] NOT NULL,
	[LocationID] [bigint] NOT NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_CalendarHolidays] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


