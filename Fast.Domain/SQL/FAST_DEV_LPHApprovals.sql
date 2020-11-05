USE [FAST_DEV]
GO

/****** Object:  Table [dbo].[LPHApprovals]    Script Date: 12/11/2019 6:22:53 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[LPHApprovals](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[LPHSubmissionID] [bigint] NOT NULL,
	[UserID] [bigint] NOT NULL,
	[Status] [nchar](10) NOT NULL,
	[Notes] [nvarchar](150) NULL,
	[LocationID] [bigint] NOT NULL,
	[ApproverID] [bigint] NOT NULL,
	[Shift] [varchar](10) NULL,
	[Date] [date] NOT NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [nvarchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_LPHApprovals] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


