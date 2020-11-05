USE [FAST_QA]
GO

/****** Object:  Table [dbo].[WPPSTPs]    Script Date: 04/03/2020 09:08:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[WPPSTPs](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Date] [date] NULL,
	[Location] [varchar](50) NULL,
	[Brand] [varchar](50) NULL,
	[Description] [varchar](500) NULL,
	[Packer] [varchar](50) NULL,
	[Maker] [varchar](50) NULL,
	[Shift1] [decimal](15, 4) NULL,
	[Shift2] [decimal](15, 4) NULL,
	[Shift3] [decimal](15, 4) NULL,
	[Activity] [varchar](20) NULL,
	[PONumber] [char](20) NULL,
	[OPSNumber] [char](20) NULL,
	[BatchSAP] [char](20) NULL,
	[Others] [varchar](50) NULL,
	[LocationID] [bigint] NULL,
	[StartDate] [date] NULL,
	[EndDate] [date] NULL,
	[Actual1] [decimal](15, 4) NULL,
	[Actual2] [decimal](15, 4) NULL,
	[Actual3] [decimal](15, 4) NULL,
	[Percentage] [decimal](15, 4) NULL,
	[ReportNote] [varchar](200) NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_WPPSTPs] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[WPPSTPs] ADD  CONSTRAINT [DF_WPPSTPs_Actual1]  DEFAULT ((0)) FOR [Actual1]
GO

ALTER TABLE [dbo].[WPPSTPs] ADD  CONSTRAINT [DF_WPPSTPs_Actual2]  DEFAULT ((0)) FOR [Actual2]
GO

ALTER TABLE [dbo].[WPPSTPs] ADD  CONSTRAINT [DF_WPPSTPs_Actual3]  DEFAULT ((0)) FOR [Actual3]
GO


