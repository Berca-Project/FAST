USE [FAST_TRAINING]
GO

/****** Object:  Table [dbo].[WPPPrimaries]    Script Date: 23/01/2020 17.12.56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[WPPPrimaries](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Blend] [varchar](500) NOT NULL,
	[VolPerOps] [int] NOT NULL,
	[StartDate] [date] NOT NULL,
	[Monday] [int] NOT NULL,
	[Tuesday] [int] NOT NULL,
	[Wednesday] [int] NOT NULL,
	[Thursday] [int] NOT NULL,
	[Friday] [int] NOT NULL,
	[Saturday] [int] NOT NULL,
	[Sunday] [int] NOT NULL,
	[LocationID] [bigint] NOT NULL,
	[Location] [varchar](50) NULL,
	[Week] [int] NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_WPPPrimaries] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[WPPPrimaries] ADD  CONSTRAINT [DF_WPPPrimaries_Week]  DEFAULT ((0)) FOR [Week]
GO


