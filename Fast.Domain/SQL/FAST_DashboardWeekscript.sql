use [FAST_QA]
go
drop table [dbo].[DashboardWeeks];
CREATE TABLE [dbo].[DashboardWeeks](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[Week] [bigint] NOT NULL,
	[Year] [bigint] NOT NULL,
	[Submitted] [int] NOT NULL,
	[Approved] [int] NOT NULL,
	[Location] [varchar](500) NOT NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedDate] [datetime]  NOT NULL
 CONSTRAINT [PK_DashboardWeeks] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
