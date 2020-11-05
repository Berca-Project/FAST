USE [FAST_DEV]
GO

/****** Object:  Table [dbo].[UserRoles]    Script Date: 01/12/2019 18:47:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[UserRoles](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[UserID] [bigint] NOT NULL,
	[RoleName] [nvarchar](50) NOT NULL,
	[IsDeleted] [bit] NULL,
	[ModifiedBy] [varchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_UserRoles] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


