USE [FAST_DEV]
GO
/****** Object:  Table [dbo].[ManPowers]    Script Date: 09/11/2019 17:53:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ManPowers](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[JobTitleID] [bigint] NULL,
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
/****** Object:  Table [dbo].[ReferenceDetails]    Script Date: 09/11/2019 17:53:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReferenceDetails](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[ReferenceID] [bigint] NOT NULL,
	[Code] [nvarchar](150) NOT NULL,
	[Description] [nvarchar](500) NULL,
	[IsDeleted] [bit] NOT NULL,
	[ModifiedBy] [nvarchar](50) NULL,
	[ModifiedDate] [datetime] NULL,
 CONSTRAINT [PK_ReferenceDetails] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[ManPowers] ON 
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (1, 76, 105, CAST(1.00 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (2, 76, 106, CAST(1.00 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (3, 76, 159, CAST(1.00 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (4, 62, 105, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (5, 62, 106, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (6, 9, 105, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (7, 9, 106, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (8, 134, 105, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (9, 134, 106, CAST(0.50 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (10, 134, 159, CAST(1.00 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (11, 134, 160, CAST(1.00 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (12, 13, 105, CAST(0.25 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (13, 13, 106, CAST(0.25 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (14, 13, 159, CAST(0.25 AS Decimal(4, 2)), 0, NULL, NULL)
GO
INSERT [dbo].[ManPowers] ([ID], [JobTitleID], [MachineTypeID], [Value], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (15, 13, 160, CAST(0.25 AS Decimal(4, 2)), 0, NULL, NULL)
GO
SET IDENTITY_INSERT [dbo].[ManPowers] OFF
GO
SET IDENTITY_INSERT [dbo].[ReferenceDetails] ON 
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (1, 1, N'4G', N'4G type', 0, N'admin', CAST(N'2019-09-15T18:34:07.607' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (2, 1, N'3G', N'3G type', 0, N'admin', CAST(N'2019-09-15T18:34:07.613' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (3, 1, N'2G', N'2G type', 0, N'admin', CAST(N'2019-09-15T18:34:07.617' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (4, 1, N'NS', N'NS type', 0, N'admin', CAST(N'2019-09-15T18:34:07.623' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (5, 5, N'PJ', N'Sampoerna Sukorejo', 0, N'admin', CAST(N'2019-09-15T22:33:00.633' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (6, 5, N'PK', N'Sampoerna Karawang', 0, N'admin', CAST(N'2019-09-15T22:33:00.650' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (7, 5, N'PI', N'PT Philip Morris Ind Karawang', 0, N'admin', CAST(N'2019-09-15T22:33:00.650' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (8, 5, N'PB', N'PT SIS Sukorejo', 0, N'admin', CAST(N'2019-09-15T22:33:00.650' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (13, 2, N'3066720AIX', N'IS Planning', 0, N'admin', CAST(N'2019-09-15T23:01:15.043' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (14, 2, N'3066720DIY', N'IS G&A', 0, N'admin', CAST(N'2019-09-15T23:01:15.047' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (15, 2, N'3066720GIY', N'IS Operations', 0, N'admin', CAST(N'2019-09-15T23:01:15.050' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (22, 4, N'ID', N'Indonesia', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (23, 4, N'MY', N'Malaysia', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (24, 6, N'PP', N'Primary Processing', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (25, 6, N'SP', N'Secondary Processing', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (26, 6, N'SO', N'Secondary OTP', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (54, 7, N'CL', N'Clove', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.750' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (55, 7, N'LM', N'Lamina', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.767' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (56, 7, N'CR', N'CRES', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.767' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (57, 7, N'LS', N'Laser', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.767' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (58, 7, N'CB', N'Case Baller', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.767' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (59, 7, N'PT', N'Perforating', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.767' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (60, 7, N'CS', N'Casing', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.783' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (61, 7, N'DT', N'DIET', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.783' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (62, 7, N'PW', N'Primary White', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.783' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (63, 7, N'CO', N'Clove Oil', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.783' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (64, 7, N'PL', N'Packer', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.797' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (65, 7, N'FT', N'KDF Filter', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.797' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (66, 7, N'RIP', N'Ripper', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.797' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (67, 7, N'SZ', N'Senzani', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.797' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (68, 7, N'RP', N'Robot Palletizer', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.797' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (69, 7, N'MT', N'Mentholation', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (70, 7, N'MP', N'Manual Process', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (71, 7, N'RW', N'Rework', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (72, 7, N'OR', N'Overolling', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (73, 7, N'CT', N'Cutting', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (74, 7, N'TD', N'TDC', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.813' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (75, 7, N'RT', N'RTC', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.830' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (76, 7, N'MK', N'Maker', 0, N'ADMIN', CAST(N'2019-09-19T21:26:50.830' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (84, 9, N'LU51', N'LU51 Desc', 0, N'ADMIN', CAST(N'2019-09-20T18:54:46.213' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (85, 9, N'LU52', N'LU52 Desc', 0, N'ADMIN', CAST(N'2019-09-20T18:54:46.223' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (86, 8, N'PC', N'Production Center', 0, N'ADMIN', CAST(N'2019-09-22T15:20:09.913' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (87, 8, N'Dep', N'Department', 0, N'ADMIN', CAST(N'2019-09-22T15:20:09.930' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (88, 8, N'SubDep', N'Sub Department', 0, N'ADMIN', CAST(N'2019-09-22T15:20:09.930' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (89, 8, N'Country', N'Country', 1, N'ADMIN', CAST(N'2019-09-22T15:20:09.930' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (105, 10, N'Packer', N'test', 0, N'ADMIN', CAST(N'2019-11-09T17:18:16.600' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (106, 10, N'Maker', N'test', 0, N'ADMIN', CAST(N'2019-11-09T17:18:16.653' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (107, 10, N'Demo', N'test', 1, N'ADMIN', CAST(N'2019-09-25T11:53:15.163' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (108, 10, N'Protos', NULL, 0, N'ADMIN', CAST(N'2019-11-09T17:18:16.660' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (119, 11, N'SGC', N'SIS Gruduk Cleaning / IZORA Cleaning', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.047' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (120, 11, N'CBM', N'Conditional Based Maintenance', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.063' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (121, 11, N'PM', N'Preventive Maintenance', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.063' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (122, 11, N'DC', N'Deep Cleaning', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.077' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (123, 11, N'DT', N'Down Time / Break Down', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.077' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (124, 11, N'ND', N'No Demand', 0, N'ADMIN', CAST(N'2019-09-25T23:13:50.077' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (125, 3, N'FA067814.08', N'SAMPOERNA U MILD LT BOX 16 SLI', 0, N'ADMIN', CAST(N'2019-09-26T06:10:50.057' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (126, 3, N'FA067814.09', N'MARLBORO', 0, N'ADMIN', CAST(N'2019-09-26T06:10:50.070' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (127, 12, N'#ffff00', N'Cuti Bersama HMS', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.490' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (128, 12, N'#ff3300', N'Libur Nasional', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.503' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (129, 12, N'#88cc00', N'Cuti Bersama Pemerintah', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.513' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (130, 12, N'#cc0000', N'Shift Off', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.520' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (131, 12, N'#3366cc', N'Acara Internal HMS', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.527' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (132, 4, N'SG', N'Singapore', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (133, 5, N'SS', N'Sampoerna Serawak', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (134, 5, N'PX', N'PT Test', 0, N'ADMIN', CAST(N'2019-10-07T19:44:48.183' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (135, 5, N'PZ', N'PT Semarang', 0, N'ADMIN', CAST(N'2019-10-07T20:00:09.270' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (137, 6, N'PT1', N'Primary Test', 0, N'ADMIN', CAST(N'2019-10-08T01:56:17.573' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (138, 7, N'LT1', N'Laser Test', 0, N'ADMIN', CAST(N'2019-10-08T02:00:54.307' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (139, 5, N'PS1', N'PT Surabaya', 0, N'ADMIN', CAST(N'2019-10-08T02:02:00.323' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (140, 6, N'PT2', N'Primary Test2', 0, N'ADMIN', CAST(N'2019-10-08T02:02:22.527' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (141, 7, N'LM1', N'Lamina Test ZS', 0, N'ADMIN', CAST(N'2019-10-08T04:27:02.733' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (142, 4, N'VH', N'Vietnam', 0, N'ADMIN', CAST(N'2019-10-08T02:25:52.083' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (143, 5, N'PS2', N'PT Sidoarjo X', 0, N'ADMIN', CAST(N'2019-10-08T04:27:53.530' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (144, 6, N'PT3', N'Primary Test3 D', 0, N'ADMIN', CAST(N'2019-10-08T04:27:30.257' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (145, 7, N'CB1', N'CB Test', 0, N'ADMIN', CAST(N'2019-10-08T02:27:31.507' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (146, 5, N'HN', N'Hanoi', 0, N'ADMIN', CAST(N'2019-10-08T03:19:38.327' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (147, 5, N'HC', N'Ho Chi Min', 0, N'ADMIN', CAST(N'2019-10-08T03:20:54.317' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (148, 13, N'P700', N'HAUNI', 1, N'ADMIN', CAST(N'2019-10-21T12:24:31.920' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (149, 13, N'GDX', N'GDX', 0, N'ADMIN', CAST(N'2019-10-21T12:24:31.930' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (150, 14, N'V0N5PC
', N'V0N5PC
', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (151, 14, N'V0CGI
', N'V0CGI
', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (152, 14, N'D0A1VC
', N'D0A1VC
', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (153, 12, N'#111111', N'ISENG', 0, N'ADMIN', CAST(N'2019-10-23T08:55:59.537' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (154, 5, N'FK', N'FERY TEST', 0, N'ADMIN', CAST(N'2019-10-25T01:59:55.873' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (155, 13, N'Garbuio', N'Garbuio', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (156, 15, N'HMS KRW', N'HMS KRW', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (157, 15, N'HMS SKJ', N'HMS SKJ', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (158, 15, N'PMIID', N'PMIID', 0, NULL, NULL)
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (159, 10, N'Case Packer', N'CP desc', 0, N'ADMIN', CAST(N'2019-11-09T17:18:16.670' AS DateTime))
GO
INSERT [dbo].[ReferenceDetails] ([ID], [ReferenceID], [Code], [Description], [IsDeleted], [ModifiedBy], [ModifiedDate]) VALUES (160, 10, N'Robot', N'Robot Desc', 0, N'ADMIN', CAST(N'2019-11-09T17:18:21.483' AS DateTime))
GO
SET IDENTITY_INSERT [dbo].[ReferenceDetails] OFF
GO
ALTER TABLE [dbo].[ReferenceDetails] ADD  CONSTRAINT [DF_ReferenceDetails_IsDeleted]  DEFAULT ((0)) FOR [IsDeleted]
GO
ALTER TABLE [dbo].[ReferenceDetails]  WITH NOCHECK ADD  CONSTRAINT [FK_ReferenceDetails_References] FOREIGN KEY([ReferenceID])
REFERENCES [dbo].[References] ([ID])
NOT FOR REPLICATION 
GO
ALTER TABLE [dbo].[ReferenceDetails] NOCHECK CONSTRAINT [FK_ReferenceDetails_References]
GO
