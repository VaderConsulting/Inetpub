IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'Domain Control')
	DROP DATABASE [Domain Control]
GO

CREATE DATABASE [Domain Control]  ON (NAME = N'Domain Control_dat', FILENAME = N'C:\MSSQL7\data\Domain Control.mdf' , SIZE = 3, FILEGROWTH = 10%) LOG ON (NAME = N'Domain Control_log', FILENAME = N'C:\MSSQL7\data\Domain Control.ldf' , SIZE = 3, FILEGROWTH = 10%)
GO

exec sp_dboption N'Domain Control', N'autoclose', N'true'
GO

exec sp_dboption N'Domain Control', N'bulkcopy', N'true'
GO

exec sp_dboption N'Domain Control', N'trunc. log', N'true'
GO

exec sp_dboption N'Domain Control', N'torn page detection', N'false'
GO

exec sp_dboption N'Domain Control', N'read only', N'false'
GO

exec sp_dboption N'Domain Control', N'dbo use', N'false'
GO

exec sp_dboption N'Domain Control', N'single', N'false'
GO

exec sp_dboption N'Domain Control', N'autoshrink', N'true'
GO

exec sp_dboption N'Domain Control', N'ANSI null default', N'false'
GO

exec sp_dboption N'Domain Control', N'recursive triggers', N'false'
GO

exec sp_dboption N'Domain Control', N'ANSI nulls', N'false'
GO

exec sp_dboption N'Domain Control', N'concat null yields null', N'false'
GO

exec sp_dboption N'Domain Control', N'cursor close on commit', N'false'
GO

exec sp_dboption N'Domain Control', N'default to local cursor', N'false'
GO

exec sp_dboption N'Domain Control', N'quoted identifier', N'false'
GO

exec sp_dboption N'Domain Control', N'ANSI warnings', N'false'
GO

exec sp_dboption N'Domain Control', N'auto create statistics', N'true'
GO

exec sp_dboption N'Domain Control', N'auto update statistics', N'true'
GO

use [Domain Control]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblAgentReg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblAgentReg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblCPU]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblCPU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblConnections]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblConnections]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblContacts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblContacts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblDiskSpace]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblDiskSpace]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblEmail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblEmail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblErrors]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblErrors]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblEvents]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblEvents]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblHosts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblHosts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblLocations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblLocations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblRoute]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblRoute]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblRouteNames]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblRouteNames]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblSubnets]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblSubnets]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTaskPath]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTaskPath]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTasks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTasks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTrace]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTrace]
GO

if not exists (select * from master.dbo.syslogins where loginname = N'Air_Test')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'British'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'Air_Test', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'ConstructionUser')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'us_english'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'ConstructionUser', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DELLP1000\IUSR_DELLP1000')
	exec sp_grantlogin N'DELLP1000\IUSR_DELLP1000'
	exec sp_defaultdb N'DELLP1000\IUSR_DELLP1000', N'TOMSMonitor'
	exec sp_defaultlanguage N'DELLP1000\IUSR_DELLP1000', N'British'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'distributor_admin')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'us_english'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'distributor_admin', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\Administrator')
	exec sp_grantlogin N'DOMAIN\Administrator'
	exec sp_defaultdb N'DOMAIN\Administrator', N'master'
	exec sp_defaultlanguage N'DOMAIN\Administrator', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\IUSR_SERVER1')
	exec sp_grantlogin N'DOMAIN\IUSR_SERVER1'
	exec sp_defaultdb N'DOMAIN\IUSR_SERVER1', N'Process_Control_Data'
	exec sp_defaultlanguage N'DOMAIN\IUSR_SERVER1', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\IUSR_Server3')
	exec sp_grantlogin N'DOMAIN\IUSR_Server3'
	exec sp_defaultdb N'DOMAIN\IUSR_Server3', N'master'
	exec sp_defaultlanguage N'DOMAIN\IUSR_Server3', N'British'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\ProgTrainee3')
	exec sp_grantlogin N'DOMAIN\ProgTrainee3'
	exec sp_defaultdb N'DOMAIN\ProgTrainee3', N'master'
	exec sp_defaultlanguage N'DOMAIN\ProgTrainee3', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\stanner')
	exec sp_grantlogin N'DOMAIN\stanner'
	exec sp_defaultdb N'DOMAIN\stanner', N'master'
	exec sp_defaultlanguage N'DOMAIN\stanner', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\webdeveloper2')
	exec sp_grantlogin N'DOMAIN\webdeveloper2'
	exec sp_defaultdb N'DOMAIN\webdeveloper2', N'master'
	exec sp_defaultlanguage N'DOMAIN\webdeveloper2', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'ITP_User')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'British'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'ITP_User', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'JiwaUser')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'us_english'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'JiwaUser', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'SIAAdmin')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'British'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'SIAAdmin', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'TQMUser')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'master', @loginlang = N'British'
	if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'TQMUser', null, @logindb, @loginlang
END
GO

if not exists (select * from master.dbo.syslogins where loginname = N'TUSK\IUSR_MAMOTH02')
	exec sp_grantlogin N'TUSK\IUSR_MAMOTH02'
	exec sp_defaultdb N'TUSK\IUSR_MAMOTH02', N'TOMSMonitor'
	exec sp_defaultlanguage N'TUSK\IUSR_MAMOTH02', N'British'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'DOMAIN\Domain Users')
	exec sp_grantlogin N'DOMAIN\Domain Users'
	exec sp_defaultdb N'DOMAIN\Domain Users', N'Process_Control_Data'
	exec sp_defaultlanguage N'DOMAIN\Domain Users', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'TUSK\Domain Users')
	exec sp_grantlogin N'TUSK\Domain Users'
	exec sp_defaultdb N'TUSK\Domain Users', N'Process_Control_Data'
	exec sp_defaultlanguage N'TUSK\Domain Users', N'us_english'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'TUSK\Programmers')
	exec sp_grantlogin N'TUSK\Programmers'
	exec sp_defaultdb N'TUSK\Programmers', N'master'
	exec sp_defaultlanguage N'TUSK\Programmers', N'British'
GO

if not exists (select * from master.dbo.syslogins where loginname = N'TUSK\SQLDBAs')
	exec sp_grantlogin N'TUSK\SQLDBAs'
	exec sp_defaultdb N'TUSK\SQLDBAs', N'master'
	exec sp_defaultlanguage N'TUSK\SQLDBAs', N'British'
GO

exec sp_addsrvrolemember N'TUSK\SQLDBAs', sysadmin
GO

if not exists (select * from dbo.sysusers where name = N'BUILTIN\Administrators' and uid < 16382)
	EXEC sp_grantdbaccess N'BUILTIN\Administrators', N'BUILTIN\Administrators'
GO

if not exists (select * from dbo.sysusers where name = N'DELLP1000\IUSR_DELLP1000' and uid < 16382)
	EXEC sp_grantdbaccess N'DELLP1000\IUSR_DELLP1000', N'DELLP1000\IUSR_DELLP1000'
GO

exec sp_addrolemember N'db_owner', N'BUILTIN\Administrators'
GO

CREATE TABLE [dbo].[tblAgentReg] (
	[AgentID] [int] IDENTITY (1, 1) NOT NULL ,
	[Hostname] [varchar] (15) NULL ,
	[TimerInterval] [int] NULL ,
	[MaxProcesses] [int] NULL ,
	[ExclusiveMode] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblCPU] (
	[Value] [int] NULL ,
	[DateTime] [smalldatetime] NOT NULL ,
	[Hostname] [varchar] (15) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblConnections] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (50) NULL ,
	[Domainname] [varchar] (50) NULL ,
	[Username] [varchar] (50) NULL ,
	[Password] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblContacts] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[FName] [varchar] (50) NULL ,
	[SName] [varchar] (50) NULL ,
	[Title] [varchar] (50) NULL ,
	[Phone] [varchar] (11) NULL ,
	[Mobile] [varchar] (12) NULL ,
	[Email] [varchar] (50) NULL ,
	[LocationID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDiskSpace] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[HostID] [int] NOT NULL ,
	[DriveLetter] [varchar] (1) NOT NULL ,
	[UsedSpace] [varchar] (15) NULL ,
	[TotalSpace] [varchar] (15) NULL ,
	[DateTime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblEmail] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Address] [varchar] (255) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblErrors] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Time] [datetime] NULL ,
	[Source] [nvarchar] (50) NULL ,
	[Number] [int] NULL ,
	[Description] [nvarchar] (255) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblEvents] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Time] [datetime] NULL ,
	[HostID] [int] NULL ,
	[Event] [varchar] (255) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblHosts] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Domainname] [varchar] (50) NULL ,
	[Hostname] [varchar] (50) NULL ,
	[IP] [varchar] (15) NULL ,
	[isPDC] [bit] NOT NULL ,
	[isDC] [bit] NOT NULL ,
	[isSQL] [bit] NOT NULL ,
	[isPrint] [bit] NOT NULL ,
	[isRAS] [bit] NOT NULL ,
	[isUnix] [bit] NOT NULL ,
	[isBrowser] [bit] NOT NULL ,
	[isTS] [bit] NOT NULL ,
	[isMSCS] [bit] NOT NULL ,
	[isTime] [bit] NOT NULL ,
	[isServer] [bit] NOT NULL ,
	[isUp] [bit] NOT NULL ,
	[Type] [varchar] (128) NULL ,
	[DateTime] [datetime] NULL ,
	[inProgress] [bit] NOT NULL ,
	[Retries] [int] NOT NULL ,
	[Timeout] [int] NOT NULL ,
	[EmailID] [int] NOT NULL ,
	[ConnectionID] [int] NOT NULL ,
	[LocationID] [int] NOT NULL ,
	[RouteID] [int] NOT NULL ,
	[TaskID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLocations] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (50) NULL ,
	[Address] [varchar] (255) NULL ,
	[Suburb] [varchar] (50) NULL ,
	[Postcode] [int] NULL ,
	[State] [varchar] (25) NULL ,
	[Country] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLog] (
	[TaskID] [int] NULL ,
	[Event] [int] NULL ,
	[Result] [varchar] (50) NULL ,
	[EventDateTime] [datetime] NULL ,
	[Agent] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblRoute] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[HostID] [int] NULL ,
	[Path] [varchar] (1024) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblRouteNames] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[IP] [varchar] (15) NOT NULL ,
	[Name] [varchar] (128) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblSubnets] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Subnet] [varchar] (15) NULL ,
	[LocationID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTaskPath] (
	[Type] [int] NULL ,
	[Params] [varchar] (50) NULL ,
	[TaskPath] [varchar] (50) NULL ,
	[DisplayName] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTasks] (
	[TaskID] [int] IDENTITY (1, 1) NOT NULL ,
	[Type] [int] NULL ,
	[Status] [int] NULL ,
	[AgentAllow] [int] NULL ,
	[AgentRunning] [int] NULL ,
	[NumInstance] [int] NULL ,
	[MaxInstance] [int] NULL ,
	[NextDateTime] [datetime] NULL ,
	[LastUpdate] [datetime] NULL ,
	[Schedule] [varchar] (10) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTrace] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[EventID] [int] NULL ,
	[Result] [varchar] (1024) NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblContacts] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblContacts] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblEvents] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblEvents] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblHosts] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblHosts_1] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLocations] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblLocations] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblRoute] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblRoute] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblRouteNames] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblRouteNames] PRIMARY KEY  CLUSTERED 
	(
		[IP]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblSubnets] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblSubnets] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTrace] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblTrace] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblCPU] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblCPU_DateTime] DEFAULT (getdate()) FOR [DateTime]
GO

ALTER TABLE [dbo].[tblContacts] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblContacts_LocationID] DEFAULT (0) FOR [LocationID]
GO

ALTER TABLE [dbo].[tblDiskSpace] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblDiskSpace_DateTime] DEFAULT (getdate()) FOR [DateTime]
GO

ALTER TABLE [dbo].[tblHosts] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblHosts_isPDC] DEFAULT (0) FOR [isPDC],
	CONSTRAINT [DF_tblHosts_isDC] DEFAULT (0) FOR [isDC],
	CONSTRAINT [DF_tblHosts_isSQL] DEFAULT (0) FOR [isSQL],
	CONSTRAINT [DF_tblHosts_isPrint] DEFAULT (0) FOR [isPrint],
	CONSTRAINT [DF_tblHosts_isRAS] DEFAULT (0) FOR [isRAS],
	CONSTRAINT [DF_tblHosts_isUnix] DEFAULT (0) FOR [isUnix],
	CONSTRAINT [DF_tblHosts_isBrowser] DEFAULT (0) FOR [isBrowser],
	CONSTRAINT [DF_tblHosts_isTS] DEFAULT (0) FOR [isTS],
	CONSTRAINT [DF_tblHosts_isMSCS] DEFAULT (0) FOR [isMSCS],
	CONSTRAINT [DF_tblHosts_isTime] DEFAULT (0) FOR [isTime],
	CONSTRAINT [DF_tblHosts_isServer] DEFAULT (0) FOR [isServer],
	CONSTRAINT [DF_tblHosts_isUp] DEFAULT (0) FOR [isUp],
	CONSTRAINT [DF_tblHosts_inProgress] DEFAULT (0) FOR [inProgress],
	CONSTRAINT [DF_tblHosts_Retries] DEFAULT (0) FOR [Retries],
	CONSTRAINT [DF_tblHosts_Timeout] DEFAULT (250) FOR [Timeout],
	CONSTRAINT [DF_tblHosts_EmailID] DEFAULT (0) FOR [EmailID],
	CONSTRAINT [DF_tblHosts_ConnectionID] DEFAULT (0) FOR [ConnectionID],
	CONSTRAINT [DF_tblHosts_LocationID] DEFAULT (0) FOR [LocationID],
	CONSTRAINT [DF_tblHosts_RouteID] DEFAULT (0) FOR [RouteID],
	CONSTRAINT [DF_tblHosts_TaskID] DEFAULT (0) FOR [TaskID]
GO

 CREATE  INDEX [IX_tblDiskSpace] ON [dbo].[tblDiskSpace]([HostID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

GRANT  SELECT  ON [dbo].[tblEmail]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblErrors]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblEvents]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblHosts]  TO [public]
GO

GRANT  SELECT  ON [dbo].[tblHosts]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblRoute]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblRouteNames]  TO [DELLP1000\IUSR_DELLP1000]
GO

GRANT  SELECT  ON [dbo].[tblTrace]  TO [DELLP1000\IUSR_DELLP1000]
GO

