USE [TEST001]
GO

/****** Object:  StoredProcedure [dbo].[InsertLog]    Script Date: 2020/03/15 13:32:31 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create procedure [dbo].[InsertLog] 
(
	@level varchar(max),
	@callSite varchar(max),
	@type varchar(max),
	@message varchar(max),
	@stackTrace varchar(max),
	@innerException varchar(max),
	@additionalInfo varchar(max)
)
as
insert into dbo.Logs
(
	[Level],
	CallSite,
	[Type],
	[Message],
	StackTrace,
	InnerException,
	AdditionalInfo
)
values
(
	@level,
	@callSite,
	@type,
	@message,
	@stackTrace,
	@innerException,
	@additionalInfo
)

GO

