/****** SSMS の SelectTopNRows コマンドのスクリプト  ******/
SELECT TOP (1000) [LogId]
      ,[Level]
      ,[CallSite]
      ,[Type]
      ,[Message]
      ,[StackTrace]
      ,[InnerException]
      ,[AdditionalInfo]
      ,[LoggedOnDate]
  FROM [TEST001].[dbo].[Logs]
