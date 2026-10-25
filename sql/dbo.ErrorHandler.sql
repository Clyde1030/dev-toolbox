USE [Claims]
GO

/****** Object:  StoredProcedure [dbo].[ErrorHandler]    Script Date: 2/2/2023 1:14:16 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[ErrorHandler]
AS

	-- NO TRANSACTION BECAUSE THIS PROC DOESN'T AFFECT THE DATABASE

	SET NOCOUNT ON

	DECLARE @errmsg		nvarchar(2048)
	DECLARE @severity	tinyint
	DECLARE @state		tinyint
	DECLARE @errno		int
	DECLARE @proc		sysname
	DECLARE @lineno		int
           
	-- Fill local variables with error info from CATCH block
	SELECT	 @errmsg	= error_message()
			,@severity	= error_severity()
			,@state		= error_state()
			,@errno		= error_number()
			,@proc		= error_procedure()
			,@lineno	= error_line()

	-- Messages without '***' are new and need to be parsed.
	-- Error messages with '***' have already been passed through this procedure and can simply be repeated
	IF @errmsg NOT LIKE '***%'
	BEGIN 
	   SELECT @errmsg = '*** ' + coalesce(quotename(@proc), '<dynamic SQL>') + 
						', ' + ltrim(str(@lineno)) + '. Errno ' + 
						ltrim(str(@errno)) + ': ' + @errmsg
	   RAISERROR(@errmsg, @severity, @state)
	END
	ELSE
	   RAISERROR(@errmsg, @severity, @state)



GO


