
-- 7. Ensure GetNextAccno function handles alphabetical and numeric account prefixes safely
CREATE OR ALTER FUNCTION [dbo].[GetNextAccno](@AccountName AS VARCHAR(255),@AccType AS VARCHAR(50),@ParentAccount As Varchar(255)=NULL,@IsParent As BIT=0)
RETURNS VARCHAR(255) AS  
BEGIN
DECLARE @AccNo AS VARCHAR(255)
DECLARE @NewVal AS VARCHAR(255), @Prefix As VARCHAR(50)

IF @IsParent=1
	BEGIN
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) from Accounts WHERE SubAccOf=@ParentAccount)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
ELSE
	BEGIN
		SET @Prefix=ASCII(UPPER(LEFT(LTRIM(RTRIM(ISNULL(@AccountName,''))),1)))
		IF @Prefix>=65 AND @Prefix<=90
			SET @Prefix=REPLICATE('0',2-LEN(@Prefix-64)) + CAST(@Prefix-64 AS VARCHAR)
		ELSE 
			SET @Prefix='00'
		
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) FROM Accounts WHERE SubAccOf=@ParentAccount AND SUBSTRING(AccNo,LEN(Accno)-4,2)=@Prefix)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + @Prefix + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
RETURN @AccNo
END
