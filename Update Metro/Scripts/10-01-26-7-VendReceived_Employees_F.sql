SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER FUNCTION [dbo].[VendReceived_Employees_F] (@EntryID INT)
RETURNS VARCHAR(1000)
AS
BEGIN
    DECLARE @ReturnStr AS VARCHAR(1000);

    SELECT @ReturnStr = STUFF((
        SELECT ', ' + E.Name
        FROM dbo.VendReceived_Employees VRE
        INNER JOIN dbo.Employees E ON VRE.EmpID = E.EmpID
        WHERE VRE.VR_RefID = @EntryID
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '');

    RETURN ISNULL(@ReturnStr, '');
END
GO
