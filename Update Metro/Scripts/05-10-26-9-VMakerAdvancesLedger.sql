USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VMakerAdvancesLedger]    Script Date: 10/5/2026 6:11:46 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VMakerAdvancesLedger]
AS
SELECT VendID,T1.* FROM
(SELECT EntryID,AccNo,DT,Description,Amount,AccVoucherNo,DAmount,'' AS Remarks FROM MakerAdvances
UNION
SELECT MakerAmtCleared.EntryID,MakerAmtCleared.AccNo,MakerAmtCleared.DT,'Deduction in Bill',-AmtClrd,MakerAmtCleared.VchrNo,0 AS DAmount,MakerLoanClearance_Manual.Remarks FROM MakerAmtCleared
LEFT JOIN MakerLoanClearance_Manual ON MakerAmtCleared.EntryID=MakerLoanClearance_Manual.MPB_STD_EntryID
) T1 INNER JOIN Makers ON T1.AccNo=Makers.AccNo



GO


