USE [DailyMTMDB]
GO

/****** Object:  StoredProcedure [dbo].[SP_Total_Contracts_Traded_Report]    Script Date: 2024/03/11 19:31:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[SP_Total_Contracts_Traded_Report]
    @DateFrom DATE,
    @DateTo DATE
AS
BEGIN

    SET NOCOUNT ON;

    SELECT
        [File Date] = CONVERT(DATE, dm.[FileDate]),
        [Contract] = dm.[Contract],
        [Contracts Traded] = dm.[ContractsTraded],
        [% Of Total Contracts Traded] = CONVERT(DECIMAL(18, 2), 100.0 * dm.[ContractsTraded] / SUM(dm.[ContractsTraded]) OVER ())
    FROM
        [DailyMTMDB].[dbo].[DailyMTM] dm
    WHERE
        dm.[FileDate] BETWEEN @DateFrom AND @DateTo;

END;
GO


