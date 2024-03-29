USE [JSE]
GO
/****** Object:  UserDefinedTableType [dbo].[DailyMTMType]    Script Date: 2023/12/18 08:34:27 ******/
CREATE TYPE [dbo].[DailyMTMType] AS TABLE(
	[FileDate] [date] NULL,
	[Contract] [nvarchar](50) NULL,
	[ExpiryDate] [date] NULL,
	[Classification] [nvarchar](50) NULL,
	[Strike] [float] NULL,
	[CallPut] [nvarchar](50) NULL,
	[MTMYield] [float] NULL,
	[MarkPrice] [float] NULL,
	[SpotRate] [float] NULL,
	[PreviousMTM] [float] NULL,
	[PreviousPrice] [float] NULL,
	[PremiumOnOption] [float] NULL,
	[Volatility] [float] NULL,
	[Delta] [float] NULL,
	[DeltaValue] [float] NULL,
	[ContractsTraded] [float] NULL,
	[OpenInterest] [float] NULL
)
GO
/****** Object:  Table [dbo].[DailyMTM]    Script Date: 2023/12/18 08:34:27 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DailyMTM](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[FileDate] [date] NOT NULL,
	[Contract] [nvarchar](50) NOT NULL,
	[ExpiryDate] [date] NOT NULL,
	[Classification] [nvarchar](50) NOT NULL,
	[Strike] [float] NOT NULL,
	[CallPut] [nvarchar](50) NULL,
	[MTMYield] [float] NOT NULL,
	[MarkPrice] [float] NOT NULL,
	[SpotRate] [float] NOT NULL,
	[PreviousMTM] [float] NOT NULL,
	[PreviousPrice] [float] NOT NULL,
	[PremiumOnOption] [float] NULL,
	[Volatility] [float] NOT NULL,
	[Delta] [float] NOT NULL,
	[DeltaValue] [float] NOT NULL,
	[ContractsTraded] [float] NOT NULL,
	[OpenInterest] [float] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  StoredProcedure [dbo].[InsertDailyMTMData]    Script Date: 2023/12/18 08:34:27 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[InsertDailyMTMData]
    @DailyMTMData dbo.DailyMTMType READONLY
AS
BEGIN
    SET NOCOUNT ON;

    -- Check for duplicates
    MERGE INTO dbo.DailyMTM AS target
    USING (SELECT * FROM @DailyMTMData) AS source
    ON target.FileDate = source.FileDate 
		AND target.ExpiryDate = source.ExpiryDate
		AND target.MTMYield = source.MTMYield
		AND target.MarkPrice = source.MarkPrice
		AND target.PreviousMTM = source.PreviousMTM
		AND target.PreviousPrice = source.PreviousPrice

    WHEN NOT MATCHED THEN
        INSERT (
			FileDate
			, Contract
			, ExpiryDate, Classification
			, Strike
			, CallPut
			, MTMYield
			, MarkPrice
			, SpotRate
			, PreviousMTM
			, PreviousPrice
			, PremiumOnOption
			, Volatility
			, Delta
			, DeltaValue
			, ContractsTraded
			, OpenInterest
			)
        VALUES (
			source.FileDate
			, source.Contract
			, source.ExpiryDate
			, source.Classification
			, source.Strike
			, source.CallPut
			, source.MTMYield
			, source.MarkPrice
			, source.SpotRate
			, source.PreviousMTM
			, source.PreviousPrice
			, source.PremiumOnOption
			, source.Volatility
			, source.Delta
			, source.DeltaValue
			, source.ContractsTraded
			, source.OpenInterest
			);
END;
GO
/****** Object:  StoredProcedure [dbo].[SP_Total_Contracts_Traded_Report]    Script Date: 2023/12/18 08:34:27 ******/
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

    -- Calculate total contracts traded for each day in the given date range
    WITH TotalDailyTrades AS (
        SELECT [FileDate], SUM([ContractsTraded]) AS TotalContracts
        FROM [dbo].[DailyMTM]
        WHERE [FileDate] BETWEEN @DateFrom AND @DateTo
        GROUP BY [FileDate]
    )
    SELECT 
        dtm.[FileDate],
        dtm.[Contract],
        dtm.[ContractsTraded],
        CASE 
            WHEN tdt.TotalContracts > 0 THEN (dtm.[ContractsTraded] / tdt.TotalContracts) * 100
            ELSE 0
        END AS [% Of Total Contracts Traded]
    FROM [dbo].[DailyMTM] dtm
    INNER JOIN TotalDailyTrades tdt ON dtm.[FileDate] = tdt.[FileDate]
    WHERE dtm.[ContractsTraded] > 0
    AND dtm.[FileDate] BETWEEN @DateFrom AND @DateTo
    ORDER BY dtm.[FileDate], dtm.[Contract];
END;
GO
