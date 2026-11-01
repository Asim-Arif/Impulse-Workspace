using System.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;

namespace DataAccessLibrary.DAC.FixedAssets;

public sealed class FixedAssetRegisterDataAccess : IFixedAssetRegisterDataAccess
{
    private readonly string _connectionString;

    public FixedAssetRegisterDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<FixedAssetRegisterFilterOptions> GetFilterOptionsAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        var result = new FixedAssetRegisterFilterOptions();

        // 1. Departments
        await using (var cmd = new SqlCommand("SELECT DeptID, ISNULL(Name,''), 0 FROM dbo.Departments WHERE ISNULL(Active,1)=1 ORDER BY Name", connection))
        await using (var reader = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                result.Departments.Add(new FixedAssetLookupItem
                {
                    Id = ReadString(reader, 0),
                    Name = ReadString(reader, 1)
                });
            }
        }

        // 2. Categories
        await using (var cmd = new SqlCommand("SELECT EntryID, ISNULL(Description,''), 0 FROM dbo.FixedAssets_Master ORDER BY Description ASC", connection))
        await using (var reader = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                result.Categories.Add(new FixedAssetLookupItem
                {
                    Id = ReadInt(reader, 0).ToString(),
                    Name = ReadString(reader, 1)
                });
            }
        }

        // 3. Last saved depreciation info
        DateTime? maxDtFrom = null;
        DateTime? maxDtTo = null;
        await using (var cmd = new SqlCommand("SELECT TOP 1 EntryID, DTFrom, DTTo FROM dbo.FixedAssets_Depreciation_Main ORDER BY EntryID DESC", connection))
        await using (var reader = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            if (await reader.ReadAsync(cancellationToken))
            {
                maxDtFrom = ReadDate(reader, 1);
                maxDtTo = ReadDate(reader, 2);
            }
        }

        result.LastSavedDateFrom = maxDtFrom;
        result.LastSavedDateTo = maxDtTo;

        if (maxDtTo.HasValue)
        {
            var nextDt = maxDtTo.Value.AddDays(1);
            result.DefaultDateFrom = new DateTime(nextDt.Year, nextDt.Month, 1);
            result.DefaultDateTo = new DateTime(nextDt.Year, nextDt.Month, DateTime.DaysInMonth(nextDt.Year, nextDt.Month));
        }
        else
        {
            result.DefaultDateFrom = new DateTime(2012, 7, 1);
            var lastMonth = DateTime.Today.AddDays(-DateTime.Today.Day);
            result.DefaultDateTo = new DateTime(lastMonth.Year, lastMonth.Month, DateTime.DaysInMonth(lastMonth.Year, lastMonth.Month));
        }

        // 4. Fill Periods
        int lastYear = maxDtFrom.HasValue ? maxDtFrom.Value.Year + 1 : DateTime.Today.Year;
        if (lastYear < 2001) lastYear = 2001;
        for (int yr = 2001; yr <= lastYear; yr++)
        {
            result.Periods.Add(new FixedAssetRegisterPeriod
            {
                DisplayText = $"{yr} - {yr + 1}",
                DateFrom = new DateTime(yr, 7, 1),
                DateTo = new DateTime(yr + 1, 6, 30)
            });
        }

        return result;
    }

    public async Task<(bool isSaved, List<FixedAssetRegisterRow> rows)> LoadRegisterDataAsync(DateTime dateFrom, DateTime dateTo, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        // Check if saved
        bool isSaved = false;
        await using (var checkCmd = new SqlCommand("SELECT COUNT(*) FROM dbo.FixedAssets_Depreciation_Main WHERE DTFrom = @DTFrom", connection))
        {
            checkCmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
            var cnt = Convert.ToInt32(await checkCmd.ExecuteScalarAsync(cancellationToken));
            isSaved = cnt > 0;
        }

        if (isSaved)
        {
            return (true, await LoadFromSavedAsync(connection, dateFrom, cancellationToken));
        }
        else
        {
            return (false, await CalculateRegisterAsync(connection, dateFrom, dateTo, cancellationToken));
        }
    }

    private static async Task<List<FixedAssetRegisterRow>> LoadFromSavedAsync(SqlConnection connection, DateTime dateFrom, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT d.AssetID, f.AcquiredDT, ISNULL(f.Description,'') AS Description, ISNULL(f.Location,'') AS Location,
                   ISNULL(d.Cost, 0) AS Cost, ISNULL(d.BookValue, 0) AS BookValue,
                   ISNULL(f.Disposed, 0) AS Disposed, f.DTDisposal, ISNULL(d.DisposedValue, 0) AS DisposedValue,
                   ISNULL(d.OpeningBalance, 0) AS OpeningBalance, ISNULL(d.AccountingYear, 0) AS AccountingYear,
                   ISNULL(d.AccumulatedBalance, 0) AS AccumulatedBalance, ISNULL(d.WrittenDownValue, 0) AS WrittenDownValue,
                   ISNULL(d.Rate, 0) AS Rate, ISNULL(d.Adjustment_Opening, 0) AS Adjustment_Opening,
                   ISNULL(d.AdjustedValue, 0) AS AdjustedValue, ISNULL(f.DeptID, '') AS DeptID,
                   ISNULL(m.Description, '') AS MasterDescription, ISNULL(f.InActive, 0) AS InActive,
                   f.InActiveDT, f.RetirementDT, ISNULL(f.Disposed_Type, 0) AS Disposed_Type
            FROM dbo.FixedAssets_Depreciation_Detail d
            INNER JOIN dbo.FixedAssets_Depreciation_Main dm ON d.FA_D_M_RefID = dm.EntryID
            LEFT JOIN dbo.FixedAssets f ON d.AssetID = f.AssetID
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID = m.EntryID
            WHERE dm.DTFrom = @DTFrom
            ORDER BY CASE WHEN ISNUMERIC(d.AssetID)=1 THEN CAST(d.AssetID AS int) ELSE 2147483647 END, d.AssetID
            """;

        var list = new List<FixedAssetRegisterRow>();
        await using var cmd = new SqlCommand(sql, connection);
        cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);

        await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var disposed = ReadBool(reader, 6);
            var dtDisposal = ReadDate(reader, 7);

            list.Add(new FixedAssetRegisterRow
            {
                AssetId = ReadString(reader, 0),
                AcquiredDate = ReadDate(reader, 1),
                Particulars = ReadString(reader, 2),
                Location = ReadString(reader, 3),
                Cost = ReadDecimal(reader, 4),
                BookValue = ReadDecimal(reader, 5),
                Disposed = disposed,
                DisposalDate = dtDisposal,
                DisposalDateText = (disposed && dtDisposal.HasValue) ? dtDisposal.Value.ToString("dd-MMM-yyyy") : "-",
                DisposedValue = ReadDecimal(reader, 8),
                OpeningBalance = ReadDecimal(reader, 9),
                AccountingYear = ReadDecimal(reader, 10),
                AccumulatedBalance = ReadDecimal(reader, 11),
                WrittenDownValue = ReadDecimal(reader, 12),
                Rate = ReadDecimal(reader, 13),
                AdjustmentOpening = ReadDecimal(reader, 14),
                AdjustedAmount = ReadDecimal(reader, 15),
                DeptId = ReadString(reader, 16),
                MasterDescription = ReadString(reader, 17),
                InActive = ReadBool(reader, 18),
                InActiveDate = ReadDate(reader, 19),
                RetirementDate = ReadDate(reader, 20),
                DisposedType = ReadInt(reader, 21)
            });
        }

        return list;
    }

    private static async Task<List<FixedAssetRegisterRow>> CalculateRegisterAsync(SqlConnection connection, DateTime dateFrom, DateTime dateTo, CancellationToken cancellationToken)
    {
        // 1. Determine Total Days of Year & myDTFrom
        double iTotalDaysOfYear;
        DateTime myDTFrom;
        if (dateTo.Month <= 6)
        {
            iTotalDaysOfYear = (new DateTime(dateTo.Year, 6, 30) - new DateTime(dateTo.Year - 1, 7, 1)).TotalDays + 1;
            myDTFrom = new DateTime(dateFrom.Year - 1, 7, 1);
        }
        else
        {
            iTotalDaysOfYear = (new DateTime(dateTo.Year + 1, 6, 30) - new DateTime(dateTo.Year, 7, 1)).TotalDays + 1;
            myDTFrom = new DateTime(dateFrom.Year, 7, 1);
        }

        // 2. Pre-load cached maps to avoid N+1 query execution
        var accDepMap = new Dictionary<int, decimal>();
        await using (var cmd = new SqlCommand("SELECT EntryID, ISNULL(SUM(CAST(AccountingYear AS decimal(18,2))),0) FROM dbo.VFixedAssets_Depreciation_Detail WHERE DTFrom < @DTFrom GROUP BY EntryID", connection))
        {
            cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
            await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                accDepMap[ReadInt(reader, 0)] = ReadDecimal(reader, 1);
            }
        }

        var adjOpenMap = new Dictionary<string, (decimal sum, DateTime maxDt)>(StringComparer.OrdinalIgnoreCase);
        await using (var cmd = new SqlCommand("SELECT AssetID, ISNULL(SUM(CAST(Adjustment AS decimal(18,2))),0), MAX(DT) FROM dbo.FixedAssets_Adjustment WHERE DT BETWEEN @DTFrom AND @DTTo GROUP BY AssetID", connection))
        {
            cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
            cmd.Parameters.AddWithValue("@DTTo", dateTo.Date);
            await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                var assetId = ReadString(reader, 0);
                var sum = ReadDecimal(reader, 1);
                var maxDt = ReadDate(reader, 2) ?? DateTime.MinValue;
                adjOpenMap[assetId] = (sum, maxDt);
            }
        }

        var adjPriorMap = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        await using (var cmd = new SqlCommand("SELECT AssetID, ISNULL(SUM(CAST(Adjustment AS decimal(18,2))),0) FROM dbo.FixedAssets_Adjustment WHERE DT < @DTFrom GROUP BY AssetID", connection))
        {
            cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
            await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                adjPriorMap[ReadString(reader, 0)] = ReadDecimal(reader, 1);
            }
        }

        var latestAccBalMap = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        const string latestBalSql = """
            SELECT AssetID, ISNULL(SUM(CAST(AccumulatedBalance AS decimal(18,2))),0)
            FROM dbo.VFixedAssets_Depreciation_Detail d
            WHERE DTFrom IN (SELECT MAX(DTFrom) FROM dbo.VFixedAssets_Depreciation_Detail WHERE AssetID = d.AssetID)
            GROUP BY AssetID
            """;
        await using (var cmd = new SqlCommand(latestBalSql, connection))
        await using (var readerBal = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await readerBal.ReadAsync(cancellationToken))
            {
                latestAccBalMap[ReadString(readerBal, 0)] = ReadDecimal(readerBal, 1);
            }
        }

        var masterRateMap = new Dictionary<int, decimal>();
        await using (var cmd = new SqlCommand("SELECT FA_M_RefID, ISNULL(PercentRate,0) FROM dbo.FixedAssets_Master_DepreciationRate WHERE DTFrom = @MyDTFrom", connection))
        {
            cmd.Parameters.AddWithValue("@MyDTFrom", myDTFrom.Date);
            await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                masterRateMap[ReadInt(reader, 0)] = ReadDecimal(reader, 1);
            }
        }

        // 3. Query all assets acquired on or before DTTo
        const string queryAssetsSql = """
            SELECT f.EntryID, ISNULL(f.AssetID,'') AS AssetID, ISNULL(f.Description,'') AS Description,
                   f.AcquiredDT, ISNULL(f.Location,'') AS Location, ISNULL(CAST(f.AcquisitionCost AS decimal(18,2)),0) AS AcquisitionCost,
                   ISNULL(f.Disposed,0) AS Disposed, f.DTDisposal, ISNULL(f.Disposed_Type,0) AS Disposed_Type,
                   ISNULL(f.InActive,0) AS InActive, f.InActiveDT, f.RetirementDT,
                   ISNULL(CAST(f.AcumulatedDepriciation AS decimal(18,2)),0) AS AcumulatedDepriciation,
                   ISNULL(f.MasterFA_RefID,0) AS MasterFA_RefID, ISNULL(m.Description,'') AS MasterDescription,
                   ISNULL(CAST(m.PercentRate AS decimal(18,2)),0) AS MasterPercentRate,
                   ISNULL(f.DeptID,'') AS DeptID, ISNULL(d.name,'') AS DeptName
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID = m.EntryID
            LEFT JOIN dbo.Departments d ON f.DeptID = d.deptid
            WHERE f.AcquiredDT <= @DTTo
            ORDER BY CASE WHEN ISNUMERIC(f.AssetID)=1 THEN CAST(f.AssetID AS int) ELSE 2147483647 END, f.AssetID
            """;

        var list = new List<FixedAssetRegisterRow>();
        await using var mainCmd = new SqlCommand(queryAssetsSql, connection);
        mainCmd.Parameters.AddWithValue("@DTTo", dateTo.Date);

        await using var readerAssets = await mainCmd.ExecuteReaderAsync(cancellationToken);
        while (await readerAssets.ReadAsync(cancellationToken))
        {
            var entryId = ReadInt(readerAssets, 0);
            var assetId = ReadString(readerAssets, 1);
            var description = ReadString(readerAssets, 2);
            var acquiredDt = ReadDate(readerAssets, 3);
            var location = ReadString(readerAssets, 4);
            var acquisitionCost = ReadDecimal(readerAssets, 5);
            var disposed = ReadBool(readerAssets, 6);
            var dtDisposal = ReadDate(readerAssets, 7);
            var disposedType = ReadInt(readerAssets, 8);
            var inActive = ReadBool(readerAssets, 9);
            var inActiveDt = ReadDate(readerAssets, 10);
            var retirementDt = ReadDate(readerAssets, 11);
            var acumulatedDepriciation = ReadDecimal(readerAssets, 12);
            var masterFaRefId = ReadInt(readerAssets, 13);
            var masterDescription = ReadString(readerAssets, 14);
            var masterPercentRate = ReadDecimal(readerAssets, 15);
            var deptId = ReadString(readerAssets, 16);
            var deptName = ReadString(readerAssets, 17);

            // Calculations
            decimal dAccDep = accDepMap.GetValueOrDefault(entryId, 0);

            decimal dAdjustment_Opening = 0;
            DateTime dtAdjustment = DateTime.MinValue;
            if (adjOpenMap.TryGetValue(assetId, out var adjTuple))
            {
                dAdjustment_Opening = adjTuple.sum;
                dtAdjustment = adjTuple.maxDt;
            }

            decimal dAdjustedAmount = adjPriorMap.GetValueOrDefault(assetId, 0);

            bool bDisposed = false;
            bool bAlreadyDisposed = false;
            double dDisposalDays = 0;
            DateTime? effectiveRetirementDt = retirementDt;

            if (disposed)
            {
                if (dtDisposal.HasValue && dtDisposal.Value >= dateFrom && dtDisposal.Value <= dateTo)
                {
                    bDisposed = true;
                    if (inActive && inActiveDt.HasValue)
                    {
                        dDisposalDays = (inActiveDt.Value - dateFrom).TotalDays + 1;
                        effectiveRetirementDt = inActiveDt;
                    }
                    else
                    {
                        if (effectiveRetirementDt.HasValue && effectiveRetirementDt.Value < dtDisposal.Value)
                        {
                            dDisposalDays = (effectiveRetirementDt.Value - dateFrom).TotalDays + 1;
                        }
                        else
                        {
                            dDisposalDays = (dtDisposal.Value - dateFrom).TotalDays + 1;
                        }
                    }
                    if (dDisposalDays < 0) dDisposalDays = 0;
                }
                else if (dtDisposal.HasValue && dtDisposal.Value < dateFrom)
                {
                    bAlreadyDisposed = true;
                }
            }
            else if (inActive && inActiveDt.HasValue)
            {
                dDisposalDays = (inActiveDt.Value - dateFrom).TotalDays + 1;
                effectiveRetirementDt = inActiveDt;
            }

            int iDisposedType;
            if (disposed && disposedType == 1)
                iDisposedType = 0; // Disposed
            else if (disposed && disposedType == 0)
                iDisposedType = (dateFrom >= new DateTime(2025, 7, 1) ? 0 : 1); // Write Off
            else
                iDisposedType = 2;

            decimal dAcqCost;
            decimal dBookValue;
            string strDisDT;
            decimal dDisCost;
            decimal dDisAccDep;

            if (bDisposed)
            {
                dAcqCost = acquisitionCost + dAdjustedAmount;
                dBookValue = 0;
                strDisDT = dtDisposal?.ToString("dd-MMM-yyyy") ?? "-";
                dDisCost = acquisitionCost;
                dDisAccDep = dAccDep;
                dAccDep = 0;
            }
            else
            {
                dAcqCost = acquisitionCost + dAdjustedAmount;
                dBookValue = Math.Max(0, dAcqCost - dAccDep);
                strDisDT = "-";
                dDisCost = 0;
                dDisAccDep = 0;
            }

            decimal dPercentRate = masterRateMap.GetValueOrDefault(masterFaRefId, masterPercentRate);

            decimal dAddAdjustment = 0;
            decimal dBookvalue_Adjustment = 0;
            decimal dBookvalue_Adjusted = 0;

            if (dAdjustment_Opening != 0 && acquiredDt.HasValue)
            {
                double dDaysAdj = (dateTo - acquiredDt.Value).TotalDays + 1;
                dAddAdjustment = Math.Round((((dAdjustment_Opening * dPercentRate) / 100m) / (decimal)iTotalDaysOfYear), 2);
                dBookvalue_Adjustment = dAddAdjustment * (decimal)dDaysAdj;
                dBookvalue_Adjusted = dAdjustment_Opening - dBookvalue_Adjustment;
                dAccDep += dBookvalue_Adjustment;
            }
            else
            {
                dAddAdjustment = 0;
                dBookvalue_Adjustment = 0;
                dBookvalue_Adjusted = 0;
                dAccDep = 0;
            }

            decimal dAddDep = 0;
            decimal dDisDep = 0;
            decimal dOpenDep = 0;
            decimal dAdditions;

            if (acquiredDt.HasValue && acquiredDt.Value > new DateTime(2005, 6, 30))
            {
                dAdditions = dAcqCost;
                double dDaysCalc;
                if (acquiredDt.Value < dateFrom)
                {
                    var rDate = effectiveRetirementDt ?? dateTo;
                    if (rDate < dateTo)
                        dDaysCalc = (rDate - dateFrom).TotalDays + 1;
                    else
                        dDaysCalc = (dateTo - dateFrom).TotalDays + 1;

                    if (dDaysCalc < 0) dDaysCalc = 0;
                }
                else
                {
                    dDaysCalc = (dateTo - acquiredDt.Value).TotalDays + 1;
                }

                if (bDisposed || bAlreadyDisposed)
                {
                    dDisDep = acumulatedDepriciation;
                    if (bDisposed && iDisposedType == 0 && (dtDisposal.HasValue && dtDisposal.Value >= dateFrom && dtDisposal.Value <= dateTo))
                    {
                        if (acquiredDt.Value >= dateFrom && acquiredDt.Value <= dateTo)
                        {
                            dDisposalDays = (dtDisposal.Value - acquiredDt.Value).TotalDays + 1;
                        }
                        dAddDep = Math.Round((((dAdditions * dPercentRate) / 100m) / (decimal)iTotalDaysOfYear) * (decimal)dDisposalDays, 2);
                    }
                }
                else
                {
                    if (acquiredDt.Value >= dateFrom && acquiredDt.Value <= dateTo)
                    {
                        dOpenDep = 0;
                    }
                    else
                    {
                        dOpenDep = Math.Round((((dAcqCost * dPercentRate) / 100m) / (decimal)iTotalDaysOfYear) * (decimal)dDaysCalc, 2);
                    }
                    dAddDep = Math.Round((((dAdditions * dPercentRate) / 100m) / (decimal)iTotalDaysOfYear) * (decimal)dDaysCalc, 2);
                }
            }
            else
            {
                dAdditions = 0;
                double dDaysCalc = (acquiredDt.HasValue && acquiredDt.Value < dateFrom)
                    ? (dateTo - dateFrom).TotalDays + 1
                    : (dateTo - (acquiredDt ?? dateFrom)).TotalDays + 1;

                if (bDisposed || bAlreadyDisposed)
                {
                    dDisDep = acumulatedDepriciation;
                }
                else
                {
                    dOpenDep = Math.Round((((dAcqCost * dPercentRate) / 100m) / (decimal)iTotalDaysOfYear) * (decimal)dDaysCalc, 2);
                }
                dAddDep = dOpenDep;
            }

            decimal dTotal = dAddDep + dDisDep;

            if (dAcqCost - dAccDep <= 0)
            {
                dTotal = 0;
                dAddDep = 0;
            }
            if (dAcqCost - dAccDep - dTotal < 0)
            {
                dTotal = dAcqCost - dAccDep;
                dAddDep = dAcqCost - dAccDep;
            }
            if (bDisposed)
            {
                dTotal = 0;
                dAccDep = 0;
            }

            decimal colOpening = 0;
            if (bDisposed || bAlreadyDisposed)
            {
                if (strDisDT != "-")
                {
                    if (dtDisposal.HasValue && dtDisposal.Value.Year >= dateFrom.Year && dtDisposal.Value.Year <= dateTo.Year)
                        colOpening = dDisAccDep;
                    else
                        colOpening = (bDisposed || bAlreadyDisposed) ? dDisDep : (dAccDep + dTotal);
                }
                else
                {
                    colOpening = latestAccBalMap.GetValueOrDefault(assetId, 0);
                }
            }
            else
            {
                if (acquiredDt.HasValue && acquiredDt.Value >= dateFrom && acquiredDt.Value <= dateTo && !bDisposed)
                {
                    colOpening = 0;
                }
                else
                {
                    var maxAccBal = latestAccBalMap.GetValueOrDefault(assetId, 0);
                    colOpening = (bDisposed || bAlreadyDisposed) ? dDisDep : maxAccBal;
                    dAccDep = maxAccBal;
                }
            }

            decimal dAmount_Extra = 0;
            if (disposed)
            {
                dAmount_Extra = (dDisAccDep + dAddDep) - (dAcqCost + dBookvalue_Adjustment);
                if (dAmount_Extra < 0)
                {
                    if (dDisAccDep <= 0)
                        dAmount_Extra = (dAccDep + dAddDep) - (dAcqCost + dBookvalue_Adjustment);
                }
            }
            else
            {
                dAmount_Extra = (dAccDep + dAddDep) - (dAcqCost + dBookvalue_Adjustment);
            }
            if (dAmount_Extra < 0) dAmount_Extra = 0;

            if (dAmount_Extra > 0)
            {
                dTotal = Math.Max(0, (dAddDep + dBookvalue_Adjustment) - dAmount_Extra);
            }
            else
            {
                dTotal = dAddDep + dBookvalue_Adjustment;
            }

            decimal colAccYear = (dAmount_Extra > 0)
                ? Math.Max(0, (dAddDep + dBookvalue_Adjustment) - dAmount_Extra)
                : (dAddDep + dBookvalue_Adjustment);

            decimal colBalance = (bDisposed || bAlreadyDisposed)
                ? dDisDep
                : (dAccDep + dTotal);

            decimal colWDValue;
            if (bDisposed || bAlreadyDisposed)
            {
                colWDValue = 0;
            }
            else
            {
                colWDValue = Math.Max(0, dAcqCost - ((dAccDep) + dTotal - dAdjustment_Opening));
            }

            list.Add(new FixedAssetRegisterRow
            {
                AssetId = assetId,
                AcquiredDate = acquiredDt,
                Particulars = description,
                Location = location,
                Cost = dAcqCost + dAdjustment_Opening,
                BookValue = dBookValue + dBookvalue_Adjusted,
                Disposed = disposed,
                DisposalDate = dtDisposal,
                DisposalDateText = strDisDT,
                DisposedValue = dDisCost - dDisAccDep,
                OpeningBalance = colOpening,
                AccountingYear = colAccYear,
                AccumulatedBalance = colBalance,
                WrittenDownValue = colWDValue,
                Rate = dPercentRate,
                AdjustmentOpening = dAdjustment_Opening,
                AdjustedAmount = dBookvalue_Adjustment,
                DeptId = deptId,
                DeptName = deptName,
                MasterId = masterFaRefId,
                MasterDescription = masterDescription,
                InActive = inActive,
                InActiveDate = inActiveDt,
                RetirementDate = retirementDt,
                DisposedType = disposedType
            });
        }

        return list;
    }

    public async Task<FixedAssetRegisterSaveResult> SaveRegisterAsync(DateTime dateFrom, DateTime dateTo, IReadOnlyList<FixedAssetRegisterRow> rows, string userName, string machineName, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = connection.BeginTransaction();

        try
        {
            // Insert Main Header with explicit EntryID
            const string insertMainSql = """
                DECLARE @NewEntryID int;
                SELECT @NewEntryID = ISNULL(MAX(EntryID), 0) + 1 FROM dbo.FixedAssets_Depreciation_Main;
                INSERT INTO dbo.FixedAssets_Depreciation_Main (EntryID, DTFrom, DTTo, UserName, MachineName, DTEntry)
                VALUES (@NewEntryID, @DTFrom, @DTTo, @UserName, @MachineName, GETDATE());
                SELECT @NewEntryID;
                """;

            int entryId;
            await using (var cmd = new SqlCommand(insertMainSql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
                cmd.Parameters.AddWithValue("@DTTo", dateTo.Date);
                cmd.Parameters.AddWithValue("@UserName", userName ?? "Administrator");
                cmd.Parameters.AddWithValue("@MachineName", machineName ?? Environment.MachineName);
                entryId = Convert.ToInt32(await cmd.ExecuteScalarAsync(cancellationToken));
            }

            // Get initial Detail EntryID sequence
            int nextDetailEntryId;
            await using (var getDetailIdCmd = new SqlCommand("SELECT ISNULL(MAX(EntryID), 0) FROM dbo.FixedAssets_Depreciation_Detail", connection, transaction))
            {
                nextDetailEntryId = Convert.ToInt32(await getDetailIdCmd.ExecuteScalarAsync(cancellationToken));
            }

            // Insert Detail Rows
            const string insertDetailSql = """
                INSERT INTO dbo.FixedAssets_Depreciation_Detail
                (EntryID, FA_D_M_RefID, AssetID, Cost, BookValue, DisposedValue, OpeningBalance, AccountingYear, AccumulatedBalance, WrittenDownValue, Rate, Adjustment_Opening, AdjustedValue)
                VALUES
                (@EntryID, @FA_D_M_RefID, @AssetID, @Cost, @BookValue, @DisposedValue, @OpeningBalance, @AccountingYear, @AccumulatedBalance, @WrittenDownValue, @Rate, @Adjustment_Opening, @AdjustedValue)
                """;

            foreach (var row in rows)
            {
                nextDetailEntryId++;
                await using var cmd = new SqlCommand(insertDetailSql, connection, transaction);
                cmd.Parameters.AddWithValue("@EntryID", nextDetailEntryId);
                cmd.Parameters.AddWithValue("@FA_D_M_RefID", entryId);
                cmd.Parameters.AddWithValue("@AssetID", row.AssetId ?? "");
                cmd.Parameters.AddWithValue("@Cost", row.Cost);
                cmd.Parameters.AddWithValue("@BookValue", row.BookValue);
                cmd.Parameters.AddWithValue("@DisposedValue", row.DisposedValue);
                cmd.Parameters.AddWithValue("@OpeningBalance", row.OpeningBalance);
                cmd.Parameters.AddWithValue("@AccountingYear", row.AccountingYear);
                cmd.Parameters.AddWithValue("@AccumulatedBalance", row.AccumulatedBalance);
                cmd.Parameters.AddWithValue("@WrittenDownValue", row.WrittenDownValue);
                cmd.Parameters.AddWithValue("@Rate", row.Rate);
                cmd.Parameters.AddWithValue("@Adjustment_Opening", row.AdjustmentOpening);
                cmd.Parameters.AddWithValue("@AdjustedValue", row.AdjustedAmount);
                await cmd.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return new FixedAssetRegisterSaveResult
            {
                Success = true,
                EntryId = entryId,
                Message = $"Depreciation Register for period {dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy} saved successfully. Click Post Voucher to load JV rows."
            };
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync(cancellationToken);
            return new FixedAssetRegisterSaveResult
            {
                Success = false,
                Message = ex.Message
            };
        }
    }

    public async Task<(bool success, string message)> DeleteRegisterAsync(DateTime dateFrom, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        // Check if DTFrom is the latest one
        DateTime? maxDtFrom = null;
        await using (var cmd = new SqlCommand("SELECT MAX(DTFrom) FROM dbo.FixedAssets_Depreciation_Main", connection))
        {
            var val = await cmd.ExecuteScalarAsync(cancellationToken);
            if (val != null && val != DBNull.Value)
            {
                maxDtFrom = Convert.ToDateTime(val);
            }
        }

        if (maxDtFrom.HasValue && maxDtFrom.Value.Date != dateFrom.Date)
        {
            return (false, "Can only Delete the Last One.");
        }

        await using var transaction = connection.BeginTransaction();
        try
        {
            const string deleteDetailSql = """
                DELETE FROM dbo.FixedAssets_Depreciation_Detail
                WHERE FA_D_M_RefID IN (SELECT EntryID FROM dbo.FixedAssets_Depreciation_Main WHERE DTFrom = @DTFrom)
                """;
            await using (var cmd = new SqlCommand(deleteDetailSql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
                await cmd.ExecuteNonQueryAsync(cancellationToken);
            }

            const string deleteMainSql = "DELETE FROM dbo.FixedAssets_Depreciation_Main WHERE DTFrom = @DTFrom";
            await using (var cmd = new SqlCommand(deleteMainSql, connection, transaction))
            {
                cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
                await cmd.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return (true, "Register deleted successfully.");
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync(cancellationToken);
            return (false, ex.Message);
        }
    }

    public async Task<int?> GetVoucherRefIdAsync(DateTime dateFrom, DateTime dateTo, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = "SELECT TOP 1 FA_D_M_RefID FROM dbo.VFixedAssets_Depreciation_Detail WHERE DTFrom = @DTFrom AND DTTo = @DTTo";
        await using var cmd = new SqlCommand(sql, connection);
        cmd.Parameters.AddWithValue("@DTFrom", dateFrom.Date);
        cmd.Parameters.AddWithValue("@DTTo", dateTo.Date);

        var val = await cmd.ExecuteScalarAsync(cancellationToken);
        if (val != null && val != DBNull.Value && int.TryParse(val.ToString(), out var refId))
        {
            return refId;
        }

        return null;
    }

    private static string ReadString(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? "" : Convert.ToString(reader.GetValue(ordinal)) ?? "";
    private static int ReadInt(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0 : Convert.ToInt32(reader.GetValue(ordinal));
    private static int? ReadNullableInt(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal));
    private static DateTime? ReadDate(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToDateTime(reader.GetValue(ordinal));
    private static decimal ReadDecimal(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0m : Convert.ToDecimal(reader.GetValue(ordinal));
    private static bool ReadBool(SqlDataReader reader, int ordinal) => !reader.IsDBNull(ordinal) && Convert.ToBoolean(reader.GetValue(ordinal));
}
