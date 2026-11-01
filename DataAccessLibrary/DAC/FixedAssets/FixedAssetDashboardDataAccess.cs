using System.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;

namespace DataAccessLibrary.DAC.FixedAssets;

public sealed class FixedAssetDashboardDataAccess : IFixedAssetDashboardDataAccess
{
    private readonly string _connectionString;

    private static readonly string[] Palette = new[]
    {
        "#3b82f6", "#10b981", "#f59e0b", "#8b5cf6", "#ec4899",
        "#06b6d4", "#14b8a6", "#f97316", "#6366f1", "#84cc16",
        "#e11d48", "#0ea5e9", "#a855f7", "#d97706", "#64748b"
    };

    private static readonly string[] BadgeClasses = new[]
    {
        "primary", "success", "warning", "info", "danger", "secondary"
    };

    public FixedAssetDashboardDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<FixedAssetDashboardResult> GetDashboardDataAsync(FixedAssetDashboardFilter filter, CancellationToken cancellationToken = default)
    {
        var result = new FixedAssetDashboardResult();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        // 1. Lookups
        result.LookupCategories = await LoadLookupAsync(connection, "SELECT EntryID, ISNULL(Description,''), ISNULL(CAST(PercentRate AS decimal(18,2)),0) FROM dbo.FixedAssets_Master ORDER BY Description", cancellationToken);
        result.LookupDepartments = await LoadLookupAsync(connection, "SELECT DeptID, ISNULL(Name,''), 0 FROM dbo.Departments WHERE ISNULL(Active,1)=1 ORDER BY Name", cancellationToken);

        // 2. Fiscal years list
        var fiscalYears = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        await using (var cmd = new SqlCommand("SELECT DISTINCT CASE WHEN AcquiredDT IS NULL THEN 'Unknown' WHEN MONTH(AcquiredDT) >= 7 THEN CAST(YEAR(AcquiredDT) AS varchar) + '-' + CAST(YEAR(AcquiredDT)+1 AS varchar) ELSE CAST(YEAR(AcquiredDT)-1 AS varchar) + '-' + CAST(YEAR(AcquiredDT) AS varchar) END AS FY FROM dbo.FixedAssets WHERE AcquiredDT IS NOT NULL ORDER BY FY DESC", connection))
        await using (var r = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await r.ReadAsync(cancellationToken))
            {
                var fy = ReadString(r, 0);
                if (!string.IsNullOrWhiteSpace(fy) && fy != "Unknown") fiscalYears.Add(fy);
            }
        }
        var curFy = FixedAssetDashboardFilter.GetCurrentFinancialYear();
        if (!fiscalYears.Contains(curFy)) fiscalYears.Add(curFy);
        result.FiscalYears = fiscalYears.OrderByDescending(x => x).ToList();

        // Base WHERE clause
        var whereParts = new List<string>();
        var parameters = new List<SqlParameter>();

        if (filter.MasterId.HasValue && filter.MasterId.Value > 0)
        {
            whereParts.Add("f.MasterFA_RefID = @MasterId");
            parameters.Add(new SqlParameter("@MasterId", filter.MasterId.Value));
        }

        if (!string.IsNullOrWhiteSpace(filter.DepartmentId))
        {
            whereParts.Add("f.DeptID = @DeptId");
            parameters.Add(new SqlParameter("@DeptId", filter.DepartmentId));
        }

        if (!string.IsNullOrWhiteSpace(filter.SearchText))
        {
            whereParts.Add("(f.AssetID LIKE @Like OR f.Description LIKE @Like OR f.Location LIKE @Like)");
            parameters.Add(new SqlParameter("@Like", $"%{filter.SearchText.Trim()}%"));
        }

        if (!string.IsNullOrWhiteSpace(filter.FiscalYear) && filter.FiscalYear != "All")
        {
            whereParts.Add("(CASE WHEN f.AcquiredDT IS NULL THEN '' WHEN MONTH(f.AcquiredDT) >= 7 THEN CAST(YEAR(f.AcquiredDT) AS varchar) + '-' + CAST(YEAR(f.AcquiredDT)+1 AS varchar) ELSE CAST(YEAR(f.AcquiredDT)-1 AS varchar) + '-' + CAST(YEAR(f.AcquiredDT) AS varchar) END = @FiscalYear)");
            parameters.Add(new SqlParameter("@FiscalYear", filter.FiscalYear));
        }

        if (filter.QuarterIndex >= 1 && filter.QuarterIndex <= 4)
        {
            var months = filter.QuarterIndex switch
            {
                1 => "7, 8, 9",
                2 => "10, 11, 12",
                3 => "1, 2, 3",
                _ => "4, 5, 6"
            };
            whereParts.Add($"(f.AcquiredDT IS NOT NULL AND MONTH(f.AcquiredDT) IN ({months}))");
        }

        if (!string.IsNullOrWhiteSpace(filter.Status) && filter.Status != "All")
        {
            if (filter.Status.Equals("Active", StringComparison.OrdinalIgnoreCase))
                whereParts.Add("ISNULL(f.InActive,0)=0 AND ISNULL(f.Disposed,0)=0");
            else if (filter.Status.Equals("Disposed", StringComparison.OrdinalIgnoreCase))
                whereParts.Add("ISNULL(f.Disposed,0)=1");
            else if (filter.Status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                whereParts.Add("ISNULL(f.InActive,0)=1");
        }

        var whereClause = whereParts.Count > 0 ? "WHERE " + string.Join(" AND ", whereParts) : "";

        // 3. Summary Aggregations
        var summarySql = $@"
            SELECT 
                COUNT(*) AS TotalAssets,
                SUM(CASE WHEN ISNULL(f.InActive,0)=0 AND ISNULL(f.Disposed,0)=0 THEN 1 ELSE 0 END) AS ActiveAssets,
                SUM(CASE WHEN ISNULL(f.Disposed,0)=1 THEN 1 ELSE 0 END) AS DisposedAssets,
                SUM(CASE WHEN ISNULL(f.InActive,0)=1 THEN 1 ELSE 0 END) AS InactiveAssets,
                ISNULL(SUM(CAST(f.AcquisitionCost AS decimal(18,2))),0) AS TotalCost,
                ISNULL(SUM(CAST(f.AcumulatedDepriciation AS decimal(18,2))),0) AS TotalAccDep,
                ISNULL(SUM(CAST(f.SalePrice AS decimal(18,2))),0) AS TotalDisposalProceeds,
                ISNULL(SUM(CAST(ISNULL(f.SalePrice,0) - ISNULL(f.WDV,0) AS decimal(18,2))),0) AS TotalDisposalGainLoss
            FROM dbo.FixedAssets f
            {whereClause};
        ";

        await using (var cmd = new SqlCommand(summarySql, connection))
        {
            foreach (var p in parameters) cmd.Parameters.Add(new SqlParameter(p.ParameterName, p.Value));
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            if (await r.ReadAsync(cancellationToken))
            {
                result.Summary.TotalAssets = ReadInt(r, 0);
                result.Summary.ActiveAssets = ReadInt(r, 1);
                result.Summary.DisposedAssets = ReadInt(r, 2);
                result.Summary.InactiveAssets = ReadInt(r, 3);
                result.Summary.TotalAcquisitionCost = ReadDecimal(r, 4);
                result.Summary.TotalAccumulatedDepreciation = ReadDecimal(r, 5);
                result.Summary.TotalDisposalProceeds = ReadDecimal(r, 6);
                result.Summary.TotalDisposalGainLoss = ReadDecimal(r, 7);
            }
        }

        // Adjustments
        await using (var cmd = new SqlCommand("SELECT ISNULL(SUM(CAST(Adjustment AS decimal(18,2))),0) FROM dbo.FixedAssets_Adjustment", connection))
        {
            result.Summary.TotalAdjustments = ReadDecimal(await cmd.ExecuteScalarAsync(cancellationToken));
        }

        // Current FY additions
        var currentYear = DateTime.Today.Month >= 7 ? DateTime.Today.Year : DateTime.Today.Year - 1;
        var fyStart = new DateTime(currentYear, 7, 1);
        var fyEnd = new DateTime(currentYear + 1, 6, 30);
        await using (var cmd = new SqlCommand("SELECT COUNT(*), ISNULL(SUM(CAST(AcquisitionCost AS decimal(18,2))),0) FROM dbo.FixedAssets WHERE AcquiredDT >= @FyStart AND AcquiredDT <= @FyEnd", connection))
        {
            cmd.Parameters.AddWithValue("@FyStart", fyStart);
            cmd.Parameters.AddWithValue("@FyEnd", fyEnd);
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            if (await r.ReadAsync(cancellationToken))
            {
                result.Summary.CurrentYearAdditionsCount = ReadInt(r, 0);
                result.Summary.CurrentYearAdditionsCost = ReadDecimal(r, 1);
            }
        }

        // 4. Categories Breakdown
        var catSql = $@"
            SELECT 
                ISNULL(m.EntryID, 0) AS MasterId,
                ISNULL(m.Description, 'Uncategorized') AS CategoryName,
                COUNT(f.EntryID) AS AssetCount,
                ISNULL(SUM(CAST(f.AcquisitionCost AS decimal(18,2))), 0) AS TotalCost,
                ISNULL(SUM(CAST(f.AcumulatedDepriciation AS decimal(18,2))), 0) AS AccDep,
                ISNULL(AVG(CAST(m.PercentRate AS decimal(18,2))), 0) AS DepRate
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID = m.EntryID
            {whereClause}
            GROUP BY m.EntryID, m.Description
            ORDER BY TotalCost DESC;
        ";

        await using (var cmd = new SqlCommand(catSql, connection))
        {
            foreach (var p in parameters) cmd.Parameters.Add(new SqlParameter(p.ParameterName, p.Value));
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            int idx = 0;
            while (await r.ReadAsync(cancellationToken))
            {
                var cost = ReadDecimal(r, 3);
                var accDep = ReadDecimal(r, 4);
                var netBook = Math.Max(0, cost - accDep);
                result.Categories.Add(new FixedAssetCategoryCard
                {
                    MasterId = ReadInt(r, 0),
                    CategoryName = ReadString(r, 1),
                    AssetCount = ReadInt(r, 2),
                    TotalCost = cost,
                    AccumulatedDepreciation = accDep,
                    NetBookValue = netBook,
                    DepreciationRate = ReadDecimal(r, 5),
                    CostPercentage = 0,
                    BadgeColor = BadgeClasses[idx % BadgeClasses.Length]
                });
                idx++;
            }
        }

        var totalCatCost = result.Categories.Sum(x => x.TotalCost);
        foreach (var cat in result.Categories)
        {
            cat.CostPercentage = totalCatCost > 0 ? Math.Round((double)(cat.TotalCost / totalCatCost * 100), 1) : 0;
        }

        // 5. Department Breakdown
        var deptSql = $@"
            SELECT 
                ISNULL(d.DeptID, '') AS DeptId,
                ISNULL(d.Name, 'General / Unassigned') AS DeptName,
                COUNT(f.EntryID) AS AssetCount,
                ISNULL(SUM(CAST(f.AcquisitionCost AS decimal(18,2))), 0) AS TotalCost,
                ISNULL(SUM(CAST(f.AcumulatedDepriciation AS decimal(18,2))), 0) AS AccDep
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.Departments d ON f.DeptID = d.DeptID
            {whereClause}
            GROUP BY d.DeptID, d.Name
            ORDER BY TotalCost DESC;
        ";

        await using (var cmd = new SqlCommand(deptSql, connection))
        {
            foreach (var p in parameters) cmd.Parameters.Add(new SqlParameter(p.ParameterName, p.Value));
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await r.ReadAsync(cancellationToken))
            {
                var cost = ReadDecimal(r, 3);
                var accDep = ReadDecimal(r, 4);
                var netBook = Math.Max(0, cost - accDep);
                result.Departments.Add(new FixedAssetDepartmentCard
                {
                    DeptId = ReadString(r, 0),
                    DeptName = ReadString(r, 1),
                    AssetCount = ReadInt(r, 2),
                    TotalCost = cost,
                    NetBookValue = netBook,
                    CostPercentage = 0
                });
            }
        }

        var totalDeptCost = result.Departments.Sum(x => x.TotalCost);
        foreach (var dept in result.Departments)
        {
            dept.CostPercentage = totalDeptCost > 0 ? Math.Round((double)(dept.TotalCost / totalDeptCost * 100), 1) : 0;
        }

        // 6. Yearly Acquisition Trends
        const string yearTrendSql = @"
            SELECT TOP 8
                CASE 
                    WHEN AcquiredDT IS NULL THEN 'Prior / Unknown'
                    WHEN MONTH(AcquiredDT) >= 7 THEN CAST(YEAR(AcquiredDT) AS varchar) + '-' + CAST(YEAR(AcquiredDT)+1 AS varchar)
                    ELSE CAST(YEAR(AcquiredDT)-1 AS varchar) + '-' + CAST(YEAR(AcquiredDT) AS varchar)
                END AS FiscalYear,
                COUNT(*) AS AssetCount,
                ISNULL(SUM(CAST(AcquisitionCost AS decimal(18,2))), 0) AS AcquisitionCost,
                ISNULL(SUM(CAST(AcumulatedDepriciation AS decimal(18,2))), 0) AS Dep
            FROM dbo.FixedAssets
            WHERE AcquiredDT IS NOT NULL
            GROUP BY 
                CASE 
                    WHEN AcquiredDT IS NULL THEN 'Prior / Unknown'
                    WHEN MONTH(AcquiredDT) >= 7 THEN CAST(YEAR(AcquiredDT) AS varchar) + '-' + CAST(YEAR(AcquiredDT)+1 AS varchar)
                    ELSE CAST(YEAR(AcquiredDT)-1 AS varchar) + '-' + CAST(YEAR(AcquiredDT) AS varchar)
                END
            ORDER BY FiscalYear DESC;
        ";

        await using (var cmd = new SqlCommand(yearTrendSql, connection))
        await using (var r = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await r.ReadAsync(cancellationToken))
            {
                result.YearlyTrends.Add(new FixedAssetYearTrend
                {
                    FiscalYear = ReadString(r, 0),
                    AssetCount = ReadInt(r, 1),
                    AcquisitionCost = ReadDecimal(r, 2),
                    Depreciation = ReadDecimal(r, 3)
                });
            }
        }

        // 7. Status Breakdown
        var totalAssets = Math.Max(1, result.Summary.TotalAssets);
        result.StatusBreakdown = new List<FixedAssetStatusBreakdown>
        {
            new() { StatusName = "In Active Service", Count = result.Summary.ActiveAssets, TotalCost = result.Summary.TotalAcquisitionCost, Percentage = Math.Round((double)result.Summary.ActiveAssets / totalAssets * 100, 1), ColorHex = "#10b981", BadgeClass = "bg-success", Icon = "bi-check-circle" },
            new() { StatusName = "Disposed / Sold", Count = result.Summary.DisposedAssets, TotalCost = 0, Percentage = Math.Round((double)result.Summary.DisposedAssets / totalAssets * 100, 1), ColorHex = "#ef4444", BadgeClass = "bg-danger", Icon = "bi-trash3" },
            new() { StatusName = "Inactive / Idle", Count = result.Summary.InactiveAssets, TotalCost = 0, Percentage = Math.Round((double)result.Summary.InactiveAssets / totalAssets * 100, 1), ColorHex = "#f59e0b", BadgeClass = "bg-warning text-dark", Icon = "bi-pause-circle" }
        };

        // 8. Top Value Assets
        var topValSql = $@"
            SELECT TOP 10
                f.EntryID, ISNULL(f.AssetID,''), ISNULL(f.Description,''),
                ISNULL(m.Description,'Uncategorized'), ISNULL(d.Name,'Unassigned'),
                ISNULL(f.Location,''), f.AcquiredDT,
                ISNULL(CAST(f.AcquisitionCost AS decimal(18,2)),0),
                ISNULL(CAST(f.AcumulatedDepriciation AS decimal(18,2)),0),
                ISNULL(f.Disposed,0), ISNULL(f.InActive,0)
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID = m.EntryID
            LEFT JOIN dbo.Departments d ON f.DeptID = d.DeptID
            {whereClause}
            ORDER BY f.AcquisitionCost DESC;
        ";

        await using (var cmd = new SqlCommand(topValSql, connection))
        {
            foreach (var p in parameters) cmd.Parameters.Add(new SqlParameter(p.ParameterName, p.Value));
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await r.ReadAsync(cancellationToken))
            {
                var cost = ReadDecimal(r, 7);
                var accDep = ReadDecimal(r, 8);
                var disposed = ReadBool(r, 9);
                var inactive = ReadBool(r, 10);

                result.TopValueAssets.Add(new FixedAssetTableRow
                {
                    EntryId = ReadInt(r, 0),
                    AssetId = ReadString(r, 1),
                    Description = ReadString(r, 2),
                    Category = ReadString(r, 3),
                    Department = ReadString(r, 4),
                    Location = ReadString(r, 5),
                    AcquiredDate = ReadDate(r, 6),
                    Cost = cost,
                    AccumulatedDepreciation = accDep,
                    BookValue = Math.Max(0, cost - accDep),
                    Disposed = disposed,
                    InActive = inactive,
                    Status = disposed ? "Disposed" : (inactive ? "Inactive" : "Active")
                });
            }
        }

        // 9. Recent Assets
        var recentSql = $@"
            SELECT TOP 10
                f.EntryID, ISNULL(f.AssetID,''), ISNULL(f.Description,''),
                ISNULL(m.Description,'Uncategorized'), ISNULL(d.Name,'Unassigned'),
                ISNULL(f.Location,''), f.AcquiredDT,
                ISNULL(CAST(f.AcquisitionCost AS decimal(18,2)),0),
                ISNULL(CAST(f.AcumulatedDepriciation AS decimal(18,2)),0),
                ISNULL(f.Disposed,0), ISNULL(f.InActive,0)
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID = m.EntryID
            LEFT JOIN dbo.Departments d ON f.DeptID = d.DeptID
            {whereClause}
            ORDER BY f.EntryID DESC;
        ";

        await using (var cmd = new SqlCommand(recentSql, connection))
        {
            foreach (var p in parameters) cmd.Parameters.Add(new SqlParameter(p.ParameterName, p.Value));
            await using var r = await cmd.ExecuteReaderAsync(cancellationToken);
            while (await r.ReadAsync(cancellationToken))
            {
                var cost = ReadDecimal(r, 7);
                var accDep = ReadDecimal(r, 8);
                var disposed = ReadBool(r, 9);
                var inactive = ReadBool(r, 10);

                result.RecentAssets.Add(new FixedAssetTableRow
                {
                    EntryId = ReadInt(r, 0),
                    AssetId = ReadString(r, 1),
                    Description = ReadString(r, 2),
                    Category = ReadString(r, 3),
                    Department = ReadString(r, 4),
                    Location = ReadString(r, 5),
                    AcquiredDate = ReadDate(r, 6),
                    Cost = cost,
                    AccumulatedDepreciation = accDep,
                    BookValue = Math.Max(0, cost - accDep),
                    Disposed = disposed,
                    InActive = inactive,
                    Status = disposed ? "Disposed" : (inactive ? "Inactive" : "Active")
                });
            }
        }

        // 10. Recent Disposals
        const string dispSql = @"
            SELECT TOP 10
                f.EntryID, ISNULL(f.AssetID,''), ISNULL(f.Description,''),
                f.DTDisposal,
                CASE ISNULL(f.Disposed_Type, 0) WHEN 1 THEN 'Sale' WHEN 2 THEN 'Scrap / Write-off' WHEN 3 THEN 'Transfer' ELSE 'Disposed' END,
                ISNULL(CAST(f.AcquisitionCost AS decimal(18,2)), 0),
                ISNULL(CAST(f.SalePrice AS decimal(18,2)), 0),
                ISNULL(CAST(ISNULL(f.SalePrice,0) - ISNULL(f.WDV,0) AS decimal(18,2)), 0),
                ISNULL(dis.Buyer, '')
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Disposals dis ON f.AssetID = dis.AssetID
            WHERE f.Disposed = 1
            ORDER BY f.DTDisposal DESC, f.EntryID DESC;
        ";

        await using (var cmd = new SqlCommand(dispSql, connection))
        await using (var r = await cmd.ExecuteReaderAsync(cancellationToken))
        {
            while (await r.ReadAsync(cancellationToken))
            {
                result.RecentDisposals.Add(new FixedAssetDisposalRow
                {
                    EntryId = ReadInt(r, 0),
                    AssetId = ReadString(r, 1),
                    Description = ReadString(r, 2),
                    DisposalDate = ReadDate(r, 3),
                    DisposalType = ReadString(r, 4),
                    Cost = ReadDecimal(r, 5),
                    SalePrice = ReadDecimal(r, 6),
                    GainLoss = ReadDecimal(r, 7),
                    Buyer = ReadString(r, 8)
                });
            }
        }

        return result;
    }

    private static async Task<List<FixedAssetLookupItem>> LoadLookupAsync(SqlConnection connection, string sql, CancellationToken cancellationToken)
    {
        var list = new List<FixedAssetLookupItem>();
        await using var cmd = new SqlCommand(sql, connection);
        await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new FixedAssetLookupItem
            {
                Id = reader.GetValue(0)?.ToString() ?? "",
                Name = ReadString(reader, 1),
                PercentRate = ReadDecimal(reader, 2)
            });
        }
        return list;
    }

    private static string ReadString(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal).Trim();
    private static int ReadInt(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0 : Convert.ToInt32(reader.GetValue(ordinal));
    private static decimal ReadDecimal(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0m : Convert.ToDecimal(reader.GetValue(ordinal));
    private static decimal ReadDecimal(object? value) => value == null || value == DBNull.Value ? 0m : Convert.ToDecimal(value);
    private static DateTime? ReadDate(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToDateTime(reader.GetValue(ordinal));
    private static bool ReadBool(SqlDataReader reader, int ordinal) => !reader.IsDBNull(ordinal) && Convert.ToBoolean(reader.GetValue(ordinal));
}
