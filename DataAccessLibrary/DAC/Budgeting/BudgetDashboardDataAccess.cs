using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Budgeting;
using DataAccessLibrary.Models.ViewModels.Budgeting;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Budgeting
{
    public class BudgetDashboardDataAccess : IBudgetDashboardDataAccess
    {
        private readonly string _connectionString;
        private readonly IBudgetDataAccess _budgetDataAccess;

        private static readonly string[] Palette = new[]
        {
            "#10b981", "#3b82f6", "#f59e0b", "#8b5cf6", "#ec4899",
            "#06b6d4", "#14b8a6", "#f97316", "#6366f1", "#84cc16",
            "#e11d48", "#0ea5e9", "#a855f7", "#d97706", "#64748b"
        };

        private static readonly (int index, string name, string colName, int calMonth, int offsetYear)[] Months = new[]
        {
            (1, "July", "Month_July", 7, 0),
            (2, "August", "Month_August", 8, 0),
            (3, "September", "Month_September", 9, 0),
            (4, "October", "Month_October", 10, 0),
            (5, "November", "Month_November", 11, 0),
            (6, "December", "Month_December", 12, 0),
            (7, "January", "Month_January", 1, 1),
            (8, "February", "Month_February", 2, 1),
            (9, "March", "Month_March", 3, 1),
            (10, "April", "Month_April", 4, 1),
            (11, "May", "Month_May", 5, 1),
            (12, "June", "Month_June", 6, 1)
        };

        public BudgetDashboardDataAccess(IConfiguration configuration, IBudgetDataAccess budgetDataAccess)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("DefaultConnection string not found.");
            _budgetDataAccess = budgetDataAccess;
        }

        private SqlConnection CreateConnection() => new(_connectionString);

        public async Task<BudgetDashboardResult> GetDashboardDataAsync(BudgetDashboardFilter filter, CancellationToken cancellationToken = default)
        {
            var result = new BudgetDashboardResult();
            var allYears = _budgetDataAccess.GetFinancialYears();
            var curFy = BudgetDashboardFilter.GetCurrentFinancialYear();
            if (!allYears.Any(x => string.Equals(x.Code, curFy, StringComparison.OrdinalIgnoreCase)))
            {
                var parts = curFy.Split('-');
                if (parts.Length > 0 && int.TryParse(parts[0].Trim(), out int sy))
                {
                    allYears.Add(new FinancialYearOption
                    {
                        Code = curFy,
                        DisplayText = $"{curFy} [ {new DateTime(sy, 7, 1):dd-MMM-yyyy} to {new DateTime(sy + 1, 6, 30):dd-MMM-yyyy} ]",
                        StartYear = sy,
                        StartDate = new DateTime(sy, 7, 1),
                        EndDate = new DateTime(sy + 1, 6, 30)
                    });
                }
            }
            result.FinancialYears = allYears.OrderByDescending(x => x.StartYear).ToList();

            if (string.IsNullOrWhiteSpace(filter.FinancialYear) && result.FinancialYears.Count > 0)
            {
                filter.FinancialYear = curFy;
            }

            bool isExpenses = !string.Equals(filter.Mode, "Income", StringComparison.OrdinalIgnoreCase);
            result.MasterHeads = await _budgetDataAccess.GetMasterHeadsAsync(isExpenses);

            var (startYear, fyStartDate, fyEndDate) = ParseFinancialYear(filter.FinancialYear);

            DateTime filterStartDate;
            DateTime filterEndDate;

            if (filter.MonthIndex >= 1 && filter.MonthIndex <= 12)
            {
                var mInfo = Months[filter.MonthIndex - 1];
                int calYr = startYear + mInfo.offsetYear;
                filterStartDate = new DateTime(calYr, mInfo.calMonth, 1);
                filterEndDate = new DateTime(calYr, mInfo.calMonth, DateTime.DaysInMonth(calYr, mInfo.calMonth), 23, 59, 59);
            }
            else if (filter.QuarterIndex >= 1 && filter.QuarterIndex <= 4)
            {
                var (quarterStart, quarterEnd) = GetQuarterDates(startYear, filter.QuarterIndex);
                filterStartDate = quarterStart;
                filterEndDate = quarterEnd;
            }
            else
            {
                filterStartDate = fyStartDate;
                filterEndDate = fyEndDate;
            }

            await using var connection = CreateConnection();
            await connection.OpenAsync(cancellationToken);

            if (isExpenses)
            {
                await LoadExpensesDashboardAsync(connection, filter, result, startYear, filterStartDate, filterEndDate, fyStartDate, fyEndDate, cancellationToken);
            }
            else
            {
                await LoadIncomeDashboardAsync(connection, filter, result, startYear, filterStartDate, filterEndDate, fyStartDate, fyEndDate, cancellationToken);
            }

            return result;
        }

        private async Task LoadExpensesDashboardAsync(
            SqlConnection connection,
            BudgetDashboardFilter filter,
            BudgetDashboardResult result,
            int startYear,
            DateTime filterStartDate,
            DateTime filterEndDate,
            DateTime fyStartDate,
            DateTime fyEndDate,
            CancellationToken cancellationToken)
        {
            string headFilter = (filter.HeadCode != "0" && !string.IsNullOrWhiteSpace(filter.HeadCode)) ? " AND a.Type = @HeadCode" : "";
            string searchFilter = !string.IsNullOrWhiteSpace(filter.SearchText) ? " AND (a.AccNo LIKE @Like OR a.AccTitle LIKE @Like)" : "";

            double monthFraction = GetBudgetPeriodFraction(filter);

            string sql = $@"
                SELECT 
                    a.AccNo, 
                    a.AccTitle, 
                    ISNULL(a.Type, '') AS HeadCode,
                    ISNULL(h.Head, 'Unassigned') AS HeadName,
                    ISNULL(e.Proposed, 0) AS FullProposed,
                    ISNULL(e.Revised, 0) AS FullRevised,
                    ISNULL((
                        SELECT SUM(v.Debit - v.Credit)
                        FROM Vouchers v
                        WHERE v.Accno = a.AccNo
                          AND v.VDate >= @FilterStart AND v.VDate <= @FilterEnd
                    ), 0) AS ActualSpent
                FROM Accounts a
                LEFT JOIN Heads h ON a.Type = h.Code
                LEFT JOIN FinancialBudgeting_Expenses e ON a.AccNo = e.AccNo AND e.FinancialYear = @FinancialYear
                WHERE a.BudgetAccount = 1 AND a.Active = 1
                  AND (TRY_CAST(a.Type AS INT) >= 41 OR LEFT(a.AccNo, 6) = '12-001')
                  {headFilter}
                  {searchFilter}
                ORDER BY FullProposed DESC, a.AccNo;
            ";

            var accountsList = new List<BudgetAccountRow>();
            decimal totalProposed = 0;
            decimal totalRevised = 0;
            decimal totalBudget = 0;
            decimal totalActual = 0;

            await using (var cmd = new SqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@FinancialYear", filter.FinancialYear);
                cmd.Parameters.AddWithValue("@FilterStart", filterStartDate);
                cmd.Parameters.AddWithValue("@FilterEnd", filterEndDate);
                if (!string.IsNullOrWhiteSpace(filter.HeadCode) && filter.HeadCode != "0")
                    cmd.Parameters.AddWithValue("@HeadCode", filter.HeadCode);
                if (!string.IsNullOrWhiteSpace(filter.SearchText))
                    cmd.Parameters.AddWithValue("@Like", $"%{filter.SearchText.Trim()}%");

                await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
                while (await reader.ReadAsync(cancellationToken))
                {
                    var fullProp = Convert.ToDecimal(reader["FullProposed"]);
                    var fullRev = Convert.ToDecimal(reader["FullRevised"]);
                    var actual = Convert.ToDecimal(reader["ActualSpent"]);

                    var budgetThisPeriod = (fullRev > 0 ? fullRev : fullProp) * (decimal)monthFraction;
                    var propThisPeriod = fullProp * (decimal)monthFraction;
                    var revThisPeriod = fullRev * (decimal)monthFraction;

                    totalProposed += propThisPeriod;
                    totalRevised += revThisPeriod;
                    totalBudget += budgetThisPeriod;
                    totalActual += actual;

                    accountsList.Add(new BudgetAccountRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        CategoryName = reader["HeadName"]?.ToString() ?? "",
                        BudgetAmount = Math.Round(budgetThisPeriod, 0),
                        ActualAmount = Math.Round(actual, 0)
                    });
                }
            }

            result.Summary.TotalBudget = Math.Round(totalBudget, 0);
            result.Summary.TotalActual = Math.Round(totalActual, 0);
            result.Summary.BudgetProposed = Math.Round(totalProposed, 0);
            result.Summary.BudgetRevised = Math.Round(totalRevised, 0);
            result.Summary.AccountCount = accountsList.Count;

            var catMap = accountsList
                .GroupBy(x => x.CategoryName)
                .OrderByDescending(g => g.Sum(x => x.BudgetAmount))
                .ToList();

            result.Summary.HeadCount = catMap.Count;
            int colorIdx = 0;

            foreach (var g in catMap)
            {
                var bgt = g.Sum(x => x.BudgetAmount);
                var act = g.Sum(x => x.ActualAmount);
                var pctTotal = totalBudget > 0 ? (double)(bgt / totalBudget * 100) : 0;

                result.Categories.Add(new BudgetCategoryStat
                {
                    CategoryName = g.Key,
                    BudgetAmount = bgt,
                    ActualAmount = act,
                    PercentageOfTotalBudget = Math.Round(pctTotal, 1),
                    ColorHex = Palette[colorIdx % Palette.Length],
                    BadgeClass = $"cat-badge-{colorIdx % Palette.Length}"
                });
                colorIdx++;
            }

            result.MonthlyTrends = await LoadMonthlyTrendsExpensesAsync(connection, filter.FinancialYear, startYear, cancellationToken);
            result.TopAccounts = accountsList.Take(15).ToList();
        }

        private async Task<List<BudgetMonthlyTrend>> LoadMonthlyTrendsExpensesAsync(
            SqlConnection connection,
            string financialYear,
            int startYear,
            CancellationToken cancellationToken)
        {
            var trends = new List<BudgetMonthlyTrend>();

            decimal yearBudget = 0;
            await using (var cmd = new SqlCommand("SELECT ISNULL(SUM(CASE WHEN ISNULL(Revised,0)>0 THEN Revised ELSE ISNULL(Proposed,0) END),0) FROM dbo.FinancialBudgeting_Expenses WHERE FinancialYear = @FinancialYear", connection))
            {
                cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                yearBudget = Convert.ToDecimal(await cmd.ExecuteScalarAsync(cancellationToken));
            }

            decimal monthlyBudget = yearBudget / 12.0m;

            for (int i = 0; i < Months.Length; i++)
            {
                var mInfo = Months[i];
                int calYr = startYear + mInfo.offsetYear;
                var mStart = new DateTime(calYr, mInfo.calMonth, 1);
                var mEnd = new DateTime(calYr, mInfo.calMonth, DateTime.DaysInMonth(calYr, mInfo.calMonth), 23, 59, 59);

                decimal actualSpent = 0;
                const string actSql = @"
                    SELECT ISNULL(SUM(v.Debit - v.Credit), 0)
                    FROM Vouchers v
                    INNER JOIN Accounts a ON v.Accno = a.AccNo
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND (TRY_CAST(a.Type AS INT) >= 41 OR LEFT(a.AccNo, 6) = '12-001')
                      AND v.VDate >= @MStart AND v.VDate <= @MEnd;
                ";

                await using (var cmd = new SqlCommand(actSql, connection))
                {
                    cmd.Parameters.AddWithValue("@MStart", mStart);
                    cmd.Parameters.AddWithValue("@MEnd", mEnd);
                    actualSpent = Convert.ToDecimal(await cmd.ExecuteScalarAsync(cancellationToken));
                }

                trends.Add(new BudgetMonthlyTrend
                {
                    MonthIndex = mInfo.index,
                    MonthName = mInfo.name,
                    MonthCode = mInfo.name.Substring(0, 3),
                    BudgetAmount = Math.Round(monthlyBudget, 0),
                    ActualAmount = Math.Round(actualSpent, 0)
                });
            }

            return trends;
        }

        private async Task LoadIncomeDashboardAsync(
            SqlConnection connection,
            BudgetDashboardFilter filter,
            BudgetDashboardResult result,
            int startYear,
            DateTime filterStartDate,
            DateTime filterEndDate,
            DateTime fyStartDate,
            DateTime fyEndDate,
            CancellationToken cancellationToken)
        {
            string headFilter = (filter.HeadCode != "0" && !string.IsNullOrWhiteSpace(filter.HeadCode)) ? " AND a.Type = @HeadCode" : "";
            string searchFilter = !string.IsNullOrWhiteSpace(filter.SearchText) ? " AND (a.AccNo LIKE @Like OR a.AccTitle LIKE @Like)" : "";

            string activeBudgetCols = BuildPeriodColumnsIncome(filter);

            string sql = $@"
                SELECT 
                    a.AccNo, 
                    a.AccTitle, 
                    ISNULL(a.Type, '') AS HeadCode,
                    ISNULL(h.Head, 'Unassigned') AS HeadName,
                    ({activeBudgetCols}) AS PeriodBudget,
                    ISNULL((
                        SELECT SUM(v.Credit - v.Debit)
                        FROM Vouchers v
                        WHERE v.Accno = a.AccNo
                          AND v.VDate >= @FilterStart AND v.VDate <= @FilterEnd
                    ), 0) AS ActualIncome
                FROM Accounts a
                LEFT JOIN Heads h ON a.Type = h.Code
                LEFT JOIN FinancialBudgeting b ON a.AccNo = b.AccNo AND b.FinancialYear = @FinancialYear
                WHERE a.BudgetAccount = 1 AND a.Active = 1
                  AND a.Type IN ('31', '32', '33')
                  {headFilter}
                  {searchFilter}
                ORDER BY PeriodBudget DESC, a.AccNo;
            ";

            var accountsList = new List<BudgetAccountRow>();
            decimal totalBudget = 0;
            decimal totalActual = 0;

            await using (var cmd = new SqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@FinancialYear", filter.FinancialYear);
                cmd.Parameters.AddWithValue("@FilterStart", filterStartDate);
                cmd.Parameters.AddWithValue("@FilterEnd", filterEndDate);
                if (!string.IsNullOrWhiteSpace(filter.HeadCode) && filter.HeadCode != "0")
                    cmd.Parameters.AddWithValue("@HeadCode", filter.HeadCode);
                if (!string.IsNullOrWhiteSpace(filter.SearchText))
                    cmd.Parameters.AddWithValue("@Like", $"%{filter.SearchText.Trim()}%");

                await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
                while (await reader.ReadAsync(cancellationToken))
                {
                    var bgt = Convert.ToDecimal(reader["PeriodBudget"]);
                    var act = Convert.ToDecimal(reader["ActualIncome"]);

                    totalBudget += bgt;
                    totalActual += act;

                    accountsList.Add(new BudgetAccountRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        CategoryName = reader["HeadName"]?.ToString() ?? "",
                        BudgetAmount = Math.Round(bgt, 0),
                        ActualAmount = Math.Round(act, 0)
                    });
                }
            }

            result.Summary.TotalBudget = Math.Round(totalBudget, 0);
            result.Summary.TotalActual = Math.Round(totalActual, 0);
            result.Summary.BudgetProposed = Math.Round(totalBudget, 0);
            result.Summary.BudgetRevised = 0;
            result.Summary.AccountCount = accountsList.Count;

            var catMap = accountsList
                .GroupBy(x => x.CategoryName)
                .OrderByDescending(g => g.Sum(x => x.BudgetAmount))
                .ToList();

            result.Summary.HeadCount = catMap.Count;
            int colorIdx = 0;

            foreach (var g in catMap)
            {
                var bgt = g.Sum(x => x.BudgetAmount);
                var act = g.Sum(x => x.ActualAmount);
                var pctTotal = totalBudget > 0 ? (double)(bgt / totalBudget * 100) : 0;

                result.Categories.Add(new BudgetCategoryStat
                {
                    CategoryName = g.Key,
                    BudgetAmount = bgt,
                    ActualAmount = act,
                    PercentageOfTotalBudget = Math.Round(pctTotal, 1),
                    ColorHex = Palette[colorIdx % Palette.Length],
                    BadgeClass = $"cat-badge-{colorIdx % Palette.Length}"
                });
                colorIdx++;
            }

            result.MonthlyTrends = await LoadMonthlyTrendsIncomeAsync(connection, filter.FinancialYear, startYear, cancellationToken);
            result.TopAccounts = accountsList.Take(15).ToList();
        }

        private async Task<List<BudgetMonthlyTrend>> LoadMonthlyTrendsIncomeAsync(
            SqlConnection connection,
            string financialYear,
            int startYear,
            CancellationToken cancellationToken)
        {
            var trends = new List<BudgetMonthlyTrend>();

            for (int i = 0; i < Months.Length; i++)
            {
                var mInfo = Months[i];
                int calYr = startYear + mInfo.offsetYear;
                var mStart = new DateTime(calYr, mInfo.calMonth, 1);
                var mEnd = new DateTime(calYr, mInfo.calMonth, DateTime.DaysInMonth(calYr, mInfo.calMonth), 23, 59, 59);

                string colName = mInfo.colName;
                string bgtSql = $@"
                    SELECT ISNULL(SUM(CASE WHEN ISNULL(b.{colName}_R,0)>0 THEN b.{colName}_R ELSE ISNULL(b.{colName},0) END), 0)
                    FROM dbo.FinancialBudgeting b
                    WHERE b.FinancialYear = @FinancialYear;
                ";

                decimal bgt = 0;
                await using (var cmd = new SqlCommand(bgtSql, connection))
                {
                    cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                    bgt = Convert.ToDecimal(await cmd.ExecuteScalarAsync(cancellationToken));
                }

                decimal actualIncome = 0;
                const string actSql = @"
                    SELECT ISNULL(SUM(v.Credit - v.Debit), 0)
                    FROM Vouchers v
                    INNER JOIN Accounts a ON v.Accno = a.AccNo
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND a.Type IN ('31', '32', '33')
                      AND v.VDate >= @MStart AND v.VDate <= @MEnd;
                ";

                await using (var cmd = new SqlCommand(actSql, connection))
                {
                    cmd.Parameters.AddWithValue("@MStart", mStart);
                    cmd.Parameters.AddWithValue("@MEnd", mEnd);
                    actualIncome = Convert.ToDecimal(await cmd.ExecuteScalarAsync(cancellationToken));
                }

                trends.Add(new BudgetMonthlyTrend
                {
                    MonthIndex = mInfo.index,
                    MonthName = mInfo.name,
                    MonthCode = mInfo.name.Substring(0, 3),
                    BudgetAmount = Math.Round(bgt, 0),
                    ActualAmount = Math.Round(actualIncome, 0)
                });
            }

            return trends;
        }

        private static string BuildPeriodColumnsIncome(BudgetDashboardFilter filter)
        {
            if (filter.MonthIndex >= 1 && filter.MonthIndex <= 12)
            {
                var m = Months[filter.MonthIndex - 1].colName;
                return $"CASE WHEN ISNULL(b.{m}_R, 0) > 0 THEN b.{m}_R ELSE ISNULL(b.{m}, 0) END";
            }

            if (filter.QuarterIndex >= 1 && filter.QuarterIndex <= 4)
            {
                int startM = (filter.QuarterIndex - 1) * 3;
                var qMonths = Months.Skip(startM).Take(3).Select(x => x.colName).ToList();
                var parts = qMonths.Select(m => $"(CASE WHEN ISNULL(b.{m}_R, 0) > 0 THEN b.{m}_R ELSE ISNULL(b.{m}, 0) END)");
                return string.Join(" + ", parts);
            }

            var allMonths = Months.Select(x => $"(CASE WHEN ISNULL(b.{x.colName}_R, 0) > 0 THEN b.{x.colName}_R ELSE ISNULL(b.{x.colName}, 0) END)");
            return string.Join(" + ", allMonths);
        }

        private static double GetBudgetPeriodFraction(BudgetDashboardFilter filter)
        {
            if (filter.MonthIndex >= 1 && filter.MonthIndex <= 12) return 1.0 / 12.0;
            if (filter.QuarterIndex >= 1 && filter.QuarterIndex <= 4) return 3.0 / 12.0;
            return 1.0;
        }

        private static (int startYear, DateTime start, DateTime end) ParseFinancialYear(string? fy)
        {
            int sy = DateTime.Today.Year;
            if (!string.IsNullOrWhiteSpace(fy))
            {
                var parts = fy.Split('-');
                if (parts.Length > 0 && int.TryParse(parts[0].Trim(), out int parsed))
                {
                    sy = parsed;
                }
            }
            return (sy, new DateTime(sy, 7, 1), new DateTime(sy + 1, 6, 30, 23, 59, 59));
        }

        private static (DateTime start, DateTime end) GetQuarterDates(int startYear, int quarter)
        {
            return quarter switch
            {
                1 => (new DateTime(startYear, 7, 1), new DateTime(startYear, 9, 30, 23, 59, 59)),
                2 => (new DateTime(startYear, 10, 1), new DateTime(startYear, 12, 31, 23, 59, 59)),
                3 => (new DateTime(startYear + 1, 1, 1), new DateTime(startYear + 1, 3, 31, 23, 59, 59)),
                4 => (new DateTime(startYear + 1, 4, 1), new DateTime(startYear + 1, 6, 30, 23, 59, 59)),
                _ => (new DateTime(startYear, 7, 1), new DateTime(startYear + 1, 6, 30, 23, 59, 59))
            };
        }
    }
}

