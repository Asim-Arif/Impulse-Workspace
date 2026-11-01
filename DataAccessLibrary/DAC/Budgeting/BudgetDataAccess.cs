using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Budgeting;
using DataAccessLibrary.Models.ViewModels.Budgeting;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Budgeting
{
    public class BudgetDataAccess : IBudgetDataAccess
    {
        private readonly string _connectionString;

        public BudgetDataAccess(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? throw new InvalidOperationException("DefaultConnection string not found.");
        }

        public List<FinancialYearOption> GetFinancialYears(int startYear = 2010)
        {
            var list = new List<FinancialYearOption>();
            int currentYear = DateTime.Today.Year;
            for (int i = startYear; i <= currentYear + 1; i++)
            {
                var startDate = new DateTime(i, 7, 1);
                var endDate = new DateTime(i + 1, 6, 30);
                var code = $"{i} - {i + 1}";
                var label = $"{code} [ {startDate:dd-MMM-yyyy} to {endDate:dd-MMM-yyyy} ]";

                list.Add(new FinancialYearOption
                {
                    Code = code,
                    DisplayText = label,
                    StartYear = i,
                    StartDate = startDate,
                    EndDate = endDate
                });
            }
            list.Reverse();
            return list;
        }

        public async Task<List<MasterHeadOption>> GetMasterHeadsAsync(bool isExpenses)
        {
            var list = new List<MasterHeadOption>
            {
                new MasterHeadOption { Code = "0", Head = "<All Heads>", Type = "" }
            };

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string query;
            if (isExpenses)
            {
                query = @"
                    SELECT Code, Head, type, ISNULL(BudgetAccount_Glance, 0) AS Glance
                    FROM Heads 
                    WHERE Code IN (
                        SELECT DISTINCT Type FROM Accounts 
                        WHERE BudgetAccount = 1 AND (TRY_CAST(Type AS INT) >= 41 OR LEFT(AccNo, 6) = '12-001')
                    )
                    OR TRY_CAST(Code AS INT) >= 41
                    ORDER BY Code";
            }
            else
            {
                query = @"
                    SELECT Code, Head, type, ISNULL(BudgetAccount_Glance, 0) AS Glance
                    FROM Heads 
                    WHERE Code IN (
                        SELECT DISTINCT Type FROM Accounts 
                        WHERE BudgetAccount = 1 AND Type IN ('31', '32', '33')
                    )
                    OR Code IN ('31', '32', '33')
                    ORDER BY Code";
            }

            using var cmd = new SqlCommand(query, conn);
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new MasterHeadOption
                {
                    Code = reader["Code"]?.ToString() ?? "",
                    Head = reader["Head"]?.ToString() ?? "",
                    Type = reader["type"]?.ToString() ?? "",
                    BudgetAccountGlance = reader["Glance"] != DBNull.Value && Convert.ToBoolean(reader["Glance"])
                });
            }

            return list;
        }

        public async Task<List<MonthwiseBudgetRow>> GetMonthwiseBudgetAsync(string financialYear, string headCode = "0", bool isRevision = false)
        {
            var list = new List<MonthwiseBudgetRow>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = @"
                SELECT 
                    a.AccNo, 
                    a.AccTitle, 
                    a.Type,
                    ISNULL(h.Head, '') AS HeadName,
                    " + (isRevision ? @"
                    ISNULL(b.Month_July_R, 0) AS M_Jul,
                    ISNULL(b.Month_August_R, 0) AS M_Aug,
                    ISNULL(b.Month_September_R, 0) AS M_Sep,
                    ISNULL(b.Month_October_R, 0) AS M_Oct,
                    ISNULL(b.Month_November_R, 0) AS M_Nov,
                    ISNULL(b.Month_December_R, 0) AS M_Dec,
                    ISNULL(b.Month_January_R, 0) AS M_Jan,
                    ISNULL(b.Month_February_R, 0) AS M_Feb,
                    ISNULL(b.Month_March_R, 0) AS M_Mar,
                    ISNULL(b.Month_April_R, 0) AS M_Apr,
                    ISNULL(b.Month_May_R, 0) AS M_May,
                    ISNULL(b.Month_June_R, 0) AS M_Jun
                    " : @"
                    ISNULL(b.Month_July, 0) AS M_Jul,
                    ISNULL(b.Month_August, 0) AS M_Aug,
                    ISNULL(b.Month_September, 0) AS M_Sep,
                    ISNULL(b.Month_October, 0) AS M_Oct,
                    ISNULL(b.Month_November, 0) AS M_Nov,
                    ISNULL(b.Month_December, 0) AS M_Dec,
                    ISNULL(b.Month_January, 0) AS M_Jan,
                    ISNULL(b.Month_February, 0) AS M_Feb,
                    ISNULL(b.Month_March, 0) AS M_Mar,
                    ISNULL(b.Month_April, 0) AS M_Apr,
                    ISNULL(b.Month_May, 0) AS M_May,
                    ISNULL(b.Month_June, 0) AS M_Jun
                    ") + @"
                FROM Accounts a
                LEFT JOIN Heads h ON a.Type = h.Code
                LEFT JOIN FinancialBudgeting b ON a.AccNo = b.AccNo AND b.FinancialYear = @FinancialYear
                WHERE a.BudgetAccount = 1 
                  AND a.Active = 1 
                  AND a.Type IN ('31', '32', '33')
                  " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                ORDER BY a.AccNo";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
            if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
            {
                cmd.Parameters.AddWithValue("@HeadCode", headCode);
            }

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new MonthwiseBudgetRow
                {
                    AccNo = reader["AccNo"]?.ToString() ?? "",
                    AccTitle = reader["AccTitle"]?.ToString() ?? "",
                    Type = reader["Type"]?.ToString() ?? "",
                    HeadName = reader["HeadName"]?.ToString() ?? "",
                    MonthJuly = reader["M_Jul"] != DBNull.Value ? Convert.ToDecimal(reader["M_Jul"]) : 0,
                    MonthAugust = reader["M_Aug"] != DBNull.Value ? Convert.ToDecimal(reader["M_Aug"]) : 0,
                    MonthSeptember = reader["M_Sep"] != DBNull.Value ? Convert.ToDecimal(reader["M_Sep"]) : 0,
                    MonthOctober = reader["M_Oct"] != DBNull.Value ? Convert.ToDecimal(reader["M_Oct"]) : 0,
                    MonthNovember = reader["M_Nov"] != DBNull.Value ? Convert.ToDecimal(reader["M_Nov"]) : 0,
                    MonthDecember = reader["M_Dec"] != DBNull.Value ? Convert.ToDecimal(reader["M_Dec"]) : 0,
                    MonthJanuary = reader["M_Jan"] != DBNull.Value ? Convert.ToDecimal(reader["M_Jan"]) : 0,
                    MonthFebruary = reader["M_Feb"] != DBNull.Value ? Convert.ToDecimal(reader["M_Feb"]) : 0,
                    MonthMarch = reader["M_Mar"] != DBNull.Value ? Convert.ToDecimal(reader["M_Mar"]) : 0,
                    MonthApril = reader["M_Apr"] != DBNull.Value ? Convert.ToDecimal(reader["M_Apr"]) : 0,
                    MonthMay = reader["M_May"] != DBNull.Value ? Convert.ToDecimal(reader["M_May"]) : 0,
                    MonthJune = reader["M_Jun"] != DBNull.Value ? Convert.ToDecimal(reader["M_Jun"]) : 0
                });
            }

            return list;
        }

        public async Task<bool> SaveMonthwiseCellAsync(string financialYear, string accNo, string monthColumn, decimal amount, bool isRevision = false)
        {
            var validMonths = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Month_July", "Month_August", "Month_September", "Month_October", "Month_November", "Month_December",
                "Month_January", "Month_February", "Month_March", "Month_April", "Month_May", "Month_June"
            };

            string targetCol = monthColumn;
            if (!validMonths.Contains(targetCol))
            {
                targetCol = "Month_" + monthColumn.Trim();
                if (!validMonths.Contains(targetCol)) return false;
            }

            if (isRevision && !targetCol.EndsWith("_R", StringComparison.OrdinalIgnoreCase))
            {
                targetCol += "_R";
            }

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = $@"
                IF EXISTS (SELECT 1 FROM FinancialBudgeting WHERE FinancialYear = @FinancialYear AND AccNo = @AccNo)
                BEGIN
                    UPDATE FinancialBudgeting 
                    SET {targetCol} = @Amount
                    WHERE FinancialYear = @FinancialYear AND AccNo = @AccNo;
                END
                ELSE
                BEGIN
                    INSERT INTO FinancialBudgeting (FinancialYear, AccNo, {targetCol})
                    VALUES (@FinancialYear, @AccNo, @Amount);
                END";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
            cmd.Parameters.AddWithValue("@AccNo", accNo);
            cmd.Parameters.AddWithValue("@Amount", (long)amount);

            var rows = await cmd.ExecuteNonQueryAsync();
            return rows > 0;
        }

        public async Task<bool> UpdateMonthwiseRevisionsAsync(string financialYear)
        {
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = @"
                UPDATE FinancialBudgeting 
                SET 
                    Month_July_R = Month_July,
                    Month_August_R = Month_August,
                    Month_September_R = Month_September,
                    Month_October_R = Month_October,
                    Month_November_R = Month_November,
                    Month_December_R = Month_December,
                    Month_January_R = Month_January,
                    Month_February_R = Month_February,
                    Month_March_R = Month_March,
                    Month_April_R = Month_April,
                    Month_May_R = Month_May,
                    Month_June_R = Month_June
                WHERE FinancialYear = @FinancialYear";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);

            await cmd.ExecuteNonQueryAsync();
            return true;
        }

        public async Task<List<ExpenseBudgetRow>> GetExpenseBudgetAsync(string financialYear, string headCode = "0")
        {
            var list = new List<ExpenseBudgetRow>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = @"
                SELECT 
                    a.AccNo, 
                    a.AccTitle, 
                    a.Type,
                    ISNULL(h.Head, '') AS HeadName,
                    ISNULL(e.Proposed, 0) AS Proposed,
                    ISNULL(e.Revised, 0) AS Revised
                FROM Accounts a
                LEFT JOIN Heads h ON a.Type = h.Code
                LEFT JOIN FinancialBudgeting_Expenses e ON a.AccNo = e.AccNo AND e.FinancialYear = @FinancialYear
                WHERE a.BudgetAccount = 1 
                  AND a.Active = 1 
                  AND (TRY_CAST(a.Type AS INT) >= 41 OR LEFT(a.AccNo, 6) = '12-001')
                  " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                ORDER BY a.AccNo";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
            if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
            {
                cmd.Parameters.AddWithValue("@HeadCode", headCode);
            }

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new ExpenseBudgetRow
                {
                    AccNo = reader["AccNo"]?.ToString() ?? "",
                    AccTitle = reader["AccTitle"]?.ToString() ?? "",
                    Type = reader["Type"]?.ToString() ?? "",
                    HeadName = reader["HeadName"]?.ToString() ?? "",
                    Proposed = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                    Revised = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0
                });
            }

            return list;
        }

        public async Task<bool> SaveExpenseCellAsync(string financialYear, string accNo, string column, decimal amount)
        {
            string targetCol = column.Equals("Revised", StringComparison.OrdinalIgnoreCase) ? "Revised" : "Proposed";

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = $@"
                IF EXISTS (SELECT 1 FROM FinancialBudgeting_Expenses WHERE FinancialYear = @FinancialYear AND AccNo = @AccNo)
                BEGIN
                    UPDATE FinancialBudgeting_Expenses 
                    SET {targetCol} = @Amount
                    WHERE FinancialYear = @FinancialYear AND AccNo = @AccNo;
                END
                ELSE
                BEGIN
                    INSERT INTO FinancialBudgeting_Expenses (FinancialYear, AccNo, {targetCol})
                    VALUES (@FinancialYear, @AccNo, @Amount);
                END";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
            cmd.Parameters.AddWithValue("@AccNo", accNo);
            cmd.Parameters.AddWithValue("@Amount", (long)amount);

            var rows = await cmd.ExecuteNonQueryAsync();
            return rows > 0;
        }

        public async Task<bool> UpdateExpenseRevisionsAsync(string financialYear)
        {
            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = @"
                UPDATE FinancialBudgeting_Expenses 
                SET Revised = Proposed
                WHERE FinancialYear = @FinancialYear";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);

            await cmd.ExecuteNonQueryAsync();
            return true;
        }

        public async Task<List<BudgetReportRow>> GetQuarterlyReportDataAsync(string financialYear, int quarter, string headCode, bool isExpenses)
        {
            var list = new List<BudgetReportRow>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            if (isExpenses)
            {
                int monthsRatio = Math.Clamp(quarter, 1, 4) * 3;
                string sql = $@"
                    SELECT 
                        a.AccNo, 
                        a.AccTitle, 
                        a.Type AS HeadCode,
                        ISNULL(h.Head, '') AS HeadName,
                        (ISNULL(e.Proposed, 0) * {monthsRatio}) / 12.0 AS Proposed,
                        (ISNULL(e.Revised, 0) * {monthsRatio}) / 12.0 AS Revised,
                        ISNULL((
                            SELECT SUM(v.Debit - v.Credit)
                            FROM Vouchers v
                            WHERE v.Accno = a.AccNo
                              AND v.VDate >= @StartDate AND v.VDate <= @EndDate
                        ), 0) AS Actual
                    FROM Accounts a
                    LEFT JOIN Heads h ON a.Type = h.Code
                    LEFT JOIN FinancialBudgeting_Expenses e ON a.AccNo = e.AccNo AND e.FinancialYear = @FinancialYear
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND (TRY_CAST(a.Type AS INT) >= 41 OR LEFT(a.AccNo, 6) = '12-001')
                      " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                    ORDER BY a.AccNo";

                var (qStart, qEnd) = GetQuarterDates(financialYear, quarter);

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                cmd.Parameters.AddWithValue("@StartDate", qStart);
                cmd.Parameters.AddWithValue("@EndDate", qEnd);
                if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
                {
                    cmd.Parameters.AddWithValue("@HeadCode", headCode);
                }

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    list.Add(new BudgetReportRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        HeadCode = reader["HeadCode"]?.ToString() ?? "",
                        HeadName = reader["HeadName"]?.ToString() ?? "",
                        ProposedAmount = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                        RevisedAmount = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0,
                        ActualAmount = reader["Actual"] != DBNull.Value ? Convert.ToDecimal(reader["Actual"]) : 0
                    });
                }
            }
            else
            {
                string propCols = quarter switch
                {
                    1 => "ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0)",
                    2 => "ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0) + ISNULL(b.Month_October, 0) + ISNULL(b.Month_November, 0) + ISNULL(b.Month_December, 0)",
                    3 => "ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0) + ISNULL(b.Month_October, 0) + ISNULL(b.Month_November, 0) + ISNULL(b.Month_December, 0) + ISNULL(b.Month_January, 0) + ISNULL(b.Month_February, 0) + ISNULL(b.Month_March, 0)",
                    4 => "ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0) + ISNULL(b.Month_October, 0) + ISNULL(b.Month_November, 0) + ISNULL(b.Month_December, 0) + ISNULL(b.Month_January, 0) + ISNULL(b.Month_February, 0) + ISNULL(b.Month_March, 0) + ISNULL(b.Month_April, 0) + ISNULL(b.Month_May, 0) + ISNULL(b.Month_June, 0)",
                    _ => "ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0)"
                };

                string revCols = quarter switch
                {
                    1 => "ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0)",
                    2 => "ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0) + ISNULL(b.Month_October_R, 0) + ISNULL(b.Month_November_R, 0) + ISNULL(b.Month_December_R, 0)",
                    3 => "ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0) + ISNULL(b.Month_October_R, 0) + ISNULL(b.Month_November_R, 0) + ISNULL(b.Month_December_R, 0) + ISNULL(b.Month_January_R, 0) + ISNULL(b.Month_February_R, 0) + ISNULL(b.Month_March_R, 0)",
                    4 => "ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0) + ISNULL(b.Month_October_R, 0) + ISNULL(b.Month_November_R, 0) + ISNULL(b.Month_December_R, 0) + ISNULL(b.Month_January_R, 0) + ISNULL(b.Month_February_R, 0) + ISNULL(b.Month_March_R, 0) + ISNULL(b.Month_April_R, 0) + ISNULL(b.Month_May_R, 0) + ISNULL(b.Month_June_R, 0)",
                    _ => "ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0)"
                };

                string sql = $@"
                    SELECT 
                        a.AccNo, 
                        a.AccTitle, 
                        a.Type AS HeadCode,
                        ISNULL(h.Head, '') AS HeadName,
                        ({propCols}) AS Proposed,
                        ({revCols}) AS Revised,
                        ISNULL((
                            SELECT SUM(v.Credit - v.Debit)
                            FROM Vouchers v
                            WHERE v.Accno = a.AccNo
                              AND v.VDate >= @StartDate AND v.VDate <= @EndDate
                        ), 0) AS Actual
                    FROM Accounts a
                    LEFT JOIN Heads h ON a.Type = h.Code
                    LEFT JOIN FinancialBudgeting b ON a.AccNo = b.AccNo AND b.FinancialYear = @FinancialYear
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND a.Type IN ('31', '32', '33')
                      " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                    ORDER BY a.AccNo";

                var (qStart, qEnd) = GetQuarterDates(financialYear, quarter);

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                cmd.Parameters.AddWithValue("@StartDate", qStart);
                cmd.Parameters.AddWithValue("@EndDate", qEnd);
                if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
                {
                    cmd.Parameters.AddWithValue("@HeadCode", headCode);
                }

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    list.Add(new BudgetReportRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        HeadCode = reader["HeadCode"]?.ToString() ?? "",
                        HeadName = reader["HeadName"]?.ToString() ?? "",
                        ProposedAmount = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                        RevisedAmount = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0,
                        ActualAmount = reader["Actual"] != DBNull.Value ? Convert.ToDecimal(reader["Actual"]) : 0
                    });
                }
            }

            return list;
        }

        public async Task<List<BudgetReportRow>> GetDateRangeReportDataAsync(string financialYear, DateTime fromDate, DateTime toDate, string headCode, bool isExpenses)
        {
            var list = new List<BudgetReportRow>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            int monthsCount = ((toDate.Year - fromDate.Year) * 12) + toDate.Month - fromDate.Month + 1;
            if (monthsCount < 1) monthsCount = 1;

            if (isExpenses)
            {
                string sql = @"
                    SELECT 
                        a.AccNo, 
                        a.AccTitle, 
                        a.Type AS HeadCode,
                        ISNULL(h.Head, '') AS HeadName,
                        (ISNULL(e.Proposed, 0) / 12.0) * @MonthsCount AS Proposed,
                        (ISNULL(e.Revised, 0) / 12.0) * @MonthsCount AS Revised,
                        ISNULL((
                            SELECT SUM(v.Debit - v.Credit)
                            FROM Vouchers v
                            WHERE v.Accno = a.AccNo
                              AND v.VDate >= @StartDate AND v.VDate <= @EndDate
                        ), 0) AS Actual
                    FROM Accounts a
                    LEFT JOIN Heads h ON a.Type = h.Code
                    LEFT JOIN FinancialBudgeting_Expenses e ON a.AccNo = e.AccNo AND e.FinancialYear = @FinancialYear
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND (TRY_CAST(a.Type AS INT) >= 41 OR LEFT(a.AccNo, 6) = '12-001')
                      " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                    ORDER BY a.AccNo";

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                cmd.Parameters.AddWithValue("@StartDate", fromDate);
                cmd.Parameters.AddWithValue("@EndDate", toDate);
                cmd.Parameters.AddWithValue("@MonthsCount", monthsCount);
                if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
                {
                    cmd.Parameters.AddWithValue("@HeadCode", headCode);
                }

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    list.Add(new BudgetReportRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        HeadCode = reader["HeadCode"]?.ToString() ?? "",
                        HeadName = reader["HeadName"]?.ToString() ?? "",
                        ProposedAmount = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                        RevisedAmount = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0,
                        ActualAmount = reader["Actual"] != DBNull.Value ? Convert.ToDecimal(reader["Actual"]) : 0
                    });
                }
            }
            else
            {
                var monthCols = BuildMonthColumnsForDateRange(fromDate, toDate, false);
                var monthRevCols = BuildMonthColumnsForDateRange(fromDate, toDate, true);

                string sql = $@"
                    SELECT 
                        a.AccNo, 
                        a.AccTitle, 
                        a.Type AS HeadCode,
                        ISNULL(h.Head, '') AS HeadName,
                        ({monthCols}) AS Proposed,
                        ({monthRevCols}) AS Revised,
                        ISNULL((
                            SELECT SUM(v.Credit - v.Debit)
                            FROM Vouchers v
                            WHERE v.Accno = a.AccNo
                              AND v.VDate >= @StartDate AND v.VDate <= @EndDate
                        ), 0) AS Actual
                    FROM Accounts a
                    LEFT JOIN Heads h ON a.Type = h.Code
                    LEFT JOIN FinancialBudgeting b ON a.AccNo = b.AccNo AND b.FinancialYear = @FinancialYear
                    WHERE a.BudgetAccount = 1 AND a.Active = 1
                      AND a.Type IN ('31', '32', '33')
                      " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                    ORDER BY a.AccNo";

                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
                cmd.Parameters.AddWithValue("@StartDate", fromDate);
                cmd.Parameters.AddWithValue("@EndDate", toDate);
                if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
                {
                    cmd.Parameters.AddWithValue("@HeadCode", headCode);
                }

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    list.Add(new BudgetReportRow
                    {
                        AccNo = reader["AccNo"]?.ToString() ?? "",
                        AccTitle = reader["AccTitle"]?.ToString() ?? "",
                        HeadCode = reader["HeadCode"]?.ToString() ?? "",
                        HeadName = reader["HeadName"]?.ToString() ?? "",
                        ProposedAmount = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                        RevisedAmount = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0,
                        ActualAmount = reader["Actual"] != DBNull.Value ? Convert.ToDecimal(reader["Actual"]) : 0
                    });
                }
            }

            return list;
        }

        public async Task<List<BudgetReportRow>> GetBudgetAtGlanceDataAsync(string financialYear, string headCode)
        {
            var list = new List<BudgetReportRow>();

            using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync();

            string sql = @"
                SELECT 
                    a.AccNo, 
                    a.AccTitle, 
                    a.Type AS HeadCode,
                    ISNULL(h.Head, '') AS HeadName,
                    (ISNULL(b.Month_July, 0) + ISNULL(b.Month_August, 0) + ISNULL(b.Month_September, 0) +
                     ISNULL(b.Month_October, 0) + ISNULL(b.Month_November, 0) + ISNULL(b.Month_December, 0) +
                     ISNULL(b.Month_January, 0) + ISNULL(b.Month_February, 0) + ISNULL(b.Month_March, 0) +
                     ISNULL(b.Month_April, 0) + ISNULL(b.Month_May, 0) + ISNULL(b.Month_June, 0)) AS Proposed,
                    (ISNULL(b.Month_July_R, 0) + ISNULL(b.Month_August_R, 0) + ISNULL(b.Month_September_R, 0) +
                     ISNULL(b.Month_October_R, 0) + ISNULL(b.Month_November_R, 0) + ISNULL(b.Month_December_R, 0) +
                     ISNULL(b.Month_January_R, 0) + ISNULL(b.Month_February_R, 0) + ISNULL(b.Month_March_R, 0) +
                     ISNULL(b.Month_April_R, 0) + ISNULL(b.Month_May_R, 0) + ISNULL(b.Month_June_R, 0)) AS Revised,
                    ISNULL((
                        SELECT SUM(v.Credit - v.Debit)
                        FROM Vouchers v
                        WHERE v.Accno = a.AccNo
                    ), 0) AS Actual
                FROM Accounts a
                LEFT JOIN Heads h ON a.Type = h.Code
                LEFT JOIN FinancialBudgeting b ON a.AccNo = b.AccNo AND b.FinancialYear = @FinancialYear
                WHERE a.BudgetAccount_Glance = 1 AND a.Active = 1
                  " + (headCode != "0" && !string.IsNullOrWhiteSpace(headCode) ? " AND a.Type = @HeadCode" : "") + @"
                ORDER BY a.AccNo";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@FinancialYear", financialYear);
            if (headCode != "0" && !string.IsNullOrWhiteSpace(headCode))
            {
                cmd.Parameters.AddWithValue("@HeadCode", headCode);
            }

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new BudgetReportRow
                {
                    AccNo = reader["AccNo"]?.ToString() ?? "",
                    AccTitle = reader["AccTitle"]?.ToString() ?? "",
                    HeadCode = reader["HeadCode"]?.ToString() ?? "",
                    HeadName = reader["HeadName"]?.ToString() ?? "",
                    ProposedAmount = reader["Proposed"] != DBNull.Value ? Convert.ToDecimal(reader["Proposed"]) : 0,
                    RevisedAmount = reader["Revised"] != DBNull.Value ? Convert.ToDecimal(reader["Revised"]) : 0,
                    ActualAmount = reader["Actual"] != DBNull.Value ? Convert.ToDecimal(reader["Actual"]) : 0
                });
            }

            return list;
        }

        private static (DateTime Start, DateTime End) GetQuarterDates(string financialYear, int quarter)
        {
            int startYear = int.TryParse(financialYear.Split('-')[0].Trim(), out int y) ? y : DateTime.Today.Year;
            return quarter switch
            {
                1 => (new DateTime(startYear, 7, 1), new DateTime(startYear, 9, 30)),
                2 => (new DateTime(startYear, 7, 1), new DateTime(startYear, 12, 31)),
                3 => (new DateTime(startYear, 7, 1), new DateTime(startYear + 1, 3, 31)),
                4 => (new DateTime(startYear, 7, 1), new DateTime(startYear + 1, 6, 30)),
                _ => (new DateTime(startYear, 7, 1), new DateTime(startYear + 1, 6, 30))
            };
        }

        private static string BuildMonthColumnsForDateRange(DateTime from, DateTime to, bool isRevision)
        {
            var cols = new List<string>();
            var cur = new DateTime(from.Year, from.Month, 1);
            var end = new DateTime(to.Year, to.Month, 1);

            while (cur <= end)
            {
                string colName = "Month_" + cur.ToString("MMMM", CultureInfo.InvariantCulture);
                if (isRevision) colName += "_R";
                cols.Add($"ISNULL(b.{colName}, 0)");
                cur = cur.AddMonths(1);
            }

            return cols.Count > 0 ? string.Join(" + ", cols) : "0";
        }
    }
}

