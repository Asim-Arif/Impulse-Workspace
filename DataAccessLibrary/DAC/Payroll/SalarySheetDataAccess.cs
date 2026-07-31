using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class SalarySheetDataAccess : ISalarySheetDataAccess
    {
        private readonly string _connectionString;

        public SalarySheetDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<bool> IsSalarySheetFinalizedAsync(int year, int month, bool isExternal)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT COUNT(*)
                FROM MonthlySalaries
                WHERE MONTH(DT) = @Month AND YEAR(DT) = @Year AND IsForSA = @IsForSA";

            int count = await db.ExecuteScalarAsync<int>(sql, new
            {
                Month   = month,
                Year    = year,
                IsForSA = isExternal ? 1 : 0
            });

            return count > 0;
        }

        public async Task<List<SalarySheetRowModel>> GetFinalizedSalarySheetAsync(SalarySheetFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            string sql = @"
                SELECT
                    v.EmpID,
                    ISNULL(e.Name, v.EmpID) AS EmpName,
                    v.DeptID,
                    ISNULL(d.Name, v.DeptID) AS DeptName,
                    e.Designation,
                    ISNULL(v.BSal, 0) AS BasicSalary,
                    ISNULL(v.Rate, 0) AS DailyRate,
                    ISNULL(v.TotalMonthHrs, 0) AS MonthDays,
                    ISNULL(v.presentDays, 0) AS PresentDays,
                    ISNULL(v.ADays, 0) / @HrsMultiple AS AbsentDays,
                    ISNULL(v.Leaves, 0) / @HrsMultiple AS LeaveDays,
                    ISNULL(v.SDays, 0) / @HrsMultiple AS EarnedDays,
                    ISNULL(v.SAmt, 0) AS EarnedSalary,
                    -- Legacy stores OT/Late hours as 'HH:MM' strings (MinsToTime), e.g. '29:09', '-02:-32'
                    CASE
                        WHEN ISNUMERIC(v.OHrs) = 1 THEN CAST(v.OHrs AS float)
                        WHEN CHARINDEX(':', v.OHrs) > 0 THEN
                            (CASE WHEN LEFT(v.OHrs, 1) = '-' THEN -1.0 ELSE 1.0 END) *
                            (ABS(CAST(LEFT(v.OHrs, CHARINDEX(':', v.OHrs) - 1) AS float)) +
                             ABS(CAST(SUBSTRING(v.OHrs, CHARINDEX(':', v.OHrs) + 1, 10) AS float)) / 60.0)
                        ELSE 0
                    END AS OTHours,
                    ISNULL(v.OAmt, 0) AS OTAmount,
                    ISNULL(v.AAllowAmt, 0) AS Allowances,
                    CASE
                        WHEN ISNUMERIC(v.lHrs) = 1 THEN CAST(v.lHrs AS float)
                        WHEN CHARINDEX(':', v.lHrs) > 0 THEN
                            (CASE WHEN LEFT(v.lHrs, 1) = '-' THEN -1.0 ELSE 1.0 END) *
                            (ABS(CAST(LEFT(v.lHrs, CHARINDEX(':', v.lHrs) - 1) AS float)) +
                             ABS(CAST(SUBSTRING(v.lHrs, CHARINDEX(':', v.lHrs) + 1, 10) AS float)) / 60.0)
                        ELSE 0
                    END AS ShortHours,
                    -- Legacy splits late (LAmt) and short (ShortHrsAmt) amounts; grid shows them combined
                    ISNULL(v.lAmt, 0) + ISNULL(v.ShortHrsAmt, 0) AS ShortHoursDeduction,
                    ISNULL(v.total, 0) AS GrossEarnings,
                    ISNULL(v.Tax, 0) AS TaxDeduction,
                    ISNULL(v.NetTtl, 0) AS NetEarnings,
                    ISNULL(v.ShortTerm, 0) AS ShortTermAdvanceDeduction,
                    ISNULL(v.LongTerm, 0) AS LongTermLoanDeduction,
                    ISNULL(v.AdvSal, 0) AS AdvSalaryPaid,
                    ISNULL(v.UnionFund, 0) AS UnionFundDeduction,
                    ISNULL(v.Fine, 0) AS FineDeduction,
                    ISNULL(v.EOBI, 0) AS EOBIDeduction,
                    ISNULL(v.GPHrsAmt, 0) AS GPHrsDeduction,
                    ISNULL(v.Balance, 0) AS NetPayable,
                    ISNULL(v.Salary_Paid, 0) AS IsPaid,
                    ISNULL(v.PrevLTLoan, 0) AS PrevLTLoanBalance,
                    1 AS IsFinalized
                FROM VMonthlySalaries v
                LEFT JOIN Employees e ON v.EmpID = e.EmpID
                LEFT JOIN Departments d ON v.DeptID = d.DeptID
                WHERE MONTH(v.DT) = @Month AND YEAR(v.DT) = @Year AND v.IsForSA = @IsForSA";

            var parameters = new DynamicParameters();
            parameters.Add("Month", filter.Month);
            parameters.Add("Year", filter.Year);
            parameters.Add("IsForSA", filter.IsExternal ? 1 : 0);
            // Legacy stores ADays/Leaves/SDays in hours on the external sheet (HrsMultiple = 8)
            parameters.Add("HrsMultiple", filter.IsExternal ? 8 : 1);

            if (!string.IsNullOrEmpty(filter.DeptID) && filter.DeptID != "0")
            {
                sql += " AND v.DeptID = @DeptID";
                parameters.Add("DeptID", filter.DeptID);
            }

            if (!string.IsNullOrEmpty(filter.EmpID) && filter.EmpID != "0")
            {
                sql += " AND v.EmpID = @EmpID";
                parameters.Add("EmpID", filter.EmpID);
            }

            sql += " ORDER BY v.DeptID, v.EmpID ASC";

            var result = await db.QueryAsync<SalarySheetRowModel>(sql, parameters);
            return result.ToList();
        }

        public async Task<List<SalarySheetRowModel>> GetUnfinalizedSalarySheetAsync(SalarySheetFilterDto filter)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            // Fetch department name map to ensure DeptName is never blank
            const string deptSql = "SELECT DeptID, Name FROM Departments";
            var deptMap = (await db.QueryAsync<(string DeptID, string Name)>(deptSql))
                .ToDictionary(x => x.DeptID, x => x.Name, StringComparer.OrdinalIgnoreCase);

            DateTime targetDate = new DateTime(filter.Year, filter.Month, DateTime.DaysInMonth(filter.Year, filter.Month));

            string spName = (filter.DeptID != "0" && filter.EmpID == "0") ? "VAdvanceSalaryDeptwise" : "VAdvanceSalary";

            var parameters = new DynamicParameters();
            parameters.Add("@ParDate", targetDate);

            if (spName == "VAdvanceSalaryDeptwise")
            {
                parameters.Add("@DeptID", filter.DeptID);
            }
            else
            {
                parameters.Add("@EmpID", filter.EmpID == "0" ? "%" : filter.EmpID);
            }

            var rawRows = await db.QueryAsync<dynamic>(spName, parameters, commandType: CommandType.StoredProcedure);

            int monthDays = DateTime.DaysInMonth(filter.Year, filter.Month);
            List<SalarySheetRowModel> list = new List<SalarySheetRowModel>();

            // Cache of Holidays-table counts keyed by TotalSalDays (same for most employees)
            var holidayCountCache = new Dictionary<int, int>();

            foreach (var r in rawRows)
            {
                var dict = (IDictionary<string, object>)r;

                string empId       = GetStringValue(dict, "EmpID", "EMPID");
                string empName     = GetStringValue(dict, "Name", "EmpName", "EMP_NAME", "EmployeeName");
                string deptId      = GetStringValue(dict, "DeptID", "DEPTID", "Dept_ID");
                string deptName    = GetStringValue(dict, "DeptName", "Department", "DEPTNAME", "Dept_Name");
                string designation = GetStringValue(dict, "Designation", "DESIGNATION");

                if (string.IsNullOrEmpty(deptName) && !string.IsNullOrEmpty(deptId) && deptMap.ContainsKey(deptId))
                {
                    deptName = deptMap[deptId];
                }

                // Legacy (frmSalarySheet.frm) computes Holidays in form code — the SP does NOT return it:
                // Sundays up to TotalSalDays + Holidays-table entries.
                // (Legacy's Sunday-absents adjustment is effectively always 0, so it is not ported.)
                DateTime? joinDate       = GetDateValue(dict, "JoinDate");
                DateTime? discontinuedOn = GetDateValue(dict, "DiscontinuedOn");

                int totalSalDays = monthDays;
                if (discontinuedOn.HasValue && discontinuedOn.Value.Year == filter.Year
                    && discontinuedOn.Value.Month == filter.Month && discontinuedOn.Value <= targetDate)
                {
                    totalSalDays = discontinuedOn.Value.Day;
                }

                bool joinedThisMonth = joinDate.HasValue
                    && joinDate.Value.Year == filter.Year && joinDate.Value.Month == filter.Month;

                int holidayCount;
                if (joinedThisMonth)
                {
                    holidayCount = await db.ExecuteScalarAsync<int>(
                        "SELECT COUNT(DT) FROM Holidays WHERE DT BETWEEN @FromDT AND @ToDT",
                        new { FromDT = joinDate.Value, ToDT = targetDate });
                }
                else
                {
                    if (!holidayCountCache.TryGetValue(totalSalDays, out holidayCount))
                    {
                        holidayCount = await db.ExecuteScalarAsync<int>(
                            "SELECT COUNT(DT) FROM Holidays WHERE MONTH(DT) = @Month AND YEAR(DT) = @Year AND DAY(DT) <= @TotalSalDays",
                            new { Month = filter.Month, Year = filter.Year, TotalSalDays = totalSalDays });
                        holidayCountCache[totalSalDays] = holidayCount;
                    }
                }

                double holidays = TotalSundays(filter.Year, filter.Month, totalSalDays,
                    joinedThisMonth ? joinDate.Value.Day : 1) + holidayCount;

                decimal basicSal   = GetDecimalValue(dict, "StartingSalary", "BSal", "BasicSalary", "Salary");
                decimal dailyRate  = Math.Round(basicSal / 30.0m, 2);
                // Legacy uses the raw (unrounded) rate for the earned salary calculation
                decimal calcDailyRate = basicSal / 30.0m;
                decimal hourlyRate = dailyRate / 8.0m;

                double presentDays      = GetDoubleValue(dict, "presentDays", "PresentDays", "Presents");
                double absAllowed       = GetDoubleValue(dict, "AbsAllowed", "AbsentAllowed");
                double leaves           = GetDoubleValue(dict, "Leaves", "TotalLeaves", "LeaveDays");
                double approvedLeaves   = GetDoubleValue(dict, "ApprovedLeaves");
                double wpLeaves         = GetDoubleValue(dict, "WPLeaves", "WithoutPayLeaves");
                double halfLeaves       = GetDoubleValue(dict, "HalfLeaves");

                // Legacy: internal sheet counts APPROVED leaves only; external sheet counts all leaves.
                // AbsAllowed is NOT part of the earned days formula (display/absent-allowed amount only).
                double leavesForCalc = filter.IsExternal ? leaves : approvedLeaves;

                // Legacy does not clamp present/absent days in the earned salary block
                double dPresentDays = presentDays + holidays + leavesForCalc - wpLeaves - halfLeaves;
                double dAbsentDays  = monthDays - dPresentDays;

                double dAbsDaysForSalCalc = dAbsentDays;
                bool deductionSpecialAbsents = GetBoolValue(dict, "DeductionSpecialAbsents");
                double spAbsentDays          = GetDoubleValue(dict, "SPAbsentDays", "SpecialAbsentDays");

                if (deductionSpecialAbsents)
                {
                    dPresentDays       = dPresentDays - spAbsentDays;
                    dAbsDaysForSalCalc += spAbsentDays;
                }

                bool payFullSalary = GetBoolValue(dict, "Pay_Full_Salary", "PayFullSalary");

                decimal earnedSalary = 0;
                double earnedDays = 0;

                if (payFullSalary)
                {
                    earnedSalary = basicSal;
                    // Legacy leaves dPresentDays at its computed value for full-salary employees
                    earnedDays   = dPresentDays;
                }
                else
                {
                    if (dPresentDays > dAbsDaysForSalCalc)
                    {
                        earnedSalary = Math.Round(basicSal - (calcDailyRate * (decimal)dAbsDaysForSalCalc), 0);
                        earnedDays   = 30.0 - dAbsDaysForSalCalc;
                    }
                    else
                    {
                        earnedSalary = Math.Round((decimal)dPresentDays * calcDailyRate, 0);
                        earnedDays   = dPresentDays;
                    }
                }

                double otHrs  = GetDoubleValue(dict, "OverTimeHours", "OHrs", "OTHours");
                double otRate = GetDoubleValue(dict, "OTRate", "OverTimeRate");
                if (otRate == 0) otRate = 1.3333;

                decimal otAmt = Math.Round((decimal)otHrs * hourlyRate * (decimal)otRate, 0);

                decimal fixAllow   = GetDecimalValue(dict, "FixAllowance", "FixAllow");
                decimal foodAllow  = GetDecimalValue(dict, "Food_Allowance", "FoodAllowance", "LunchAmt");
                decimal totalAllowances = fixAllow + foodAllow;

                double lateComingHrs      = GetDoubleValue(dict, "LateComingHrs", "lHrs", "LateHrs");
                double lateComingAdjusted = GetDoubleValue(dict, "LateComingHrs_Adjusted");
                double netLateComingHrs   = Math.Max(0, lateComingHrs - lateComingAdjusted);

                double hrsWorked = GetDoubleValue(dict, "HrsWorked", "WorkedHrs", "HOURS", "SalHrs");
                double salHrs = 0;
                if (hrsWorked > 0)
                {
                    salHrs = hrsWorked + (holidays * 8.0) + (leaves * 8.0) - (wpLeaves * 8.0);
                }
                else
                {
                    salHrs = (presentDays + holidays + absAllowed + leaves - wpLeaves) * 8.0;
                }

                double dLateHrs = (monthDays * 8.0) - salHrs - (dAbsentDays * 8.0) - lateComingAdjusted;
                if (dLateHrs < 0) dLateHrs = 0;

                double dShortHrs = dLateHrs - netLateComingHrs;
                if (dShortHrs < 0) dShortHrs = 0;

                // Legacy (frmSalarySheet.frm:2291): full-salary employees get no late/short deduction
                if (payFullSalary)
                {
                    netLateComingHrs = 0;
                    dShortHrs        = 0;
                    dLateHrs         = 0;
                }

                decimal deductionPercentage = GetDecimalValue(dict, "DeductionPercentage");
                if (deductionPercentage == 0) deductionPercentage = 100;

                decimal lateAmt = Math.Round((decimal)netLateComingHrs * hourlyRate * (deductionPercentage / 100m), 0);
                decimal shortAmt = Math.Round((decimal)dShortHrs * hourlyRate, 0);
                decimal totalLateShortDed = lateAmt + shortAmt;

                // ── Legacy-layout values for MonthlySalaries persistence only ────────
                // Everything below mirrors frmSalarySheet.frm display/storage logic and
                // has NO effect on the salary calculation above or the grid.

                // Display absent days (frm:2041-2103) — based on SalDays, not dAbsDaysForSalCalc
                double totalSalDaysDisp = totalSalDays;
                if (filter.IsExternal) totalSalDaysDisp -= holidays; // frm:1862

                int daysToExclude = joinedThisMonth ? joinDate.Value.Day - 1 : 0;

                double salDays = filter.IsExternal
                    ? presentDays + absAllowed               // frm:2051
                    : presentDays + holidays + absAllowed;   // frm:2041
                if (salDays > totalSalDaysDisp) salDays = totalSalDaysDisp;
                if (salDays < 0) salDays = 0;

                // Internal sheet only: unapproved leaves reduce SalDays (frm:2095)
                if (!filter.IsExternal) salDays -= Math.Max(0, leaves - approvedLeaves);

                // Employees with DeductAbsents = 0 are always fully present (frm:2097)
                string deductAbsentsRaw = GetStringValue(dict, "DeductAbsents");
                if (deductAbsentsRaw == "0" || deductAbsentsRaw.Equals("False", StringComparison.OrdinalIgnoreCase))
                {
                    salDays = totalSalDaysDisp - daysToExclude;
                }

                double displayAbsentDays = totalSalDaysDisp - salDays - daysToExclude;
                if (displayAbsentDays < 0) displayAbsentDays = 0;

                // Absent-allowed amount (frm:2137-2144)
                double netAbsAllowed = 0;
                if (displayAbsentDays > 0)
                {
                    netAbsAllowed = absAllowed - displayAbsentDays <= 0 ? absAllowed : absAllowed - displayAbsentDays;
                }
                decimal absentAllowedAmount = Math.Round((decimal)netAbsAllowed * dailyRate, 2);

                // Legacy SalHrs (frm:2158-2207) — external path ported; internal is best-effort
                double legacySalHrs;
                if (filter.IsExternal)
                {
                    if (payFullSalary)
                    {
                        legacySalHrs = monthDays * 8.0; // txtTotalHrs
                    }
                    else
                    {
                        double extraHrsToDeduct = GetDoubleValue(dict, "ExtraHrsToDeduct");
                        legacySalHrs = Math.Round(hrsWorked, 2) + (holidays * 8.0) + (leaves * 8.0)
                                     - (wpLeaves * 8.0) - extraHrsToDeduct;
                        if (legacySalHrs > monthDays * 8.0) legacySalHrs = monthDays * 8.0; // IAA <> Dr-Frgz
                    }
                    if (deductionSpecialAbsents) legacySalHrs -= spAbsentDays * 8.0;
                    if (legacySalHrs < 0) legacySalHrs = 0;
                }
                else
                {
                    legacySalHrs = salDays - wpLeaves - (GetDoubleValue(dict, "SickLeaves", "SL") / 2.0);
                    if (legacySalHrs > monthDays) legacySalHrs = monthDays;
                    if (legacySalHrs < 0) legacySalHrs = 0;
                }

                // Legacy dLateHrs (frm:2264-2270), unclamped — stored as 'HH:MM' in LHrs
                double legacyLateHrs = (monthDays * 8.0) - legacySalHrs - (wpLeaves * 8.0)
                                     - (dAbsentDays * 8.0) - lateComingAdjusted;
                if (deductionSpecialAbsents) legacyLateHrs -= spAbsentDays * 8.0;
                if (payFullSalary) legacyLateHrs = 0;

                // Sunday OT: legacy stores the AMOUNT in SundayOTRate (frm:2366)
                double sundayOTHrs  = GetDoubleValue(dict, "SundayOTHrs");
                double otRateRaw    = GetDoubleValue(dict, "OTRate", "OverTimeRate");
                double sundayOTAmt  = sundayOTHrs * ((double)dailyRate / 8.0) * otRateRaw;

                // Legacy stores NET late-coming hours, rounded to 4dp, unclamped (frm:2215, 2520).
                // Full-salary employees are zeroed (frm:2291-2295 runs before the value is stored).
                double netLateComingHrsRaw = payFullSalary ? 0 : Math.Round(lateComingHrs - lateComingAdjusted, 4);

                // ── End of legacy-layout values ──────────────────────────────────────

                decimal grossEarnings = earnedSalary + otAmt + totalAllowances - totalLateShortDed;

                decimal tax         = GetDecimalValue(dict, "TaxAmt", "Tax", "TaxDeduction");
                decimal netEarnings = grossEarnings - tax;

                // VAdvanceSalary returns Short Term advance deduction in 'Payed' column
                decimal stAdv = GetDecimalValue(dict, "Payed", "ShortTerm", "STAdv", "ShortTermAdvance");

                decimal lDedAmt = GetDecimalValue(dict, "DedAmt");
                decimal lTotAmt = GetDecimalValue(dict, "totamt");
                decimal lClsAmt = GetDecimalValue(dict, "ClsAmt");
                decimal remBal  = lTotAmt - lClsAmt;

                decimal ltLoan = 0;
                decimal amtToDeduct = GetDecimalValue(dict, "AmtToDeduct");
                if (amtToDeduct > 0)
                {
                    ltLoan = amtToDeduct;
                }
                else
                {
                    if (remBal > 0 && lDedAmt > remBal)
                    {
                        ltLoan = remBal;
                    }
                    else
                    {
                        ltLoan = lDedAmt;
                    }
                }

                decimal advSalPaid = GetDecimalValue(dict, "AdvSalpaid", "AdvSal", "AdvanceSalary");
                decimal fine       = GetDecimalValue(dict, "TotalFine", "Fine", "FineDeduction");
                decimal eobi       = GetDecimalValue(dict, "EOBIAmt", "EOBI");

                double gpHrs      = GetDoubleValue(dict, "GPHrs", "GatePassHours");
                decimal gpHrsAmt  = Math.Round((decimal)gpHrs * hourlyRate, 0);

                bool unionFundFlag = GetBoolValue(dict, "UnionFund");
                decimal unionFund  = unionFundFlag ? 20m : 0m;

                decimal netPayable = grossEarnings + fine - (stAdv + ltLoan + advSalPaid + eobi + unionFund + gpHrsAmt);

                list.Add(new SalarySheetRowModel
                {
                    EmpID                      = empId,
                    EmpName                    = !string.IsNullOrEmpty(empName) ? empName : empId,
                    DeptID                     = deptId,
                    DeptName                   = !string.IsNullOrEmpty(deptName) ? deptName : deptId,
                    Designation                = designation,
                    BasicSalary                = basicSal,
                    DailyRate                  = dailyRate,
                    MonthDays                  = monthDays,
                    PresentDays                = presentDays,
                    AbsentDays                 = dAbsDaysForSalCalc,
                    LeaveDays                  = leaves,
                    EarnedDays                 = earnedDays,
                    EarnedSalary               = earnedSalary,
                    OTHours                    = otHrs,
                    OTAmount                   = otAmt,
                    Allowances                 = totalAllowances,
                    ShortHours                 = dShortHrs + netLateComingHrs,
                    ShortHoursDeduction        = totalLateShortDed,
                    GrossEarnings              = grossEarnings,
                    TaxDeduction               = tax,
                    NetEarnings                = netEarnings,
                    ShortTermAdvanceDeduction = stAdv,
                    LongTermLoanDeduction      = ltLoan,
                    AdvSalaryPaid              = advSalPaid,
                    UnionFundDeduction         = unionFund,
                    FineDeduction              = fine,
                    EOBIDeduction              = eobi,
                    GPHrsDeduction             = gpHrsAmt,
                    NetPayable                 = netPayable,
                    IsPaid                     = false,
                    IsFinalized                = false,
                    PrevLTLoanBalance          = Math.Max(remBal, ltLoan),

                    // Complete 45-column MonthlySalaries fields (no calculation impact)
                    CasualLeaves               = GetDoubleValue(dict, "CasualLeaves", "CL"),
                    SickLeaves                 = GetDoubleValue(dict, "SickLeaves", "SL"),
                    AnnualLeaves               = GetDoubleValue(dict, "AnnualLeaves", "AL"),
                    CompensatoryLeaves         = GetDoubleValue(dict, "CompensatoryLeaves", "CPL"),
                    WPLeaves                   = wpLeaves,
                    MaternityLeaves            = GetDoubleValue(dict, "MaternityLeaves", "ML"),
                    AbsentAllowed              = absAllowed,
                    AbsentAllowedAmount        = absentAllowedAmount,
                    AbsentDeductionAmount      = Math.Round((decimal)displayAbsentDays * dailyRate, 0),
                    HrsPerDay                  = filter.IsExternal ? 8.0 : 1.0, // legacy HrsMultiple (frm:2004,2009)
                    TotalMonthHrs              = monthDays * 8,
                    FakeWorkingHrs             = monthDays * 8,
                    FakeRate                   = 0m,
                    FakeSalary                 = 0m,
                    SundayOTHrs                = sundayOTHrs,
                    SundayOTRate               = sundayOTAmt, // legacy stores the Sunday OT AMOUNT here (frm:2366)
                    FixAllowance               = fixAllow,
                    HoldSalaryAmt              = GetDecimalValue(dict, "HoldAmt", "HoldSalaryAmt"),
                    LateComingHrs              = netLateComingHrsRaw,
                    AmtPaid                    = GetDecimalValue(dict, "AmtPaid"), // legacy: SalaryAlreadyPaid only (frm:2519)
                    GPHrs                      = gpHrs,
                    OTDinnerCount              = (int)GetDoubleValue(dict, "DinnerCount", "OTDinnerCount"),
                    OTDinnerAmount             = GetDecimalValue(dict, "LunchAmt", "OTDinnerAmount"),
                    DedOnePercent              = 0m, // legacy hardcodes 0 — 1% deduction is disabled (frm:2469)
                    ZeroAbsentBonus            = GetDecimalValue(dict, "ZeroAbsentBonus"),
                    OTHrs_Original             = GetDoubleValue(dict, "OTHrs_Actual"),
                    LateHrs_Original           = GetDoubleValue(dict, "LateComingHrs_Actual"),
                    OTHrs_Net                  = otHrs,
                    LateHrs_Net                = GetDoubleValue(dict, "LateComingHrs_Actual") - lateComingAdjusted,

                    // Legacy-layout persistence values (used only by FinalizeSalarySheetAsync)
                    DisplayAbsentDays          = displayAbsentDays,
                    LegacySalHrs               = legacySalHrs,
                    LegacyLateHrs              = legacyLateHrs,
                    ShortHoursOnly             = dShortHrs,
                    LateAmount                 = lateAmt,
                    ShortAmount                = shortAmt
                });
            }

            return list;
        }

        // Port of legacy TotalSundays (Salaries.bas) — counts Sundays from startDay to endDay,
        // replicating the legacy quirk of comparing StartDay against the weekday number.
        private static int TotalSundays(int year, int month, int endDay, int startDay = 1)
        {
            int firstDay = (int)new DateTime(year, month, startDay).DayOfWeek + 1; // 1 = Sunday .. 7 = Saturday
            int firstSunday = startDay != firstDay ? startDay + (8 - firstDay) : startDay;

            int count = 0;
            for (int d = firstSunday; d <= endDay; d += 7) count++;
            return count;
        }

        // Port of legacy MinsToTime (PayrollFunctions.bas:27) — hours to 'HH:MM' string,
        // including the legacy negative quirk (e.g. -1.5333 hrs -> "-02:-32").
        private static string MinsToTime(double hours)
        {
            long minutes = (long)Math.Round(hours * 60.0, MidpointRounding.ToEven); // VB double->Long is banker's
            long hrs  = (long)Math.Floor(minutes / 60.0);                           // VB Int()
            long mins = minutes % 60;                                               // VB Mod keeps dividend sign
            return hrs.ToString("00") + ":" + mins.ToString("00");
        }

        private static DateTime? GetDateValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var kvp = dict.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
                if (kvp.Key != null && kvp.Value != null && DateTime.TryParse(kvp.Value.ToString(), out DateTime val))
                {
                    return val;
                }
            }
            return null;
        }

        private static string GetStringValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var kvp = dict.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
                if (kvp.Key != null && kvp.Value != null)
                {
                    return kvp.Value.ToString() ?? "";
                }
            }
            return "";
        }

        private static decimal GetDecimalValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var kvp = dict.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
                if (kvp.Key != null && kvp.Value != null && decimal.TryParse(kvp.Value.ToString(), out decimal val))
                {
                    return val;
                }
            }
            return 0;
        }

        private static double GetDoubleValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var kvp = dict.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
                if (kvp.Key != null && kvp.Value != null && double.TryParse(kvp.Value.ToString(), out double val))
                {
                    return val;
                }
            }
            return 0;
        }

        private static bool GetBoolValue(IDictionary<string, object> dict, params string[] keys)
        {
            foreach (var key in keys)
            {
                var kvp = dict.FirstOrDefault(x => string.Equals(x.Key, key, StringComparison.OrdinalIgnoreCase));
                if (kvp.Key != null && kvp.Value != null)
                {
                    if (bool.TryParse(kvp.Value.ToString(), out bool bVal)) return bVal;
                    if (int.TryParse(kvp.Value.ToString(), out int iVal)) return iVal != 0;
                }
            }
            return false;
        }

        public async Task FinalizeSalarySheetAsync(
            int year,
            int month,
            bool isExternal,
            List<SalarySheetRowModel> rows,
            string userName,
            string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open) db.Open();

            string stAccSql = "SELECT DataValue FROM GeneralData WHERE DataName = 'EmpShortTermAccNo'";
            string ltAccSql = "SELECT DataValue FROM GeneralData WHERE DataName = 'EmpLongTermAccNo'";
            string empShortTermAccNo = (await db.ExecuteScalarAsync<string>(stAccSql)) ?? "";
            string empLongTermAccNo  = (await db.ExecuteScalarAsync<string>(ltAccSql)) ?? "";

            using IDbTransaction trans = db.BeginTransaction();

            try
            {
                DateTime dtMonthEnd = new DateTime(year, month, DateTime.DaysInMonth(year, month));

                // Legacy layout (frmSalarySheet.frm finalize): external sheet stores hours
                // (HrsMultiple = 8), hourly rate and 'HH:MM' time strings; internal stores days.
                decimal hrsMultiple = isExternal ? 8m : 1m;

                // ── 1. Insert into MonthlySalaries ───────────────────────────
                const string insertSalSql = @"
                    INSERT INTO MonthlySalaries
                        (EmpID, DT, BSal, Rate, ADays, AAmt, AAllow, AAllowAmt, SDays, SAmt, Leaves, LeaveAmt, OHrs, OAmt,
                         LHrs, LAmt, Total, Tax, NetTtl, ShortTerm, LongTerm, AdvSal, Unionfund, Fine, Bonus, Lunch, EOBI, Balance,
                         PrevLTLoan, CasualLeaves, SickLeaves, AnnualLeaves, CompensatoryLeaves, WPLeaves, MaternityLeaves,
                         HrsPerDay, TotalMonthHrs, FakeWorkingHrs, FakeRate, FakeSalary, SundayOTHrs, SundayOTRate,
                         FixAllowance, HoldSalaryAmt, PresentDays, LeaveDays, LateComingHrs, ShortHrs, AmtPaid, GPHrs, GPHrsAmt,
                         DeptID, OTDinnerCount, OTDinnerAmount, DedOnePercent, ShortHrsAmt, ZeroAbsentBonus,
                         OTHrs_Original, LateHrs_Original, OTHrs_Net, LateHrs_Net, IsForSA, DTFinal, Salary_Paid)
                    VALUES
                        (@EmpID, @DT, @BSal, @Rate, @ADays, @AAmt, @AAllow, @AAllowAmt, @SDays, @SAmt, @Leaves, @LeaveAmt, @OHrs, @OAmt,
                         @LHrs, @LAmt, @Total, @Tax, @NetTtl, @ShortTerm, @LongTerm, @AdvSal, @Unionfund, @Fine, @Bonus, @Lunch, @EOBI, @Balance,
                         @PrevLTLoan, @CasualLeaves, @SickLeaves, @AnnualLeaves, @CompensatoryLeaves, @WPLeaves, @MaternityLeaves,
                         @HrsPerDay, @TotalMonthHrs, @FakeWorkingHrs, @FakeRate, @FakeSalary, @SundayOTHrs, @SundayOTRate,
                         @FixAllowance, @HoldSalaryAmt, @PresentDays, @LeaveDays, @LateComingHrs, @ShortHrs, @AmtPaid, @GPHrs, @GPHrsAmt,
                         @DeptID, @OTDinnerCount, @OTDinnerAmount, @DedOnePercent, @ShortHrsAmt, @ZeroAbsentBonus,
                         @OTHrs_Original, @LateHrs_Original, @OTHrs_Net, @LateHrs_Net, @IsForSA, CAST(GETDATE() AS DATE), 0)";

                foreach (var row in rows)
                {
                    await db.ExecuteAsync(insertSalSql, new
                    {
                        EmpID              = row.EmpID,
                        DT                 = dtMonthEnd,
                        BSal               = row.BasicSalary,
                        Rate               = Math.Round(row.DailyRate / hrsMultiple, 2, MidpointRounding.AwayFromZero),
                        ADays              = row.DisplayAbsentDays * (double)hrsMultiple,
                        AAmt               = row.AbsentDeductionAmount,
                        AAllow             = row.AbsentAllowed,
                        AAllowAmt          = row.AbsentAllowedAmount,
                        SDays              = row.LegacySalHrs,
                        SAmt               = row.EarnedSalary,
                        Leaves             = row.LeaveDays * (double)hrsMultiple,
                        LeaveAmt           = 0m,
                        OHrs               = MinsToTime(row.OTHours),
                        OAmt               = row.OTAmount,
                        LHrs               = MinsToTime(row.LegacyLateHrs),
                        LAmt               = row.LateAmount,
                        Total              = row.GrossEarnings,
                        Tax                = row.TaxDeduction,
                        NetTtl             = row.NetEarnings,
                        ShortTerm          = row.ShortTermAdvanceDeduction,
                        LongTerm           = row.LongTermLoanDeduction,
                        AdvSal             = row.AdvSalaryPaid,
                        Unionfund          = row.UnionFundDeduction,
                        Fine               = row.FineDeduction,
                        Bonus              = 0m,
                        Lunch              = 0m,
                        EOBI               = row.EOBIDeduction,
                        Balance            = row.NetPayable,
                        PrevLTLoan         = row.PrevLTLoanBalance,
                        CasualLeaves       = row.CasualLeaves,
                        SickLeaves         = row.SickLeaves,
                        AnnualLeaves       = row.AnnualLeaves,
                        CompensatoryLeaves = row.CompensatoryLeaves,
                        WPLeaves           = row.WPLeaves,
                        MaternityLeaves    = row.MaternityLeaves,
                        HrsPerDay          = row.HrsPerDay,
                        TotalMonthHrs      = row.TotalMonthHrs,
                        FakeWorkingHrs     = row.FakeWorkingHrs,
                        FakeRate           = row.FakeRate,
                        FakeSalary         = row.FakeSalary,
                        SundayOTHrs        = row.SundayOTHrs,
                        SundayOTRate       = row.SundayOTRate,
                        FixAllowance       = row.FixAllowance,
                        HoldSalaryAmt      = row.HoldSalaryAmt,
                        PresentDays        = row.EarnedDays - row.LeaveDays, // frm:2532
                        LeaveDays          = row.LeaveDays,
                        LateComingHrs      = row.LateComingHrs,
                        ShortHrs           = row.ShortHoursOnly,
                        AmtPaid            = row.AmtPaid,
                        GPHrs              = row.GPHrs,
                        GPHrsAmt           = row.GPHrsDeduction,
                        DeptID             = row.DeptID,
                        OTDinnerCount      = row.OTDinnerCount,
                        OTDinnerAmount     = row.OTDinnerAmount,
                        DedOnePercent      = row.DedOnePercent,
                        ShortHrsAmt        = row.ShortAmount,
                        ZeroAbsentBonus    = row.ZeroAbsentBonus,
                        OTHrs_Original     = row.OTHrs_Original,
                        LateHrs_Original   = row.LateHrs_Original,
                        OTHrs_Net          = row.OTHrs_Net,
                        LateHrs_Net        = row.LateHrs_Net,
                        IsForSA            = isExternal ? 1 : 0
                    }, transaction: trans);

                    // ── 2. Update AmtCleared for Long Term loans ─────────────
                    if (row.LongTermLoanDeduction > 0)
                    {
                        await db.ExecuteAsync("DELETE FROM AmtCleared WHERE EmpID = @EmpID AND DT = @DT", new { EmpID = row.EmpID, DT = dtMonthEnd }, transaction: trans);
                        await db.ExecuteAsync("INSERT INTO AmtCleared (EmpID, DT, AmtClrd) VALUES (@EmpID, @DT, @AmtClrd)", new { EmpID = row.EmpID, DT = dtMonthEnd, AmtClrd = row.LongTermLoanDeduction }, transaction: trans);
                    }
                }

                // ── 3. Update Short Term advances cleared ────────────────────
                const string updateSTAdvSql = @"
                    UPDATE Advances
                    SET Cleared = 1
                    WHERE MONTH(DeductionStartDT) = @Month AND YEAR(DeductionStartDT) = @Year AND Type = 0 AND Cleared = 0";

                await db.ExecuteAsync(updateSTAdvSql, new { Month = month, Year = year }, transaction: trans);

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task UpdateSalaryPaidStatusAsync(string empId, int year, int month, bool isPaid, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                UPDATE MonthlySalaries
                SET Salary_Paid = @IsPaid,
                    Salary_Paid_UserName = @UserName,
                    Salary_Paid_MachineName = @MachineName,
                    Salary_Paid_DTEntry = GETDATE()
                WHERE EmpID = @EmpID AND MONTH(DT) = @Month AND YEAR(DT) = @Year";

            await db.ExecuteAsync(sql, new
            {
                IsPaid      = isPaid ? 1 : 0,
                UserName    = userName,
                MachineName = machineName,
                EmpID       = empId,
                Month       = month,
                Year        = year
            });
        }

        public async Task UpdateLongTermDeductionOverrideAsync(string empId, DateTime monthDate, decimal amount, string userName, string machineName)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string updateSql = @"
                UPDATE LongTermDeductionsMonthly
                SET AmtToDeduct = @Amount, DTEntry = GETDATE(), UserName = @UserName, MachineName = @MachineName
                WHERE EmpID = @EmpID AND DedMonth = @DedMonth";

            int rows = await db.ExecuteAsync(updateSql, new
            {
                Amount      = amount,
                UserName    = userName,
                MachineName = machineName,
                EmpID       = empId,
                DedMonth    = monthDate.Date
            });

            if (rows == 0)
            {
                const string insertSql = @"
                    INSERT INTO LongTermDeductionsMonthly (DedMonth, EmpID, AmtToDeduct, UserName, MachineName)
                    VALUES (@DedMonth, @EmpID, @Amount, @UserName, @MachineName)";

                await db.ExecuteAsync(insertSql, new
                {
                    DedMonth    = monthDate.Date,
                    EmpID       = empId,
                    Amount      = amount,
                    UserName    = userName,
                    MachineName = machineName
                });
            }
        }

        public async Task TransferNegativeSalaryAsync(string empId, decimal amountToAdjust, DateTime monthDate)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open) db.Open();
            using IDbTransaction trans = db.BeginTransaction();

            try
            {
                decimal amountLeft = amountToAdjust;
                int month = monthDate.Month;
                int year = monthDate.Year;

                const string selectSql = @"
                    SELECT EntryID, Amount 
                    FROM Advances 
                    WHERE EmpID = @EmpID 
                      AND MONTH(DeductionStartDT) = @Month 
                      AND YEAR(DeductionStartDT) = @Year 
                      AND Type = 0 
                    ORDER BY Amount DESC";

                var advances = (await db.QueryAsync<(long EntryID, decimal Amount)>(
                    selectSql, 
                    new { EmpID = empId, Month = month, Year = year }, 
                    transaction: trans)).ToList();

                foreach (var adv in advances)
                {
                    decimal tempAmt;
                    if (adv.Amount >= amountLeft)
                    {
                        tempAmt = amountLeft;
                        amountLeft = 0;
                    }
                    else
                    {
                        tempAmt = adv.Amount;
                        amountLeft -= tempAmt;
                    }

                    decimal remain = adv.Amount - tempAmt;
                    string desc = $"Amount Adjusted for Negative Salary, Orignal Amt={adv.Amount}, Amt Transferred={tempAmt} Remain={remain}";

                    const string updateSql = @"
                        UPDATE Advances 
                        SET Amount = @NewAmount, SoftwareDescription = @Desc 
                        WHERE EntryID = @EntryID";

                    await db.ExecuteAsync(updateSql, new { NewAmount = remain, Desc = desc, EntryID = adv.EntryID }, transaction: trans);

                    if (amountLeft == 0)
                        break;
                }

                DateTime nextMonth = new DateTime(year, month, 1).AddMonths(1);
                decimal transferredAmt = amountToAdjust - amountLeft;

                const string insertSql = @"
                    INSERT INTO Advances (EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, AccVoucherNo, DeductionStartDT)
                    VALUES (@EmpID, @DT, 0, 'Negative Salary Adjustment Entry', @Amount, 0, 0, '', @DT)";

                await db.ExecuteAsync(insertSql, new { EmpID = empId, DT = nextMonth, Amount = transferredAmt }, transaction: trans);

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<List<DepartmentSalaryVoucherSummaryDto>> GetDepartmentSalaryVoucherSummariesAsync(int year, int month)
        {
            using IDbConnection db = new SqlConnection(_connectionString);

            const string sql = @"
                SELECT
                    d.Name,
                    ISNULL(d.AccruedAccNo, '') AS AccruedAccNo,
                    ISNULL(d.SalaryAccNo, '') AS SalaryAccNo,
                    ISNULL(d.AccruedFoodAccNo, '') AS AccruedFoodAccNo,
                    ISNULL(d.AccruedUnionAccNo, '') AS AccruedUnionAccNo,
                    ISNULL(d.SchAccNo, '') AS SchAccNo,
                    ISNULL(d.AccruedEOBIAccNo, '') AS AccruedEOBIAccNo,
                    ISNULL(d.TaxAccNo, '') AS TaxAccNo,
                    ISNULL(d.OnePercentAccNo, '') AS OnePercentAccNo,
                    t.DeptID,
                    t.TotalAmt,
                    t.SchAmt,
                    t.LunchAmt,
                    t.UnionAmt,
                    t.BalanceAmt,
                    t.ShortAmt,
                    t.LongAmt,
                    t.EOBIAmt,
                    t.TaxAmt,
                    t.OTAmt,
                    t.DedOnePercentAmt,
                    t.ZeroAbsentBonus,
                    t.GPHrsAmt
                FROM (
                    SELECT
                        DeptID,
                        SUM(ISNULL(Total,0) - ISNULL(OAmt,0) - ISNULL(ZeroAbsentBonus,0) - ISNULL(GPHrsAmt,0)) AS TotalAmt,
                        SUM(ISNULL(Bonus,0)) AS SchAmt,
                        SUM(ISNULL(Lunch,0)) AS LunchAmt,
                        SUM(ISNULL(UnionFund,0)) AS UnionAmt,
                        SUM(ISNULL(Balance,0)) AS BalanceAmt,
                        SUM(ISNULL(ShortTerm,0)) AS ShortAmt,
                        SUM(ISNULL(LongTerm,0)) AS LongAmt,
                        SUM(ISNULL(EOBI,0)) AS EOBIAmt,
                        SUM(ISNULL(Tax,0)) AS TaxAmt,
                        SUM(ISNULL(OAmt,0)) AS OTAmt,
                        SUM(ISNULL(Fine,0)) AS DedOnePercentAmt,
                        SUM(ISNULL(ZeroAbsentBonus,0)) AS ZeroAbsentBonus,
                        SUM(ISNULL(GPHrsAmt,0)) AS GPHrsAmt
                    FROM VMonthlySalaries
                    WHERE MONTH(DT) = @Month AND YEAR(DT) = @Year
                    GROUP BY DeptID
                ) t
                INNER JOIN Departments d ON t.DeptID = d.DeptID
                WHERE ISNULL(d.AccruedAccNo, '') <> ''
                  AND ISNULL(d.SalaryAccNo, '') <> ''
                  AND ISNULL(d.SchAccNo, '') <> ''
                  AND ISNULL(d.OnePercentAccNo, '') <> ''";

            var list = await db.QueryAsync<DepartmentSalaryVoucherSummaryDto>(sql, new { Month = month, Year = year });
            return list.AsList();
        }

        public async Task<(string ShortTermAccNo, string LongTermAccNo)> GetSalaryVoucherGeneralAccountsAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            string stAccSql = "SELECT ISNULL(DataValue, '') FROM GeneralData WHERE DataName = 'EmpShortTermAccNo'";
            string ltAccSql = "SELECT ISNULL(DataValue, '') FROM GeneralData WHERE DataName = 'EmpLongTermAccNo'";
            string stAccNo = (await db.ExecuteScalarAsync<string>(stAccSql)) ?? "";
            string ltAccNo = (await db.ExecuteScalarAsync<string>(ltAccSql)) ?? "";
            return (stAccNo, ltAccNo);
        }
    }
}
