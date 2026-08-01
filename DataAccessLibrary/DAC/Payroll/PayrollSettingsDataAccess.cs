using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Payroll
{
    public interface IPayrollSettingsDataAccess
    {
        Task<GeneralSettingsModel> GetSettingsByDeptAsync(string deptId);
        Task<bool> SaveSettingsAsync(GeneralSettingsModel model);
        Task<bool> UpdateAllDepartmentTimingsAsync(string enterTime, string exitTime, string lunchInTime, string lunchOutTime, string fridayBreakStart, string fridayBreakEnd, string fridayCheckOut);
        Task<List<EmpAttCategoryModel>> GetAttendanceCategoriesAsync();
        Task<List<CategoryRuleItem>> GetCategoryRulesByIdAsync(int categoryId);
    }

    public class PayrollSettingsDataAccess : IPayrollSettingsDataAccess
    {
        private readonly string _connectionString;

        public PayrollSettingsDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<GeneralSettingsModel> GetSettingsByDeptAsync(string deptId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    Entryid,
                    DeptID,
                    ISNULL(EnterTime, '08:00') AS EnterTime,
                    ISNULL(ExitTime, '17:00') AS ExitTime,
                    ISNULL(LunchTime, 1.0) AS LunchTime,
                    ISNULL(LunchInTime, '13:00') AS LunchInTime,
                    ISNULL(LunchOutTime, '14:00') AS LunchOutTime,
                    ISNULL(FridayBreakStartTime, '12:30') AS FridayBreakStartTime,
                    ISNULL(FridayBreakEndTime, '14:00') AS FridayBreakEndTime,
                    ISNULL(FridayCheckOut, '17:30') AS FridayCheckOut,
                    ISNULL(WorkingHrs, 8.0) AS WorkingHrs,
                    ISNULL(EmpAttCat_RefID, 1) AS EmpAttCat_RefID,
                    ISNULL(AbsAllowed, 0) AS AbsAllowed,
                    ISNULL(TotalAbsAllowed, 0) AS TotalAbsAllowed,
                    ISNULL(TimeAllowed, 0) AS TimeAllowed,
                    ISNULL(NTimes, 0) AS NTimes,
                    ISNULL(OverTime, 5) AS OverTime,
                    ISNULL(SocialAmt, 0) AS SocialAmt,
                    ISNULL(DinnerAmt, 0) AS DinnerAmt,
                    ISNULL(ShortTermPer, 0) AS ShortTermPer,
                    ISNULL(AdvSalPer, 0) AS AdvSalPer,
                    ISNULL(LongTermTime, 0) AS LongTermTime,
                    ISNULL(LongTermLimit, '0') AS LongTermLimit,
                    ISNULL(BGHoliday, 0) AS BGHoliday,
                    ISNULL(AGHoliday, 0) AS AGHoliday,
                    ISNULL(BPHoliday, 0) AS BPHoliday,
                    ISNULL(APHoliday, 0) AS APHoliday,
                    ISNULL(DeductAbsents, 1) AS DeductAbsents,
                    ISNULL(DeductLHrs, 1) AS DeductLHrs
                FROM GeneralSettings
                WHERE DeptID = @deptId";

            var model = await db.QueryFirstOrDefaultAsync<GeneralSettingsModel>(sql, new { deptId });
            if (model == null)
            {
                model = new GeneralSettingsModel { DeptID = deptId };
            }
            return model;
        }

        public async Task<bool> SaveSettingsAsync(GeneralSettingsModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (db.State != ConnectionState.Open) db.Open();

            using var trans = db.BeginTransaction();
            try
            {
                const string checkSql = "SELECT COUNT(1) FROM GeneralSettings WHERE DeptID = @DeptID";
                int count = await db.ExecuteScalarAsync<int>(checkSql, new { model.DeptID }, trans);

                if (count > 0)
                {
                    const string updateSql = @"
                        UPDATE GeneralSettings SET
                            EnterTime = @EnterTime,
                            ExitTime = @ExitTime,
                            LunchTime = @LunchTime,
                            LunchInTime = @LunchInTime,
                            LunchOutTime = @LunchOutTime,
                            FridayBreakStartTime = @FridayBreakStartTime,
                            FridayBreakEndTime = @FridayBreakEndTime,
                            FridayCheckOut = @FridayCheckOut,
                            WorkingHrs = @WorkingHrs,
                            EmpAttCat_RefID = @EmpAttCat_RefID,
                            AbsAllowed = @AbsAllowed,
                            TotalAbsAllowed = @TotalAbsAllowed,
                            TimeAllowed = @TimeAllowed,
                            NTimes = @NTimes,
                            OverTime = @OverTime,
                            SocialAmt = @SocialAmt,
                            DinnerAmt = @DinnerAmt,
                            ShortTermPer = @ShortTermPer,
                            AdvSalPer = @AdvSalPer,
                            LongTermTime = @LongTermTime,
                            LongTermLimit = @LongTermLimit,
                            BGHoliday = @BGHoliday,
                            AGHoliday = @AGHoliday,
                            BPHoliday = @BPHoliday,
                            APHoliday = @APHoliday,
                            DeductAbsents = @DeductAbsents,
                            DeductLHrs = @DeductLHrs
                        WHERE DeptID = @DeptID";

                    await db.ExecuteAsync(updateSql, model, trans);
                }
                else
                {
                    const string insertSql = @"
                        INSERT INTO GeneralSettings (
                            DeptID, EnterTime, ExitTime, LunchTime, LunchInTime, LunchOutTime,
                            FridayBreakStartTime, FridayBreakEndTime, FridayCheckOut, WorkingHrs,
                            EmpAttCat_RefID, AbsAllowed, TotalAbsAllowed, TimeAllowed, NTimes, OverTime,
                            SocialAmt, DinnerAmt, ShortTermPer, AdvSalPer, LongTermTime, LongTermLimit,
                            BGHoliday, AGHoliday, BPHoliday, APHoliday, DeductAbsents, DeductLHrs
                        ) VALUES (
                            @DeptID, @EnterTime, @ExitTime, @LunchTime, @LunchInTime, @LunchOutTime,
                            @FridayBreakStartTime, @FridayBreakEndTime, @FridayCheckOut, @WorkingHrs,
                            @EmpAttCat_RefID, @AbsAllowed, @TotalAbsAllowed, @TimeAllowed, @NTimes, @OverTime,
                            @SocialAmt, @DinnerAmt, @ShortTermPer, @AdvSalPer, @LongTermTime, @LongTermLimit,
                            @BGHoliday, @AGHoliday, @BPHoliday, @APHoliday, @DeductAbsents, @DeductLHrs
                        )";

                    await db.ExecuteAsync(insertSql, model, trans);
                }

                // Cascade update to EmpSettings for non-exempt employees in this department
                const string updateEmpsSql = @"
                    UPDATE EmpSettings SET
                        EnterTime = @EnterTime,
                        ExitTime = @ExitTime,
                        LunchTime = @LunchTime,
                        LunchInTime = @LunchInTime,
                        LunchOutTime = @LunchOutTime,
                        FridayBreakStartTime = @FridayBreakStartTime,
                        FridayBreakEndTime = @FridayBreakEndTime,
                        FridayCheckOut = @FridayCheckOut,
                        WorkingHrs = @WorkingHrs,
                        EmpAttCat_RefID = @EmpAttCat_RefID,
                        AbsAllowed = @AbsAllowed,
                        TotalAbsAllowed = @TotalAbsAllowed,
                        TimeAllowed = @TimeAllowed,
                        NTimes = @NTimes,
                        OverTime = @OverTime,
                        SocialAmt = @SocialAmt,
                        DinnerAmt = @DinnerAmt,
                        ShortTermPer = @ShortTermPer,
                        AdvSalPer = @AdvSalPer,
                        LongTermTime = @LongTermTime,
                        LongTermLimit = @LongTermLimit,
                        BGHoliday = @BGHoliday,
                        AGHoliday = @AGHoliday,
                        BPHoliday = @BPHoliday,
                        APHoliday = @APHoliday,
                        DeductAbsents = @DeductAbsents,
                        DeductLHrs = @DeductLHrs
                    WHERE EmpID IN (
                        SELECT EmpID FROM Employees WHERE DeptID = @DeptID AND ISNULL(ExemptSettings, 0) = 0
                    )";

                await db.ExecuteAsync(updateEmpsSql, model, trans);

                trans.Commit();
                return true;
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<bool> UpdateAllDepartmentTimingsAsync(string enterTime, string exitTime, string lunchInTime, string lunchOutTime, string fridayBreakStart, string fridayBreakEnd, string fridayCheckOut)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                UPDATE GeneralSettings SET 
                    EnterTime = @enterTime,
                    ExitTime = @exitTime,
                    LunchInTime = @lunchInTime,
                    LunchOutTime = @lunchOutTime,
                    FridayBreakStartTime = @fridayBreakStart,
                    FridayBreakEndTime = @fridayBreakEnd,
                    FridayCheckOut = @fridayCheckOut";

            int rows = await db.ExecuteAsync(sql, new
            {
                enterTime,
                exitTime,
                lunchInTime,
                lunchOutTime,
                fridayBreakStart,
                fridayBreakEnd,
                fridayCheckOut
            });

            return rows > 0;
        }

        public async Task<List<EmpAttCategoryModel>> GetAttendanceCategoriesAsync()
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT 
                    EntryID,
                    LunchTimeScanning,
                    LateDeduction,
                    LateDeductionAfter,
                    GatePassRequired,
                    LeavesRecorded,
                    OverTimePaid,
                    DeductionSpecialAbsents,
                    OTRate,
                    DeductionPercentage,
                    Remarks,
                    Saturday_Early_Out_Mins
                FROM EmployeeAttendanceCatagories
                ORDER BY EntryID";

            var list = await db.QueryAsync<EmpAttCategoryModel>(sql);
            return list.ToList();
        }

        public async Task<List<CategoryRuleItem>> GetCategoryRulesByIdAsync(int categoryId)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = @"
                SELECT * FROM EmployeeAttendanceCatagories WHERE EntryID = @categoryId";

            var cat = await db.QueryFirstOrDefaultAsync<EmpAttCategoryModel>(sql, new { categoryId });
            var rules = new List<CategoryRuleItem>();

            if (cat != null)
            {
                rules.Add(new CategoryRuleItem { SettingName = "Lunch Scanning", SettingValue = (cat.LunchTimeScanning ?? false) ? "Not Required" : "Required" });
                rules.Add(new CategoryRuleItem { SettingName = "Time Deduction", SettingValue = (cat.LateDeduction ?? false) ? "No" : "Yes" });
                rules.Add(new CategoryRuleItem { SettingName = "Time Deduction After", SettingValue = cat.LateDeductionAfter ?? "0" });
                rules.Add(new CategoryRuleItem { SettingName = "Gate Pass Required", SettingValue = (cat.GatePassRequired ?? false) ? "No" : "Yes" });
                rules.Add(new CategoryRuleItem { SettingName = "Leaves Recorded", SettingValue = (cat.LeavesRecorded ?? false) ? "No" : "Yes" });
                rules.Add(new CategoryRuleItem { SettingName = "Over Time Paid", SettingValue = (cat.OverTimePaid ?? false) ? "No" : "Yes" });
                rules.Add(new CategoryRuleItem { SettingName = "Sandwich Leaves Deduction", SettingValue = (cat.DeductionSpecialAbsents ?? false) ? "Yes" : "No" });
                rules.Add(new CategoryRuleItem { SettingName = "Over Time Rate", SettingValue = (cat.OTRate ?? 0f).ToString("0.0#") });
                rules.Add(new CategoryRuleItem { SettingName = "Deduction Percentage", SettingValue = (cat.DeductionPercentage ?? 100).ToString() });
                rules.Add(new CategoryRuleItem { SettingName = "Saturday Early Out (Mins.)", SettingValue = cat.Saturday_Early_Out_Mins.ToString() });
            }

            return rules;
        }
    }
}
