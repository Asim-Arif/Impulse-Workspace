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
    public interface IPayrollPoliciesDataAccess
    {
        Task<List<EmpAttCategoryModel>> GetCategoriesAsync();
        Task<EmpAttCategoryModel?> GetCategoryByIdAsync(int id);
        Task<bool> SaveCategoryAsync(EmpAttCategoryModel model);
        Task<bool> DeleteCategoryAsync(int id);
    }

    public class PayrollPoliciesDataAccess : IPayrollPoliciesDataAccess
    {
        private readonly string _connectionString;

        public PayrollPoliciesDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<EmpAttCategoryModel>> GetCategoriesAsync()
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
                    ISNULL(Saturday_Early_Out_Mins, 0) AS Saturday_Early_Out_Mins
                FROM EmployeeAttendanceCatagories
                ORDER BY EntryID";

            var list = await db.QueryAsync<EmpAttCategoryModel>(sql);
            return list.ToList();
        }

        public async Task<EmpAttCategoryModel?> GetCategoryByIdAsync(int id)
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
                    ISNULL(Saturday_Early_Out_Mins, 0) AS Saturday_Early_Out_Mins
                FROM EmployeeAttendanceCatagories
                WHERE EntryID = @id";

            return await db.QueryFirstOrDefaultAsync<EmpAttCategoryModel>(sql, new { id });
        }

        public async Task<bool> SaveCategoryAsync(EmpAttCategoryModel model)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            if (model.EntryID == 0)
            {
                const string insertSql = @"
                    INSERT INTO EmployeeAttendanceCatagories (
                        LunchTimeScanning, LateDeduction, LateDeductionAfter, GatePassRequired,
                        LeavesRecorded, OverTimePaid, DeductionSpecialAbsents, OTRate,
                        DeductionPercentage, Remarks, Saturday_Early_Out_Mins
                    ) VALUES (
                        @LunchTimeScanning, @LateDeduction, @LateDeductionAfter, @GatePassRequired,
                        @LeavesRecorded, @OverTimePaid, @DeductionSpecialAbsents, @OTRate,
                        @DeductionPercentage, @Remarks, @Saturday_Early_Out_Mins
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                int newId = await db.ExecuteScalarAsync<int>(insertSql, model);
                model.EntryID = newId;
                return newId > 0;
            }
            else
            {
                const string updateSql = @"
                    UPDATE EmployeeAttendanceCatagories SET
                        LunchTimeScanning = @LunchTimeScanning,
                        LateDeduction = @LateDeduction,
                        LateDeductionAfter = @LateDeductionAfter,
                        GatePassRequired = @GatePassRequired,
                        LeavesRecorded = @LeavesRecorded,
                        OverTimePaid = @OverTimePaid,
                        DeductionSpecialAbsents = @DeductionSpecialAbsents,
                        OTRate = @OTRate,
                        DeductionPercentage = @DeductionPercentage,
                        Remarks = @Remarks,
                        Saturday_Early_Out_Mins = @Saturday_Early_Out_Mins
                    WHERE EntryID = @EntryID";

                int rows = await db.ExecuteAsync(updateSql, model);
                return rows > 0;
            }
        }

        public async Task<bool> DeleteCategoryAsync(int id)
        {
            using IDbConnection db = new SqlConnection(_connectionString);
            const string sql = "DELETE FROM EmployeeAttendanceCatagories WHERE EntryID = @id";
            int rows = await db.ExecuteAsync(sql, new { id });
            return rows > 0;
        }
    }
}
