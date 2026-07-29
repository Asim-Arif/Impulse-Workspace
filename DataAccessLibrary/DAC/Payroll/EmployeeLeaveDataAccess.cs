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
    public class EmployeeLeaveDataAccess : IEmployeeLeaveDataAccess
    {
        private readonly string _connectionString;

        public EmployeeLeaveDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<LeaveBalanceMatrixDto> GetLeaveBalancesAsync(string empId, DateTime date)
        {
            var matrix = new LeaveBalanceMatrixDto();
            if (string.IsNullOrWhiteSpace(empId)) return matrix;

            // Set default Available limits (matching legacy Form_Load)
            matrix.CL.Available  = 10;
            matrix.SL.Available  = 8;
            matrix.AL.Available  = 14;
            matrix.SAL.Available = 0;
            matrix.LWP.Available = 0;
            matrix.HL.Available  = 12;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new DynamicParameters();
                parameters.Add("@EmpID", empId, DbType.String, size: 50);
                parameters.Add("@DT",    date,  DbType.DateTime);

                var row = await db.QueryFirstOrDefaultAsync<dynamic>(
                    "GetLeaveBalances",
                    parameters,
                    commandType: CommandType.StoredProcedure);

                if (row != null)
                {
                    var dict = (IDictionary<string, object>)row;

                    double GetVal(string key)
                    {
                        if (dict.TryGetValue(key, out object? val) && val != null && DBNull.Value != val)
                        {
                            if (double.TryParse(val.ToString(), out double d))
                                return d;
                        }
                        return 0;
                    }

                    // Used (prior to this month)
                    matrix.CL.Used  = GetVal("CL_E");
                    matrix.SL.Used  = GetVal("SL_E");
                    matrix.AL.Used  = GetVal("AL_E");
                    matrix.SAL.Used = GetVal("SAL_E");
                    matrix.LWP.Used = GetVal("LWP_E");
                    matrix.HL.Used  = GetVal("HL_E");

                    // This Month
                    matrix.CL.ThisMonth  = GetVal("CL");
                    matrix.SL.ThisMonth  = GetVal("SL");
                    matrix.AL.ThisMonth  = GetVal("AL");
                    matrix.SAL.ThisMonth = GetVal("SAL");
                    matrix.LWP.ThisMonth = GetVal("LWP");
                    matrix.HL.ThisMonth  = GetVal("HL");
                }
            }

            return matrix;
        }

        public async Task<List<UserLookupModel>> GetUsersAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT UserName FROM Users WHERE ISNULL(InActive, 0) = 0 ORDER BY UserName";
                var users = await db.QueryAsync<UserLookupModel>(sql);
                return users.AsList();
            }
        }

        public async Task<bool> IsProbationAsync(string empId, DateTime date)
        {
            if (string.IsNullOrWhiteSpace(empId)) return false;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT CASE WHEN JoinDate IS NOT NULL AND DATEADD(month, 3, JoinDate) > @DT THEN 1 ELSE 0 END
                    FROM Employees
                    WHERE EmpID = @EmpID";

                int result = await db.ExecuteScalarAsync<int>(sql, new { EmpID = empId, DT = date });
                return result == 1;
            }
        }

        public async Task<bool> HasExistingAttendanceAsync(string empId, DateTime fromDate, DateTime toDate)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT COUNT(*) FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) BETWEEN @FromDate AND @ToDate";
                int count = await db.ExecuteScalarAsync<int>(sql, new { EmpID = empId, FromDate = fromDate.Date, ToDate = toDate.Date });
                return count > 0;
            }
        }

        public async Task<bool> SaveEmployeeLeaveAsync(EmployeeLeaveInputDto input)
        {
            if (string.IsNullOrWhiteSpace(input.EmpID)) return false;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (IDbTransaction trans = db.BeginTransaction())
                {
                    try
                    {
                        DateTime fromDate = input.FromDate.Date;
                        DateTime toDate   = input.ToDate.Date;

                        // Delete existing attendance & leaves if full day
                        if (input.DurationMode == "FullDay")
                        {
                            string delCond = "WHERE EmpID = @EmpID AND CAST(DT AS DATE) BETWEEN @FromDate AND @ToDate";
                            await db.ExecuteAsync($"DELETE FROM AttendanceSheet {delCond}", new { EmpID = input.EmpID, FromDate = fromDate, ToDate = toDate }, trans);
                            await db.ExecuteAsync($"DELETE FROM EmpTimes {delCond}", new { EmpID = input.EmpID, FromDate = fromDate, ToDate = toDate }, trans);
                            await db.ExecuteAsync($"DELETE FROM Leaves {delCond}", new { EmpID = input.EmpID, FromDate = fromDate, ToDate = toDate }, trans);
                        }

                        // Determine duration hours & leave fraction
                        double dHrs = 8.0;
                        if (input.DurationMode == "HalfDay")
                        {
                            dHrs = 4.0;
                        }
                        else if (input.DurationMode == "ShortLeave")
                        {
                            if (DateTime.TryParse(input.FromTime, out DateTime tFrom) && DateTime.TryParse(input.ToTime, out DateTime tTo))
                            {
                                double mins = (tTo - tFrom).TotalMinutes;
                                if (mins > 0) dHrs = Math.Round(mins / 60.0, 1);
                            }
                        }

                        double dLeave = dHrs / 8.0;
                        int fullFlag  = (input.DurationMode == "FullDay") ? 1 : 0;
                        int deductFlag = input.WithDeduction ? 1 : 0;
                        int attCode   = input.IsAbsent ? 0 : (input.WithDeduction ? 3 : 2);

                        // Validate allowed column names for leaves
                        string validCode = "CL";
                        string[] allowedCols = { "CL", "SL", "AL", "CPL", "LWP", "ML", "SAL", "TML" };
                        if (allowedCols.Contains(input.LeaveTypeCode, StringComparer.OrdinalIgnoreCase))
                        {
                            validCode = input.LeaveTypeCode.ToUpper();
                        }

                        int daySpan = (toDate - fromDate).Days;
                        for (int i = 0; i <= daySpan; i++)
                        {
                            DateTime dt = fromDate.AddDays(i);

                            // Insert into AttendanceSheet
                            string attSql = @"
                                INSERT INTO AttendanceSheet(EmpID, DT, Attendance)
                                VALUES(@EmpID, @DT, @Attendance)";
                            await db.ExecuteAsync(attSql, new { EmpID = input.EmpID, DT = dt, Attendance = attCode }, trans);

                            if (!input.IsAbsent)
                            {
                                string leaveSql = $@"
                                    INSERT INTO Leaves(EmpID, DT, AuthPerson, Reason, [From], [To], [Full], Hrs, Deduct, {validCode})
                                    VALUES(@EmpID, @DT, @AuthPerson, @Reason, @FromTime, @ToTime, @FullFlag, @Hrs, @DeductFlag, @LeaveVal)";

                                await db.ExecuteAsync(leaveSql, new
                                {
                                    EmpID      = input.EmpID,
                                    DT         = dt,
                                    AuthPerson = input.AuthPerson,
                                    Reason     = input.Reason,
                                    FromTime   = input.FromTime,
                                    ToTime     = input.ToTime,
                                    FullFlag   = fullFlag,
                                    Hrs        = dHrs,
                                    DeductFlag = deductFlag,
                                    LeaveVal   = dLeave
                                }, trans);
                            }
                        }

                        trans.Commit();
                        return true;
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }

        public async Task ExecutePrintEmpLeaveBalancesAsync(DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                var parameters = new DynamicParameters();
                DateTime monthStart = new DateTime(date.Year, date.Month, 1);
                parameters.Add("@DT", monthStart, DbType.DateTime);

                await db.ExecuteAsync(
                    "SP_EmpLeaveBalances",
                    parameters,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: 120);
            }
        }
    }
}
