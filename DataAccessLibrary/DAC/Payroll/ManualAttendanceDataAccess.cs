using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class ManualAttendanceDataAccess : IManualAttendanceDataAccess
    {
        private readonly string _connectionString;
        private readonly ITakeAttendanceDataAccess _takeAttendanceDataAccess;

        public ManualAttendanceDataAccess(IConfiguration config, ITakeAttendanceDataAccess takeAttendanceDataAccess)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
            _takeAttendanceDataAccess = takeAttendanceDataAccess;
        }

        public async Task<List<EmployeeListItemModel>> GetActiveEmployeesAsync()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EmpID, Name, ISNULL(FName, '') AS FName, ISNULL(Designation, '') AS Designation,
                           ISNULL(StartingSalary, 0) AS StartingSalary, ISNULL(Active, 1) AS Active, DeptID
                    FROM Employees
                    WHERE ISNULL(Active, 1) = 1
                    ORDER BY Name";

                return (await db.QueryAsync<EmployeeListItemModel>(sql)).AsList();
            }
        }

        public async Task<byte[]?> GetEmployeePictureAsync(string empId)
        {
            if (string.IsNullOrWhiteSpace(empId)) return null;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = "SELECT EmpPic FROM Employees WHERE EmpID = @EmpID";
                return await db.ExecuteScalarAsync<byte[]?>(sql, new { EmpID = empId });
            }
        }

        public async Task<EmpShiftDefaultTimings> GetEmployeeDefaultTimingsAsync(string empId, DateTime date)
        {
            var timings = new EmpShiftDefaultTimings();
            if (string.IsNullOrWhiteSpace(empId)) return timings;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT g.EnterTime, g.ExitTime, g.LunchInTime, g.LunchOutTime,
                           g.FridayBreakStartTime, g.FridayBreakEndTime, g.FridayCheckOut
                    FROM Employees e
                    INNER JOIN GeneralSettings g ON e.DeptID = g.DeptID
                    WHERE e.EmpID = @EmpID";

                var row = await db.QueryFirstOrDefaultAsync(sql, new { EmpID = empId });
                if (row != null)
                {
                    bool isFriday = date.DayOfWeek == DayOfWeek.Friday;

                    string enter = row.EnterTime != null ? row.EnterTime.ToString().Trim() : "08:00";
                    string exit = isFriday && row.FridayCheckOut != null ? row.FridayCheckOut.ToString().Trim() : (row.ExitTime != null ? row.ExitTime.ToString().Trim() : "17:00");
                    string lunchIn = isFriday && row.FridayBreakStartTime != null ? row.FridayBreakStartTime.ToString().Trim() : (row.LunchInTime != null ? row.LunchInTime.ToString().Trim() : "13:00");
                    string lunchOut = isFriday && row.FridayBreakEndTime != null ? row.FridayBreakEndTime.ToString().Trim() : (row.LunchOutTime != null ? row.LunchOutTime.ToString().Trim() : "14:00");

                    timings.EnterTime = FormatTimeStringTo24Hr(enter, "08:00");
                    timings.ExitTime = FormatTimeStringTo24Hr(exit, "17:00");
                    timings.LunchInTime = FormatTimeStringTo24Hr(lunchIn, "13:00");
                    timings.LunchOutTime = FormatTimeStringTo24Hr(lunchOut, "14:00");
                }
            }

            return timings;
        }

        public async Task<List<EmpTimeDetailRow>> GetEmployeeAttendanceDetailsAsync(string empId, DateTime date)
        {
            if (string.IsNullOrWhiteSpace(empId)) return new List<EmpTimeDetailRow>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    SELECT EntryID, EmpID, DT, ISNULL(InTime, '') AS InTime, ISNULL(OutTime, '') AS OutTime,
                           ISNULL(Hrs, 0) AS Hrs, ISNULL(PayableHrs, 0) AS PayableHrs,
                           ISNULL(OverTime, 0) AS OverTime, ISNULL(UserName, '') AS UserName
                    FROM EmpTimes
                    WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT
                    ORDER BY EntryID";

                var rows = await db.QueryAsync<EmpTimeDetailRow>(sql, new { EmpID = empId, DT = date.Date });
                return rows.AsList();
            }
        }

        public async Task<string?> CheckLeaveOrHolidayAsync(string empId, DateTime date)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Check Leaves
                string leaveSql = "SELECT TOP 1 ISNULL(NULLIF(Reason, ''), 'Leave Recorded') FROM Leaves WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                string? leave = await db.ExecuteScalarAsync<string?>(leaveSql, new { EmpID = empId, DT = date.Date });
                if (!string.IsNullOrEmpty(leave))
                {
                    return $"Employee is on leave ({leave}) on {date:dd-MMM-yyyy}.";
                }

                // Check Holidays
                string holidaySql = "SELECT TOP 1 Description FROM Holidays WHERE CAST(DT AS DATE) = @DT";
                string? holiday = await db.ExecuteScalarAsync<string?>(holidaySql, new { DT = date.Date });
                if (!string.IsNullOrEmpty(holiday))
                {
                    return $"Official Holiday ({holiday}) on {date:dd-MMM-yyyy}.";
                }

                return null;
            }
        }

        public async Task<bool> SaveManualAttendanceAsync(ManualAttendanceInputDto input)
        {
            if (input == null || string.IsNullOrWhiteSpace(input.EmpID)) return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        DateTime attDate = input.DT.Date;

                        // 1. Delete previous records in transaction
                        string delAttSql = "DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                        await db.ExecuteAsync(delAttSql, new { EmpID = input.EmpID, DT = attDate }, trans);

                        string delTimesSql = "DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                        await db.ExecuteAsync(delTimesSql, new { EmpID = input.EmpID, DT = attDate }, trans);

                        // 2. Calculate Late Hours
                        double lateHours = 0;
                        DateTime inTimeDt = ParseTimeString(input.InTime, attDate);
                        DateTime stdStart = attDate.AddHours(8); // 08:00 AM
                        if (inTimeDt > stdStart.AddMinutes(10))
                        {
                            lateHours = Math.Round((inTimeDt - stdStart).TotalHours, 2);
                        }

                        // 3. Insert AttendanceSheet
                        string insAttSql = @"
                            INSERT INTO AttendanceSheet (EmpID, DT, Attendance, LateHours)
                            VALUES (@EmpID, @DT, 1, @LateHours)";
                        await db.ExecuteAsync(insAttSql, new { EmpID = input.EmpID, DT = attDate, LateHours = lateHours }, trans);

                        // 4. Insert Shift Slot 1 (Morning: InTime to BrkOut)
                        if (!string.IsNullOrWhiteSpace(input.InTime) && !string.IsNullOrWhiteSpace(input.BrkOut))
                        {
                            DateTime tIn = ParseTimeString(input.InTime, attDate);
                            DateTime tOut = ParseTimeString(input.BrkOut, attDate);
                            if (tOut > tIn)
                            {
                                double hrs = Math.Round((tOut - tIn).TotalHours, 2);
                                await InsertEmpTimeRowAsync(db, trans, input.EmpID, attDate, Format24Hr(tIn), Format24Hr(tOut), hrs, hrs, false, input.UserName, input.ComputerName);
                            }
                        }

                        // 5. Insert Shift Slot 2 (Afternoon: BrkIn to ChkOut)
                        if (!string.IsNullOrWhiteSpace(input.BrkIn) && !string.IsNullOrWhiteSpace(input.ChkOut))
                        {
                            DateTime tIn = ParseTimeString(input.BrkIn, attDate);
                            DateTime tOut = ParseTimeString(input.ChkOut, attDate);
                            if (tOut > tIn)
                            {
                                double hrs = Math.Round((tOut - tIn).TotalHours, 2);
                                await InsertEmpTimeRowAsync(db, trans, input.EmpID, attDate, Format24Hr(tIn), Format24Hr(tOut), hrs, hrs, false, input.UserName, input.ComputerName);
                            }
                        }

                        // 6. Insert Shift Slot 3 (OverTime: OTIn to OTOut)
                        if (input.EnableOT && !string.IsNullOrWhiteSpace(input.OTIn) && !string.IsNullOrWhiteSpace(input.OTOut))
                        {
                            DateTime tIn = ParseTimeString(input.OTIn, attDate);
                            DateTime tOut = ParseTimeString(input.OTOut, attDate);
                            if (tOut > tIn)
                            {
                                double hrs = Math.Round((tOut - tIn).TotalHours, 2);
                                await InsertEmpTimeRowAsync(db, trans, input.EmpID, attDate, Format24Hr(tIn), Format24Hr(tOut), hrs, hrs, true, input.UserName, input.ComputerName);
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

        public async Task<bool> DeleteAttendanceAsync(string empId, DateTime date)
        {
            if (string.IsNullOrWhiteSpace(empId)) return false;

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                using (var trans = db.BeginTransaction())
                {
                    try
                    {
                        DateTime attDate = date.Date;
                        string delAttSql = "DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                        await db.ExecuteAsync(delAttSql, new { EmpID = empId, DT = attDate }, trans);

                        string delTimesSql = "DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                        await db.ExecuteAsync(delTimesSql, new { EmpID = empId, DT = attDate }, trans);

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

        private async Task InsertEmpTimeRowAsync(IDbConnection db, IDbTransaction trans, string empId, DateTime date, string inTimeStr, string outTimeStr, double hrs, double payableHrs, bool isOverTime, string userName, string computerName)
        {
            string sql = @"
                INSERT INTO EmpTimes (EmpID, DT, InTime, OutTime, Hrs, PayableHrs, ManualEntry, UserName, ComputerName, EntryDT, OverTime)
                VALUES (@EmpID, @DT, @InTime, @OutTime, @Hrs, @PayableHrs, 1, @UserName, @ComputerName, GETDATE(), @OverTime)";

            await db.ExecuteAsync(sql, new
            {
                EmpID = empId,
                DT = date,
                InTime = inTimeStr,
                OutTime = outTimeStr,
                Hrs = hrs,
                PayableHrs = payableHrs,
                UserName = userName,
                ComputerName = string.IsNullOrWhiteSpace(computerName) ? "Unknown" : computerName,
                OverTime = isOverTime ? 1 : 0
            }, trans);
        }

        private string FormatTimeStringTo24Hr(string timeStr, string defaultVal)
        {
            if (string.IsNullOrWhiteSpace(timeStr)) return defaultVal;
            if (DateTime.TryParse(timeStr, out DateTime dt))
            {
                return dt.ToString("HH:mm", CultureInfo.InvariantCulture);
            }
            if (TimeSpan.TryParse(timeStr, out TimeSpan ts))
            {
                return DateTime.Today.Add(ts).ToString("HH:mm", CultureInfo.InvariantCulture);
            }
            return defaultVal;
        }

        private DateTime ParseTimeString(string timeStr, DateTime baseDate)
        {
            if (DateTime.TryParse(timeStr, out DateTime parsed))
            {
                return baseDate.Date.Add(parsed.TimeOfDay);
            }
            if (TimeSpan.TryParse(timeStr, out TimeSpan ts))
            {
                return baseDate.Date.Add(ts);
            }
            return baseDate.Date;
        }

        private string Format24Hr(DateTime dt)
        {
            return dt.ToString("HH:mm", CultureInfo.InvariantCulture);
        }

        public Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId)
        {
            return _takeAttendanceDataAccess.ReCalculateAttendanceAsync(dateFrom, dateTo, deptId, empId);
        }
    }
}
