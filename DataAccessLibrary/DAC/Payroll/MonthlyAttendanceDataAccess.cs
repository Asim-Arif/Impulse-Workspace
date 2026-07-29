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
    public class MonthlyAttendanceDataAccess : IMonthlyAttendanceDataAccess
    {
        private readonly string _connectionString;
        private readonly ITakeAttendanceDataAccess _takeAttDac;

        public MonthlyAttendanceDataAccess(IConfiguration config, ITakeAttendanceDataAccess takeAttDac)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
            _takeAttDac = takeAttDac;
        }

        public async Task<List<MonthlyAttendanceDayRow>> GetMonthlyAttendanceAsync(string empId, int year, int month)
        {
            var rows = new List<MonthlyAttendanceDayRow>();
            if (string.IsNullOrWhiteSpace(empId)) return rows;

            int daysInMonth = DateTime.DaysInMonth(year, month);
            DateTime startDate = new DateTime(year, month, 1);
            DateTime endDate   = new DateTime(year, month, daysInMonth);

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Fetch Holidays for the month
                string holSql = "SELECT CAST(DT AS DATE) AS DT, Description FROM Holidays WHERE CAST(DT AS DATE) BETWEEN @StartDate AND @EndDate";
                var holidays = (await db.QueryAsync<(DateTime DT, string Description)>(holSql, new { StartDate = startDate, EndDate = endDate }))
                    .ToDictionary(h => h.DT.Date, h => h.Description ?? "Holiday");

                // Fetch Attendance / Punches for the month
                string attSql = @"
                    SELECT 
                        CAST(DT AS DATE) AS DT,
                        MIN(InTime) AS InTime,
                        MAX(OutTime) AS OutTime,
                        MIN(OutTime) AS MinOutTime,
                        COUNT(*) AS AttCount,
                        SUM(CASE WHEN OverTime = 1 OR DATENAME(WeekDay, DT) = 'Sunday' THEN PayableHrs ELSE 0 END) AS OTHrs
                    FROM EmpTimes
                    WHERE EmpID = @EmpID AND CAST(DT AS DATE) BETWEEN @StartDate AND @EndDate
                    GROUP BY CAST(DT AS DATE)";

                var attDict = (await db.QueryAsync<dynamic>(attSql, new { EmpID = empId, StartDate = startDate, EndDate = endDate }))
                    .ToDictionary(a => ((DateTime)a.DT).Date, a => a);

                // Fetch Leaves for the month
                string leaveSql = "SELECT CAST(DT AS DATE) AS DT FROM Leaves WHERE EmpID = @EmpID AND CAST(DT AS DATE) BETWEEN @StartDate AND @EndDate";
                var leaveDates = (await db.QueryAsync<DateTime>(leaveSql, new { EmpID = empId, StartDate = startDate, EndDate = endDate }))
                    .Select(d => d.Date)
                    .ToHashSet();

                for (int day = 1; day <= daysInMonth; day++)
                {
                    DateTime dt = new DateTime(year, month, day);
                    bool isSunday = dt.DayOfWeek == DayOfWeek.Sunday;
                    bool isHoliday = holidays.TryGetValue(dt.Date, out string? holName);

                    string status = "Absent";
                    string inTimeStr = string.Empty;
                    string outTimeStr = string.Empty;
                    double otHrs = 0;

                    if (attDict.TryGetValue(dt.Date, out dynamic? att) && att != null)
                    {
                        status = "Present";
                        inTimeStr  = FormatTime24(att.InTime);
                        outTimeStr = FormatTime24(att.OutTime);
                        if (att.OTHrs != null)
                        {
                            double.TryParse(att.OTHrs.ToString(), out otHrs);
                        }
                    }
                    else if (leaveDates.Contains(dt.Date))
                    {
                        status = "Leave";
                    }

                    rows.Add(new MonthlyAttendanceDayRow
                    {
                        DayNo             = day,
                        Date              = dt,
                        DayName           = dt.ToString("dddd"),
                        IsSundayOrHoliday = isSunday || isHoliday,
                        HolidayName       = isHoliday ? holName ?? "Holiday" : (isSunday ? "Sunday" : string.Empty),
                        Status            = status,
                        InTime            = inTimeStr,
                        OutTime           = outTimeStr,
                        OtHours           = Math.Round(otHrs, 2),
                        IsModified        = false
                    });
                }
            }

            return rows;
        }

        public async Task<bool> SaveMonthlyAttendanceAsync(MonthlyAttendanceSaveDto input)
        {
            if (string.IsNullOrWhiteSpace(input.EmpID) || !input.Rows.Any()) return false;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();

                // Fetch Employee Shift Break Times
                string breakSql = @"
                    SELECT TOP 1
                        ISNULL(LunchInTime, '13:00') AS LunchInTime,
                        ISNULL(LunchOutTime, '14:00') AS LunchOutTime,
                        ISNULL(FridayBreakStartTime, '12:30') AS FridayBreakStartTime,
                        ISNULL(FridayBreakEndTime, '14:00') AS FridayBreakEndTime
                    FROM VEmpSettings
                    WHERE EmpID = @EmpID";

                var shift = await db.QueryFirstOrDefaultAsync<dynamic>(breakSql, new { EmpID = input.EmpID });
                string breakStartNormal = shift?.LunchInTime ?? "13:00";
                string breakEndNormal   = shift?.LunchOutTime ?? "14:00";
                string breakStartFri    = shift?.FridayBreakStartTime ?? "12:30";
                string breakEndFri      = shift?.FridayBreakEndTime ?? "14:00";

                foreach (var row in input.Rows)
                {
                    if (!string.IsNullOrWhiteSpace(row.InTime))
                    {
                        DateTime myDT = row.Date.Date;

                        // Delete existing Attendance & EmpTimes for this date
                        using (IDbTransaction trans = db.BeginTransaction())
                        {
                            try
                            {
                                string delCond = "WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                                await db.ExecuteAsync($"DELETE FROM AttendanceSheet {delCond}", new { EmpID = input.EmpID, DT = myDT }, trans);
                                await db.ExecuteAsync($"DELETE FROM EmpTimes {delCond}", new { EmpID = input.EmpID, DT = myDT }, trans);
                                trans.Commit();
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }

                        // Process In Time punch
                        if (!string.IsNullOrWhiteSpace(row.InTime))
                        {
                            await _takeAttDac.ProcessPunchAsync(input.EmpID, myDT, row.InTime);

                            // Process Out Time punch if present
                            if (!string.IsNullOrWhiteSpace(row.OutTime))
                            {
                                TimeSpan.TryParse(row.OutTime, out TimeSpan outTs);

                                // Determine breaks
                                string restStart = (myDT.DayOfWeek == DayOfWeek.Friday) ? breakStartFri : breakStartNormal;
                                string restEnd   = (myDT.DayOfWeek == DayOfWeek.Friday) ? breakEndFri : breakEndNormal;

                                TimeSpan.TryParse(restStart, out TimeSpan bStartTs);
                                TimeSpan.TryParse(restEnd,   out TimeSpan bEndTs);

                                TimeSpan cutOffNight = new TimeSpan(4, 0, 0); // 04:00 AM

                                // Insert break punches if OutTime extends past break start
                                if (outTs <= cutOffNight || outTs > bStartTs)
                                {
                                    await _takeAttDac.ProcessPunchAsync(input.EmpID, myDT, restStart);
                                    await _takeAttDac.ProcessPunchAsync(input.EmpID, myDT, restEnd);
                                }

                                // Insert OutTime punch (handling next-day rollover if <= 04:00 AM)
                                if (outTs <= cutOffNight)
                                {
                                    await _takeAttDac.ProcessPunchAsync(input.EmpID, myDT.AddDays(1), row.OutTime);
                                }
                                else
                                {
                                    await _takeAttDac.ProcessPunchAsync(input.EmpID, myDT, row.OutTime);
                                }
                            }
                        }
                    }
                }

                return true;
            }
        }

        public async Task<bool> ClearDateAttendanceAsync(string empId, DateTime date)
        {
            if (string.IsNullOrWhiteSpace(empId)) return false;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                string sql = @"
                    DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT;
                    DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT;";

                await db.ExecuteAsync(sql, new { EmpID = empId, DT = date.Date });
                return true;
            }
        }

        private static string FormatTime24(object? val)
        {
            if (val == null || DBNull.Value == val) return string.Empty;
            string s = val.ToString() ?? string.Empty;
            if (DateTime.TryParse(s, out DateTime dt))
                return dt.ToString("HH:mm");
            if (TimeSpan.TryParse(s, out TimeSpan ts))
                return ts.ToString(@"hh\:mm");
            return s;
        }
    }
}
