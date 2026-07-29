using Dapper;
using DataAccessLibrary.Interface.Payroll;
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
    public class TakeAttendanceDataAccess : ITakeAttendanceDataAccess
    {
        private readonly string _connectionString;

        public TakeAttendanceDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<bool> ProcessPunchAsync(string empId, DateTime date, string punchTimeStr, string userName = "System", string computerName = "Unknown", IDbConnection? db = null, IDbTransaction? trans = null)
        {
            if (string.IsNullOrWhiteSpace(empId)) return false;

            bool isLocalConn = false;
            if (db == null)
            {
                db = new SqlConnection(_connectionString);
                await ((SqlConnection)db).OpenAsync();
                isLocalConn = true;
            }

            try
            {
                // 1. Fetch Employee Details
                string empInfoSql = "SELECT DeptID, ISNULL(ExemptSettings, 0) AS ExemptSettings FROM Employees WHERE EmpID = @EmpID";
                var empInfo = await db.QueryFirstOrDefaultAsync(empInfoSql, new { EmpID = empId }, trans);
                if (empInfo == null) return false;

                string deptId = (string)empInfo.DeptID ?? "";
                bool bExempt = empInfo.ExemptSettings != null && (bool)empInfo.ExemptSettings;

                // 2. Fetch Shift Timings & Category Settings
                string enterTime = "08:00";
                bool deductLHrs = true;
                int attCat = 0;

                string shiftSql;
                if (bExempt)
                {
                    shiftSql = @"SELECT EmpAttCat_RefID, EnterTime, ExitTime, LunchInTime, LunchOutTime, ISNULL(DeductLHrs, 1) AS DeductLHrs,
                                        FridayBreakStartTime, FridayBreakEndTime, FridayCheckOut
                                 FROM EmpSettings WHERE EmpID = @EmpID";
                }
                else
                {
                    shiftSql = @"SELECT EmpAttCat_RefID, EnterTime, ExitTime, LunchInTime, LunchOutTime, ISNULL(DeductLHrs, 1) AS DeductLHrs,
                                        FridayBreakStartTime, FridayBreakEndTime, FridayCheckOut
                                 FROM GeneralSettings WHERE DeptID = @DeptID";
                }

                var shiftRes = await db.QueryFirstOrDefaultAsync(shiftSql, new { EmpID = empId, DeptID = deptId }, trans);
                if (shiftRes != null)
                {
                    attCat = shiftRes.EmpAttCat_RefID != null ? Convert.ToInt32(shiftRes.EmpAttCat_RefID) : 0;
                    if (shiftRes.DeductLHrs != null) deductLHrs = (bool)shiftRes.DeductLHrs;

                    bool isFriday = date.DayOfWeek == DayOfWeek.Friday;
                    if (isFriday && shiftRes.FridayCheckOut != null)
                    {
                        enterTime = shiftRes.EnterTime != null ? FormatTimeStringTo24Hr(Convert.ToString(shiftRes.EnterTime), "08:00") : "08:00";
                    }
                    else if (shiftRes.EnterTime != null)
                    {
                        enterTime = FormatTimeStringTo24Hr(Convert.ToString(shiftRes.EnterTime), "08:00");
                    }
                }

                // 3. Category Settings (LateDeduction, LateDeductionAfter)
                int lateDeductionAfter = 0;
                if (attCat > 0)
                {
                    string catSql = "SELECT ISNULL(LateDeductionAfter, 0) AS LateDeductionAfter FROM EmployeeAttendanceCatagories WHERE EntryID = @EntryID";
                    lateDeductionAfter = await db.ExecuteScalarAsync<int>(catSql, new { EntryID = attCat }, trans);
                }

                DateTime lateDeductionCutoff = ParseTimeString(enterTime, date).AddMinutes(lateDeductionAfter);

                // 3. Query existing entries for target date
                string existingSql = "SELECT EntryID, InTime, OutTime FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT ORDER BY EntryID";
                var existingEntries = (await db.QueryAsync(existingSql, new { EmpID = empId, DT = date.Date }, trans)).AsList();

                if (existingEntries.Count == 0)
                {
                    DateTime pTime = ParseTimeString(punchTimeStr, date);
                    double lateHours = 0;
                    if (deductLHrs && pTime > lateDeductionCutoff)
                    {
                        DateTime eTime = ParseTimeString(enterTime, date);
                        lateHours = Math.Round((pTime - eTime).TotalHours, 2);
                    }

                    await db.ExecuteAsync("DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = empId, DT = date.Date }, trans);
                    await db.ExecuteAsync("INSERT INTO AttendanceSheet (EmpID, DT, Attendance, LateHours, ExtraHrsToDeduct) VALUES (@EmpID, @DT, 1, @LateHours, 0)",
                        new { EmpID = empId, DT = date.Date, LateHours = lateHours }, trans);

                    await InsertEmpTimeRowAsync(db, trans, empId, date.Date, punchTimeStr, "", 0, 0, false, userName, computerName);
                }
                else
                {
                    var openEntry = existingEntries.FirstOrDefault(e => string.IsNullOrWhiteSpace((string)e.OutTime));
                    if (openEntry != null)
                    {
                        string inTimeStr = (string)openEntry.InTime;
                        DateTime tIn = ParseTimeString(inTimeStr, date);
                        DateTime tOut = ParseTimeString(punchTimeStr, date);
                        double hrs = tOut > tIn ? Math.Round((tOut - tIn).TotalHours, 2) : 0;

                        string updateSql = "UPDATE EmpTimes SET OutTime = @OutTime, Hrs = @Hrs, PayableHrs = @Hrs WHERE EntryID = @EntryID";
                        await db.ExecuteAsync(updateSql, new { OutTime = punchTimeStr, Hrs = hrs, EntryID = (int)openEntry.EntryID }, trans);
                    }
                    else
                    {
                        await InsertEmpTimeRowAsync(db, trans, empId, date.Date, punchTimeStr, "", 0, 0, false, userName, computerName);
                    }
                }

                return true;
            }
            finally
            {
                if (isLocalConn && db != null)
                {
                    db.Dispose();
                }
            }
        }

        public async Task<int> ReCalculateAttendanceAsync(DateTime dateFrom, DateTime dateTo, string deptId, string empId)
        {
            DateTime dtFrom = dateFrom.Date;
            DateTime dtTo = dateTo.Date;

            string sql = @"
                SELECT EmpID, DT, DeptID,
                       FirstInTime, FirstOutTime,
                       SecondInTime, SecondOutTime,
                       LunchInTime, LunchOutTime,
                       OTInTime, OTOutTime, OTHrs
                FROM VEmpTimes2
                WHERE CAST(DT AS DATE) BETWEEN @DateFrom AND @DateTo";

            if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
            {
                sql += " AND DeptID = @DeptID";
            }
            if (!string.IsNullOrWhiteSpace(empId) && empId != "0")
            {
                sql += " AND EmpID = @EmpID";
            }
            sql += " ORDER BY EmpID, DT";

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                var records = (await db.QueryAsync(sql, new { DateFrom = dtFrom, DateTo = dtTo, DeptID = deptId, EmpID = empId })).AsList();

                int count = 0;
                foreach (var r in records)
                {
                    string curEmpId = Convert.ToString(r.EmpID) ?? string.Empty;
                    DateTime myDT = Convert.ToDateTime(r.DT).Date;

                    string firstIn = r.FirstInTime != null ? Convert.ToString(r.FirstInTime) : "";
                    string firstOut = r.FirstOutTime != null ? Convert.ToString(r.FirstOutTime) : "";
                    string secondIn = r.SecondInTime != null ? Convert.ToString(r.SecondInTime) : "";
                    string secondOut = r.SecondOutTime != null ? Convert.ToString(r.SecondOutTime) : "";
                    string otIn = r.OTInTime != null ? Convert.ToString(r.OTInTime) : "";
                    string otOut = r.OTOutTime != null ? Convert.ToString(r.OTOutTime) : "";
                    string lunchIn = r.LunchInTime != null ? Convert.ToString(r.LunchInTime) : "";
                    string lunchOut = r.LunchOutTime != null ? Convert.ToString(r.LunchOutTime) : "";
                    double dOTHrs = r.OTHrs != null ? Convert.ToDouble(r.OTHrs) : 0;

                    if (!string.IsNullOrEmpty(lunchIn) && lunchIn == lunchOut)
                    {
                        if (!string.IsNullOrWhiteSpace(secondOut))
                        {
                            firstOut = secondOut;
                        }
                    }

                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            string delTimesSql = "DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT";
                            await db.ExecuteAsync(delTimesSql, new { EmpID = curEmpId, DT = myDT }, trans);

                            if (!string.IsNullOrWhiteSpace(firstIn))
                                await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(firstIn, "08:00"), "System", "127.0.0.1", db, trans);

                            if (!string.IsNullOrWhiteSpace(firstOut))
                                await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(firstOut, "13:00"), "System", "127.0.0.1", db, trans);

                            if (!string.IsNullOrWhiteSpace(secondIn))
                                await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(secondIn, "14:00"), "System", "127.0.0.1", db, trans);

                            if (!string.IsNullOrWhiteSpace(otOut))
                            {
                                if (TimeOnly.TryParse(otOut, out TimeOnly otOutTime) && otOutTime < new TimeOnly(7, 0) && dOTHrs > 3)
                                {
                                    await ProcessPunchAsync(curEmpId, myDT.AddDays(1), FormatTimeStringTo24Hr(otOut, "20:00"), "System", "127.0.0.1", db, trans);
                                }
                                else
                                {
                                    await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(otOut, "20:00"), "System", "127.0.0.1", db, trans);
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrWhiteSpace(secondOut))
                                    await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(secondOut, "17:00"), "System", "127.0.0.1", db, trans);

                                if (!string.IsNullOrWhiteSpace(otIn))
                                    await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(otIn, "17:30"), "System", "127.0.0.1", db, trans);

                                if (!string.IsNullOrWhiteSpace(otOut))
                                    await ProcessPunchAsync(curEmpId, myDT, FormatTimeStringTo24Hr(otOut, "20:00"), "System", "127.0.0.1", db, trans);
                            }

                            trans.Commit();
                            count++;
                        }
                        catch
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }

                return count;
            }
        }

        private async Task InsertEmpTimeRowAsync(IDbConnection db, IDbTransaction? trans, string empId, DateTime date, string inTimeStr, string outTimeStr, double hrs, double payableHrs, bool isOverTime, string userName, string computerName)
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
        public async Task<int> AutoGenerateAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo)
        {
            DateTime dtFrom = dateFrom.Date;
            DateTime dtTo = dateTo.Date;

            string empCond = " WHERE ISNULL(Employees.Active, 1) = 1";
            if (!string.IsNullOrWhiteSpace(empId) && empId != "0")
            {
                empCond += " AND Employees.EmpID = @EmpID";
            }
            else if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
            {
                empCond += " AND Employees.DeptID = @DeptID";
            }

            string sql = $@"
                SELECT Employees.EmpID,
                       ISNULL(VEmpSettings.LunchInTime, '13:00') AS LunchInTime,
                       ISNULL(VEmpSettings.LunchOutTime, '14:00') AS LunchOutTime,
                       ISNULL(VEmpSettings.FridayBreakStartTime, '12:30') AS FridayBreakStartTime,
                       ISNULL(VEmpSettings.FridayBreakEndTime, '14:00') AS FridayBreakEndTime,
                       ISNULL(VEmpSettings.EnterTime, '08:00') AS EnterTime,
                       ISNULL(VEmpSettings.ExitTime, '17:00') AS ExitTime
                FROM Employees
                LEFT JOIN VEmpSettings ON Employees.EmpID = VEmpSettings.EmpID
                {empCond}";

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                var empList = (await db.QueryAsync(sql, new { DeptID = deptId, EmpID = empId })).AsList();

                int generatedCount = 0;
                Random rand = new Random();

                foreach (var emp in empList)
                {
                    string curEmpId = Convert.ToString(emp.EmpID) ?? string.Empty;
                    string enterTime = FormatTimeStringTo24Hr(Convert.ToString(emp.EnterTime), "08:00");
                    string exitTime = FormatTimeStringTo24Hr(Convert.ToString(emp.ExitTime), "17:00");
                    string lunchIn = FormatTimeStringTo24Hr(Convert.ToString(emp.LunchInTime), "13:00");
                    string lunchOut = FormatTimeStringTo24Hr(Convert.ToString(emp.LunchOutTime), "14:00");
                    string friLunchIn = FormatTimeStringTo24Hr(Convert.ToString(emp.FridayBreakStartTime), "12:30");
                    string friLunchOut = FormatTimeStringTo24Hr(Convert.ToString(emp.FridayBreakEndTime), "14:00");

                    for (DateTime myDT = dtFrom; myDT <= dtTo; myDT = myDT.AddDays(1))
                    {
                        if (myDT.DayOfWeek == DayOfWeek.Sunday) continue;

                        int holCount = await db.ExecuteScalarAsync<int>("SELECT COUNT(1) FROM Holidays WHERE CAST(DT AS DATE) = @DT", new { DT = myDT });
                        if (holCount > 0) continue;

                        int attCount = await db.ExecuteScalarAsync<int>("SELECT COUNT(1) FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT });
                        if (attCount > 0) continue;

                        using (var trans = db.BeginTransaction())
                        {
                            try
                            {
                                await db.ExecuteAsync("DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT }, trans);
                                await db.ExecuteAsync("DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT }, trans);
                                await db.ExecuteAsync("DELETE FROM Leaves WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT }, trans);

                                int randMinsIn = rand.Next(0, 10);
                                DateTime dtEnter = ParseTimeString(enterTime, myDT).AddMinutes(-randMinsIn);
                                string strInTime = dtEnter.ToString("HH:mm", CultureInfo.InvariantCulture);

                                await ProcessPunchAsync(curEmpId, myDT, strInTime, "System", "127.0.0.1", db, trans);

                                bool isFriday = myDT.DayOfWeek == DayOfWeek.Friday;
                                string lIn = isFriday ? friLunchIn : lunchIn;
                                string lOut = isFriday ? friLunchOut : lunchOut;

                                if (lIn != lOut)
                                {
                                    await ProcessPunchAsync(curEmpId, myDT, lIn, "System", "127.0.0.1", db, trans);
                                    await ProcessPunchAsync(curEmpId, myDT, lOut, "System", "127.0.0.1", db, trans);
                                }

                                int randMinsOut = rand.Next(0, 6);
                                DateTime dtExit = ParseTimeString(exitTime, myDT).AddMinutes(randMinsOut);
                                string strOutTime = dtExit.ToString("HH:mm", CultureInfo.InvariantCulture);

                                await ProcessPunchAsync(curEmpId, myDT, strOutTime, "System", "127.0.0.1", db, trans);

                                trans.Commit();
                                generatedCount++;
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }
                }

                return generatedCount;
            }
        }

        public async Task<int> CompleteMissingAttendanceAsync(string deptId, string empId, DateTime dateFrom, DateTime dateTo)
        {
            DateTime dtFrom = dateFrom.Date;
            DateTime dtTo = dateTo.Date;

            string empCond = " WHERE ISNULL(Employees.Active, 1) = 1";
            if (!string.IsNullOrWhiteSpace(empId) && empId != "0")
            {
                empCond += " AND Employees.EmpID = @EmpID";
            }
            else if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
            {
                empCond += " AND Employees.DeptID = @DeptID";
            }

            string sql = $@"
                SELECT Employees.EmpID,
                       ISNULL(VEmpSettings.LunchInTime, '13:00') AS LunchInTime,
                       ISNULL(VEmpSettings.LunchOutTime, '14:00') AS LunchOutTime,
                       ISNULL(VEmpSettings.FridayBreakStartTime, '12:30') AS FridayBreakStartTime,
                       ISNULL(VEmpSettings.FridayBreakEndTime, '14:00') AS FridayBreakEndTime,
                       ISNULL(VEmpSettings.EnterTime, '08:00') AS EnterTime,
                       ISNULL(VEmpSettings.ExitTime, '17:00') AS ExitTime,
                       T1.DT, T1.FirstInTime
                FROM Employees
                LEFT JOIN VEmpSettings ON Employees.EmpID = VEmpSettings.EmpID
                INNER JOIN (
                    SELECT EmpID, DT, FirstInTime, FirstOutTime, SecondInTime, SecondOutTime
                    FROM VEmpTimes2
                    WHERE CAST(DT AS DATE) BETWEEN @DateFrom AND @DateTo
                      AND (FirstOutTime IS NULL OR SecondInTime IS NULL OR SecondOutTime IS NULL)
                ) T1 ON Employees.EmpID = T1.EmpID
                {empCond}
                ORDER BY Employees.EmpID, T1.DT";

            using (var db = new SqlConnection(_connectionString))
            {
                await db.OpenAsync();
                var records = (await db.QueryAsync(sql, new { DateFrom = dtFrom, DateTo = dtTo, DeptID = deptId, EmpID = empId })).AsList();

                int count = 0;
                Random rand = new Random();

                foreach (var r in records)
                {
                    string curEmpId = Convert.ToString(r.EmpID) ?? string.Empty;
                    DateTime myDT = Convert.ToDateTime(r.DT).Date;
                    string firstIn = r.FirstInTime != null ? FormatTimeStringTo24Hr(Convert.ToString(r.FirstInTime), "08:00") : "08:00";
                    string exitTime = FormatTimeStringTo24Hr(Convert.ToString(r.ExitTime), "17:00");
                    string lunchIn = FormatTimeStringTo24Hr(Convert.ToString(r.LunchInTime), "13:00");
                    string lunchOut = FormatTimeStringTo24Hr(Convert.ToString(r.LunchOutTime), "14:00");
                    string friLunchIn = FormatTimeStringTo24Hr(Convert.ToString(r.FridayBreakStartTime), "12:30");
                    string friLunchOut = FormatTimeStringTo24Hr(Convert.ToString(r.FridayBreakEndTime), "14:00");

                    using (var trans = db.BeginTransaction())
                    {
                        try
                        {
                            await db.ExecuteAsync("DELETE FROM AttendanceSheet WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT }, trans);
                            await db.ExecuteAsync("DELETE FROM EmpTimes WHERE EmpID = @EmpID AND CAST(DT AS DATE) = @DT", new { EmpID = curEmpId, DT = myDT }, trans);

                            await ProcessPunchAsync(curEmpId, myDT, firstIn, "System", "127.0.0.1", db, trans);

                            bool isFriday = myDT.DayOfWeek == DayOfWeek.Friday;
                            string lIn = isFriday ? friLunchIn : lunchIn;
                            string lOut = isFriday ? friLunchOut : lunchOut;

                            if (lIn != lOut)
                            {
                                await ProcessPunchAsync(curEmpId, myDT, lIn, "System", "127.0.0.1", db, trans);
                                await ProcessPunchAsync(curEmpId, myDT, lOut, "System", "127.0.0.1", db, trans);
                            }

                            int randMinsOut = rand.Next(0, 6);
                            DateTime dtExit = ParseTimeString(exitTime, myDT).AddMinutes(randMinsOut);
                            string strOutTime = dtExit.ToString("HH:mm", CultureInfo.InvariantCulture);

                            await ProcessPunchAsync(curEmpId, myDT, strOutTime, "System", "127.0.0.1", db, trans);

                            trans.Commit();
                            count++;
                        }
                        catch
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }

                return count;
            }
        }
    }
}
