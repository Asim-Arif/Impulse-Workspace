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
    public class GatePassDataAccess : IGatePassDataAccess
    {
        private readonly string _connectionString;

        public GatePassDataAccess(IConfiguration config)
        {
            _connectionString = config.GetConnectionString("DefaultConnection");
        }

        public async Task<List<GatePassRowModel>> GetGatePassListAsync(string deptId, DateTime date)
        {
            var results = new List<GatePassRowModel>();

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // Query active present employees for the date
                string empCond = "WHERE Active = 1 AND EmpID IN (SELECT EmpID FROM AttendanceSheet WHERE Attendance = 1 AND CAST(DT AS DATE) = @Date)";
                if (!string.IsNullOrWhiteSpace(deptId) && deptId != "0")
                {
                    empCond += " AND DeptID = @DeptID";
                }

                string empSql = $"SELECT RTRIM(LTRIM(EmpID)) AS EmpID, Name, ISNULL(FName,'') AS FName, ISNULL(Rel,'') AS Rel FROM Employees {empCond} ORDER BY EmpID";

                var employees = (await db.QueryAsync<(string EmpID, string Name, string FName, string Rel)>(empSql, new { Date = date.Date, DeptID = deptId })).ToList();

                // Query existing gate passes from EmpGPHrs for the date
                string gpSql = "SELECT EntryID, RTRIM(LTRIM(EmpID)) AS EmpID, GPHrs, ISNULL(Time1,'') AS Time1, ISNULL(Time2,'') AS Time2, ISNULL(Remarks,'') AS Remarks, ISNULL(Official,0) AS Official FROM EmpGPHrs WHERE CAST(DT AS DATE) = @Date";

                var gatePasses = (await db.QueryAsync<dynamic>(gpSql, new { Date = date.Date }))
                    .ToDictionary(g => (string)g.EmpID, g => g);

                foreach (var emp in employees)
                {
                    string fullName = emp.Name;
                    if (!string.IsNullOrWhiteSpace(emp.Rel) || !string.IsNullOrWhiteSpace(emp.FName))
                    {
                        fullName += $" ({emp.Rel} {emp.FName})".Trim();
                    }

                    if (gatePasses.TryGetValue(emp.EmpID, out dynamic? gp) && gp != null)
                    {
                        double gpHrs = 0;
                        if (gp.GPHrs != null)
                        {
                            double.TryParse(gp.GPHrs.ToString(), out gpHrs);
                        }

                        results.Add(new GatePassRowModel
                        {
                            EntryID    = (int)gp.EntryID,
                            EmpID      = emp.EmpID,
                            EmpName    = fullName,
                            IsSelected = true,
                            OutTime    = SanitizeTime5((string)gp.Time1),
                            InTime     = SanitizeTime5((string)gp.Time2),
                            GpHrs      = gpHrs,
                            IsOfficial = (bool)gp.Official,
                            Remarks    = (string)gp.Remarks
                        });
                    }
                    else
                    {
                        results.Add(new GatePassRowModel
                        {
                            EntryID    = 0,
                            EmpID      = emp.EmpID,
                            EmpName    = fullName,
                            IsSelected = false,
                            OutTime    = string.Empty,
                            InTime     = string.Empty,
                            GpHrs      = 0,
                            IsOfficial = false,
                            Remarks    = string.Empty
                        });
                    }
                }
            }

            return results;
        }

        public async Task<bool> SaveGatePassesAsync(GatePassSaveDto input)
        {
            if (!input.Rows.Any()) return true;

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                using (IDbTransaction trans = db.BeginTransaction())
                {
                    try
                    {
                        DateTime date = input.Date.Date;

                        foreach (var row in input.Rows)
                        {
                            string cleanOutTime = SanitizeTime5(row.OutTime);
                            string cleanInTime  = SanitizeTime5(row.InTime);

                            if (row.IsSelected)
                            {
                                if (row.EntryID > 0)
                                {
                                    // Update existing gate pass record
                                    string updateSql = @"
                                        UPDATE EmpGPHrs
                                        SET Time1 = @Time1, Time2 = @Time2, GPHrs = @GPHrs, Official = @Official, Remarks = @Remarks
                                        WHERE EntryID = @EntryID";

                                    await db.ExecuteAsync(updateSql, new
                                    {
                                        Time1    = cleanOutTime,
                                        Time2    = cleanInTime,
                                        GPHrs    = row.GpHrs,
                                        Official = row.IsOfficial ? 1 : 0,
                                        Remarks  = row.Remarks,
                                        EntryID  = row.EntryID
                                    }, trans);
                                }
                                else
                                {
                                    // Insert new gate pass record
                                    string insertSql = @"
                                        INSERT INTO EmpGPHrs(EmpID, DT, Time1, Time2, GPHrs, Official, Remarks)
                                        VALUES(@EmpID, @DT, @Time1, @Time2, @GPHrs, @Official, @Remarks)";

                                    await db.ExecuteAsync(insertSql, new
                                    {
                                        EmpID    = row.EmpID,
                                        DT       = date,
                                        Time1    = cleanOutTime,
                                        Time2    = cleanInTime,
                                        GPHrs    = row.GpHrs,
                                        Official = row.IsOfficial ? 1 : 0,
                                        Remarks  = row.Remarks
                                    }, trans);
                                }
                            }
                            else if (row.EntryID > 0)
                            {
                                // Delete unselected gate pass record
                                string delSql = "DELETE FROM EmpGPHrs WHERE EntryID = @EntryID";
                                await db.ExecuteAsync(delSql, new { EntryID = row.EntryID }, trans);
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

        private static string SanitizeTime5(string? val)
        {
            if (string.IsNullOrWhiteSpace(val)) return string.Empty;
            val = val.Trim();
            if (TimeSpan.TryParse(val, out TimeSpan ts))
                return ts.ToString(@"hh\:mm");
            if (DateTime.TryParse(val, out DateTime dt))
                return dt.ToString("HH:mm");
            if (val.Length > 5) return val.Substring(0, 5);
            return val;
        }
    }
}
