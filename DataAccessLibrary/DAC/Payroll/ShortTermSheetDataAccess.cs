using Dapper;
using DataAccessLibrary.Interface.Payroll;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace DataAccessLibrary.DAC.Payroll
{
    public class ShortTermSheetDataAccess : IShortTermSheetDataAccess
    {
        private readonly IConfiguration _config;

        public ShortTermSheetDataAccess(IConfiguration config)
        {
            _config = config;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection");

        public async Task<List<ShortTermSheetItemViewModel>> GetShortTermSheetDataAsync(DateTime date, string deptId)
        {
            var result = new List<ShortTermSheetItemViewModel>();
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

            using IDbConnection db = new SqlConnection(ConnectionString);

            // 1. Calculate calculation days limit
            string sqlDays = "SELECT ISNULL(MAX(NoofDays), 0) FROM Employees_ST_Sheet_Calculate_For_Days WHERE MONTH(DT)=@Month AND YEAR(DT)=@Year";
            int daysForCalc = await db.ExecuteScalarAsync<int>(sqlDays, new { Month = date.Month, Year = date.Year });
            if (daysForCalc <= 0) daysForCalc = 15;

            // 2. Count holidays & Sundays
            string sqlHolidays = "SELECT COUNT(DT) FROM Holidays WHERE MONTH(DT)=@Month AND YEAR(DT)=@Year AND DAY(DT)<=@Days";
            int holidayCount = await db.ExecuteScalarAsync<int>(sqlHolidays, new { Month = date.Month, Year = date.Year, Days = daysForCalc });

            int totalSundays = 0;
            for (int day = 1; day <= daysForCalc; day++)
            {
                var current = new DateTime(date.Year, date.Month, day);
                if (current.DayOfWeek == DayOfWeek.Sunday)
                {
                    totalSundays++;
                }
            }
            holidayCount += totalSundays;

            // 3. Stored Procedure Execution
            var procDate = new DateTime(date.Year, date.Month, daysForCalc);
            IEnumerable<dynamic> rawRows;

            if (string.IsNullOrWhiteSpace(deptId) || deptId == "0")
            {
                rawRows = await db.QueryAsync("VAdvanceSalary", new { ParDate = procDate, EmpID = "%" }, commandType: CommandType.StoredProcedure);
            }
            else
            {
                rawRows = await db.QueryAsync("VAdvanceSalaryDeptWise", new { ParDate = procDate, DeptID = deptId }, commandType: CommandType.StoredProcedure);
            }

            // 4. Query existing saved entries in Employees_ST_Sheet
            string sqlSaved = "SELECT EmpID, ISNULL(Demand,0) AS Demand, ISNULL(Approved,0) AS Approved FROM Employees_ST_Sheet WHERE DT=@DT";
            var savedDict = (await db.QueryAsync(sqlSaved, new { DT = lastDayOfMonth }))
                .ToDictionary(r => (string)r.EmpID, r => (Demand: Convert.ToDecimal(r.Demand), Approved: Convert.ToDecimal(r.Approved)));

            // 5. Map and compute payable calculations
            foreach (var row in rawRows)
            {
                var dict = (IDictionary<string, object>)row;

                string empId = GetString(dict, "EmpID", "EmpId");
                string name = GetString(dict, "Name");
                string empIdOld = GetString(dict, "EmpIDOld");
                string designation = GetString(dict, "Designation");
                string deptName = GetString(dict, "DeptName");
                string dId = GetString(dict, "DeptID");

                decimal startingSalary = GetDecimal(dict, "StartingSalary");
                double presentDays = GetDouble(dict, "PresentDays");
                double leaves = GetDouble(dict, "Leaves");
                double wpLeaves = GetDouble(dict, "WPLeaves");
                decimal payed = GetDecimal(dict, "Payed");
                decimal amtToDeduct = GetDecimal(dict, "AmtToDeduct");
                double otHours = GetDouble(dict, "OverTimeHours");
                double otRate = GetDouble(dict, "OTRate");
                if (otRate <= 0) otRate = 1.0;

                decimal dailyRate = startingSalary / 30m;
                double effectiveDays = presentDays + holidayCount + leaves - wpLeaves;
                if (effectiveDays > daysForCalc) effectiveDays = daysForCalc;

                decimal salaryPayable = dailyRate * (decimal)effectiveDays;
                salaryPayable = salaryPayable - payed - amtToDeduct;
                if (salaryPayable < 0) salaryPayable = 0;

                decimal otValue = (decimal)otHours * (dailyRate / 8m) * (decimal)otRate;
                salaryPayable = salaryPayable + otValue;
                if (salaryPayable > 0)
                {
                    salaryPayable = Math.Round(salaryPayable / 500m) * 500m;
                }

                decimal demand = salaryPayable;
                decimal approved = salaryPayable;

                if (savedDict.TryGetValue(empId, out var savedVal))
                {
                    demand = savedVal.Demand > 0 ? savedVal.Demand : salaryPayable;
                    approved = savedVal.Approved;
                }

                result.Add(new ShortTermSheetItemViewModel
                {
                    EmpID = empId,
                    Name = name,
                    EmpIDOld = empIdOld,
                    Designation = designation,
                    DeptName = deptName,
                    DeptID = dId,
                    StartingSalary = startingSalary,
                    PresentDays = presentDays,
                    Leaves = leaves,
                    WPLeaves = wpLeaves,
                    Payed = payed,
                    AmtToDeduct = amtToDeduct,
                    OverTimeHours = otHours,
                    OTRate = otRate,
                    CalculatedPayable = salaryPayable,
                    Demand = demand,
                    Approved = approved
                });
            }

            return result.OrderBy(r => r.EmpID).ToList();
        }

        public async Task SaveShortTermSheetEntriesAsync(DateTime date, IEnumerable<ShortTermSheetSaveDto> items)
        {
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string updateSql = "UPDATE Employees_ST_Sheet SET Demand=@Demand, Approved=@Approved WHERE EmpID=@EmpID AND DT=@DT";
                string insertSql = "INSERT INTO Employees_ST_Sheet(EmpID, DT, Demand, Approved) VALUES(@EmpID, @DT, @Demand, @Approved)";

                foreach (var item in items)
                {
                    int rows = await db.ExecuteAsync(updateSql, new
                    {
                        Demand = item.Demand,
                        Approved = item.Approved,
                        EmpID = item.EmpID,
                        DT = lastDayOfMonth
                    }, trans);

                    if (rows == 0)
                    {
                        await db.ExecuteAsync(insertSql, new
                        {
                            EmpID = item.EmpID,
                            DT = lastDayOfMonth,
                            Demand = item.Demand,
                            Approved = item.Approved
                        }, trans);
                    }
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }

        public async Task<ShortTermPostingStatusModel> GetPostingStatusAsync(DateTime date, string deptId)
        {
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            using IDbConnection db = new SqlConnection(ConnectionString);

            string sql = @"SELECT TOP 1 CashVchrNo, BankVchrNo, UserName AS PostedBy, DTEntry AS PostedDate 
                           FROM Employees_ST_Sheet_Posting 
                           WHERE MONTH(DT)=@Month AND YEAR(DT)=@Year AND DeptID=@DeptID 
                           ORDER BY EntryID DESC";

            var row = await db.QueryFirstOrDefaultAsync(sql, new { Month = date.Month, Year = date.Year, DeptID = deptId });

            if (row != null)
            {
                return new ShortTermPostingStatusModel
                {
                    IsPosted = true,
                    CashVchrNo = Convert.ToString(row.CashVchrNo) ?? string.Empty,
                    BankVchrNo = Convert.ToString(row.BankVchrNo) ?? string.Empty,
                    PostedBy = Convert.ToString(row.PostedBy) ?? string.Empty,
                    PostedDate = row.PostedDate != null ? Convert.ToDateTime(row.PostedDate) : null
                };
            }

            return new ShortTermPostingStatusModel { IsPosted = false };
        }

        public async Task SavePostingRecordAsync(DateTime date, string deptId, string vchrNo, string bankVchrNo, string chqNo, DateTime chqDT, int postingType, string userName, string machineName, IEnumerable<ShortTermSheetItemViewModel> items)
        {
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            using IDbConnection db = new SqlConnection(ConnectionString);
            db.Open();
            using var trans = db.BeginTransaction();

            try
            {
                string sqlPost = @"INSERT INTO Employees_ST_Sheet_Posting (DT, CashVchrNo, BankVchrNo, UserName, MachineName, DTEntry, ChqNo, ChqDT, DeptID, Posting_Type)
                                   VALUES (@DT, @CashVchrNo, @BankVchrNo, @UserName, @MachineName, GETDATE(), @ChqNo, @ChqDT, @DeptID, @PostingType)";

                await db.ExecuteAsync(sqlPost, new
                {
                    DT = lastDayOfMonth,
                    CashVchrNo = vchrNo,
                    BankVchrNo = bankVchrNo,
                    UserName = userName,
                    MachineName = machineName,
                    ChqNo = chqNo,
                    ChqDT = chqDT,
                    DeptID = deptId,
                    PostingType = postingType
                }, trans);

                string sqlAdv = @"INSERT INTO Advances(EmpID, DT, Type, Description, Amount, DAmount, AmountCleared, DeductionStartDT, Cleared, ActualAmountTaken, AccVoucherNo)
                                  VALUES(@EmpID, @DT, 0, 'Advance Salary Short Term', @Amount, 0, 0, @DT, 0, @Amount, @VchrNo)";

                foreach (var item in items.Where(x => x.Approved > 0))
                {
                    await db.ExecuteAsync(sqlAdv, new
                    {
                        EmpID = item.EmpID,
                        DT = lastDayOfMonth,
                        Amount = item.Approved,
                        VchrNo = vchrNo
                    }, trans);
                }

                trans.Commit();
            }
            catch
            {
                trans.Rollback();
                throw;
            }
        }


        private static Dictionary<string, object> CreateCaseInsensitiveDict(IDictionary<string, object> dict)
        {
            var ciDict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in dict)
            {
                if (kvp.Key != null && !ciDict.ContainsKey(kvp.Key))
                {
                    ciDict[kvp.Key] = kvp.Value;
                }
            }
            return ciDict;
        }

        private static string GetString(IDictionary<string, object> dict, params string[] keys)
        {
            var ciDict = CreateCaseInsensitiveDict(dict);
            foreach (var key in keys)
            {
                if (ciDict.TryGetValue(key, out var val) && val != null)
                {
                    return val.ToString() ?? string.Empty;
                }
            }
            return string.Empty;
        }

        private static decimal GetDecimal(IDictionary<string, object> dict, params string[] keys)
        {
            var ciDict = CreateCaseInsensitiveDict(dict);
            foreach (var key in keys)
            {
                if (ciDict.TryGetValue(key, out var val) && val != null && decimal.TryParse(val.ToString(), out decimal res))
                {
                    return res;
                }
            }
            return 0m;
        }

        private static double GetDouble(IDictionary<string, object> dict, params string[] keys)
        {
            var ciDict = CreateCaseInsensitiveDict(dict);
            foreach (var key in keys)
            {
                if (ciDict.TryGetValue(key, out var val) && val != null && double.TryParse(val.ToString(), out double res))
                {
                    return res;
                }
            }
            return 0;
        }
    }
}


