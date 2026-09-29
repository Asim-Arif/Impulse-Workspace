using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.Production
{
    public class EmployeePerformanceDataAccess : IEmployeePerformanceDataAccess
    {
        private readonly IConfiguration _config;
        private readonly ILogger<EmployeePerformanceDataAccess> _logger;

        public EmployeePerformanceDataAccess(IConfiguration config, ILogger<EmployeePerformanceDataAccess> logger)
        {
            _config = config;
            _logger = logger;
        }

        private string ConnectionString => _config.GetConnectionString("DefaultConnection")
            ?? _config.GetConnectionString("SMBI_AWM")
            ?? throw new InvalidOperationException("DefaultConnection string is not configured.");

        public async Task<List<DepartmentLookupItem>> GetDepartmentsLookupAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT deptid AS DeptID, name AS DeptName FROM Departments ORDER BY name ASC";
            return (await db.QueryAsync<DepartmentLookupItem>(sql)).ToList();
        }

        public async Task<List<ProcessLookupItem>> GetProcessesLookupAsync()
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = "SELECT ProcessID, Description FROM Processes ORDER BY Description ASC";
            return (await db.QueryAsync<ProcessLookupItem>(sql)).ToList();
        }

        public async Task<List<PerformanceEmployeeLookupItem>> GetEmployeesLookupAsync(int? deptId = null)
        {
            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT EmpID, Name, ISNULL(Designation, '') AS Designation 
                FROM Employees 
                WHERE ISNULL(Active, 1) = 1 
                  AND (@DeptID IS NULL OR deptid = @DeptID)
                ORDER BY Name ASC";

            return (await db.QueryAsync<PerformanceEmployeeLookupItem>(sql, new { DeptID = deptId })).ToList();
        }

        public async Task<EmployeePerformanceDashboardDto> GetPerformanceDashboardAsync(EmployeePerformanceFilter filter)
        {
            var result = new EmployeePerformanceDashboardDto();

            using IDbConnection db = new SqlConnection(ConnectionString);
            var p = new DynamicParameters();

            // Date Range setup
            DateTime dtFrom = filter.DtFrom.Date;
            DateTime dtTo = filter.DtTo.Date.AddDays(1).AddTicks(-1);

            p.Add("@DtFrom", dtFrom);
            p.Add("@DtTo", dtTo);
            p.Add("@DeptID", filter.DepartmentId);
            p.Add("@ProcessID", filter.ProcessId);
            p.Add("@EmpID", filter.EmpId);

            // Calculate working days in range (excluding Sundays)
            int workingDays = 0;
            for (DateTime date = filter.DtFrom.Date; date <= filter.DtTo.Date; date = date.AddDays(1))
            {
                if (date.DayOfWeek != DayOfWeek.Sunday)
                {
                    workingDays++;
                }
            }
            if (workingDays == 0) workingDays = 1;
            result.WorkingDaysInPeriod = workingDays;

            string sql = @"
                ;WITH EmpTargets AS (
                    SELECT 
                        edt.EmpID,
                        ISNULL(SUM(CASE WHEN edt.OverTime = 0 THEN edt.Qty ELSE 0 END), 0) AS DailyCapacity,
                        ISNULL(SUM(CASE WHEN edt.OverTime = 1 THEN edt.Qty ELSE 0 END), 0) AS OvertimeDailyCapacity,
                        MAX(p.Description) AS PrimaryProcessName
                    FROM EmpDailyTargets edt WITH (NOLOCK)
                    LEFT JOIN Processes p WITH (NOLOCK) ON edt.ProcessID = p.ProcessID
                    WHERE (@ProcessID IS NULL OR edt.ProcessID = @ProcessID)
                    GROUP BY edt.EmpID
                ),
                EmpAssigned AS (
                    SELECT 
                        vi.IssEmpID AS EmpID,
                        ISNULL(SUM(vid.IssQty), 0) AS AssignedQty,
                        COUNT(DISTINCT vid.LotNo) AS AssignedLotsCount
                    FROM VendIssued vi WITH (NOLOCK)
                    INNER JOIN VendIssdDetail vid WITH (NOLOCK) ON vi.EntryID = vid.RefID
                    WHERE vi.DT BETWEEN @DtFrom AND @DtTo
                      AND ISNULL(vi.IssEmpID, '') <> ''
                      AND (@ProcessID IS NULL OR vi.ProcessID = @ProcessID OR vid.RcvProcessID = @ProcessID)
                    GROUP BY vi.IssEmpID
                ),
                EmpCompleted AS (
                    SELECT 
                        vi.IssEmpID AS EmpID,
                        ISNULL(SUM(vrd.RcvdQty), 0) AS CompletedQty,
                        COUNT(DISTINCT vrd.LotNo) AS CompletedLotsCount,
                        COUNT(DISTINCT CASE WHEN vr.DT > vid.ReturnDT AND vid.ReturnDT IS NOT NULL THEN vrd.LotNo END) AS DelayedLotsCount
                    FROM VendReceived vr WITH (NOLOCK)
                    INNER JOIN VendRcvdDetail vrd WITH (NOLOCK) ON vr.EntryID = vrd.RefID
                    INNER JOIN VendIssdDetail vid WITH (NOLOCK) ON vrd.Issue_RefID = vid.EntryID
                    INNER JOIN VendIssued vi WITH (NOLOCK) ON vid.RefID = vi.EntryID
                    WHERE vr.DT BETWEEN @DtFrom AND @DtTo
                      AND ISNULL(vi.IssEmpID, '') <> ''
                      AND (@ProcessID IS NULL OR vr.ProcessID = @ProcessID)
                    GROUP BY vi.IssEmpID
                ),
                EmpActiveWip AS (
                    SELECT 
                        vi.IssEmpID AS EmpID,
                        ISNULL(SUM(vid.IssQty - vid.RcvdQty), 0) AS ActiveBalanceQty,
                        COUNT(DISTINCT vid.LotNo) AS ActiveLotsCount,
                        COUNT(DISTINCT CASE WHEN vid.ReturnDT < CAST(GETDATE() AS DATE) AND vid.ReturnDT IS NOT NULL THEN vid.LotNo END) AS OverdueLotsCount,
                        ISNULL(SUM(CASE WHEN vid.ReturnDT < CAST(GETDATE() AS DATE) AND vid.ReturnDT IS NOT NULL THEN (vid.IssQty - vid.RcvdQty) ELSE 0 END), 0) AS OverdueQty
                    FROM VendIssued vi WITH (NOLOCK)
                    INNER JOIN VendIssdDetail vid WITH (NOLOCK) ON vi.EntryID = vid.RefID
                    WHERE vid.IssQty > vid.RcvdQty
                      AND vid.LotNo <> '0'
                      AND ISNULL(vi.IssEmpID, '') <> ''
                      AND NOT EXISTS (SELECT 1 FROM Lots_Closed lc WITH (NOLOCK) WHERE lc.LotNo = vid.LotNo)
                      AND (@ProcessID IS NULL OR vi.ProcessID = @ProcessID OR vid.RcvProcessID = @ProcessID)
                    GROUP BY vi.IssEmpID
                )
                SELECT 
                    e.EmpID,
                    e.Name AS EmployeeName,
                    ISNULL(e.Designation, '') AS Designation,
                    ISNULL(d.name, '') AS DepartmentName,
                    ISNULL(t.PrimaryProcessName, 'Production') AS PrimaryProcessName,
                    ISNULL(t.DailyCapacity, 0) AS DailyCapacity,
                    ISNULL(t.OvertimeDailyCapacity, 0) AS OvertimeDailyCapacity,
                    ISNULL(a.AssignedQty, 0) AS AssignedQty,
                    ISNULL(a.AssignedLotsCount, 0) AS AssignedLotsCount,
                    ISNULL(c.CompletedQty, 0) AS CompletedQty,
                    ISNULL(c.CompletedLotsCount, 0) AS CompletedLotsCount,
                    ISNULL(c.DelayedLotsCount, 0) AS DelayedLotsCount,
                    ISNULL(w.ActiveBalanceQty, 0) AS ActiveBalanceQty,
                    ISNULL(w.ActiveLotsCount, 0) AS ActiveLotsCount,
                    ISNULL(w.OverdueLotsCount, 0) AS OverdueLotsCount,
                    ISNULL(w.OverdueQty, 0) AS OverdueQty
                FROM Employees e WITH (NOLOCK)
                LEFT JOIN Departments d WITH (NOLOCK) ON e.deptid = d.deptid
                LEFT JOIN EmpTargets t ON e.EmpID = t.EmpID
                LEFT JOIN EmpAssigned a ON e.EmpID = a.EmpID
                LEFT JOIN EmpCompleted c ON e.EmpID = c.EmpID
                LEFT JOIN EmpActiveWip w ON e.EmpID = w.EmpID
                WHERE ISNULL(e.Active, 1) = 1
                  AND (ISNULL(t.DailyCapacity, 0) > 0 OR ISNULL(a.AssignedQty, 0) > 0 OR ISNULL(w.ActiveBalanceQty, 0) > 0 OR ISNULL(c.CompletedQty, 0) > 0)
                  AND (@DeptID IS NULL OR e.deptid = @DeptID)
                  AND (@EmpID IS NULL OR @EmpID = '' OR e.EmpID = @EmpID)
                ORDER BY e.Name ASC";

            var list = (await db.QueryAsync<EmployeePerformanceSummaryDto>(sql, p)).ToList();

            foreach (var emp in list)
            {
                emp.WorkingDays = workingDays;
            }

            // Status Filtering if applicable
            if (filter.StatusFilter == 1) // Overloaded
            {
                list = list.Where(x => x.PerformanceStatus == "Overloaded").ToList();
            }
            else if (filter.StatusFilter == 2) // Has Overdue
            {
                list = list.Where(x => x.OverdueLotsCount > 0).ToList();
            }
            else if (filter.StatusFilter == 3) // Optimal
            {
                list = list.Where(x => x.PerformanceStatus == "Optimal").ToList();
            }

            result.Employees = list;
            result.TotalWorkers = list.Count;
            result.TotalAssignedQty = list.Sum(x => x.AssignedQty);
            result.TotalCompletedQty = list.Sum(x => x.CompletedQty);
            result.TotalActiveWipQty = list.Sum(x => x.ActiveBalanceQty);
            result.TotalOverdueLots = list.Sum(x => x.OverdueLotsCount);

            int totalPeriodCapacity = list.Sum(x => x.PeriodCapacity);
            result.OverallUtilizationPct = totalPeriodCapacity > 0
                ? Math.Min(200.0, Math.Round((result.TotalCompletedQty * 100.0) / totalPeriodCapacity, 1))
                : 0;

            return result;
        }

        public async Task<List<EmployeeLotDetailDto>> GetEmployeeLotsDetailAsync(string empId, DateTime dtFrom, DateTime dtTo)
        {
            if (string.IsNullOrWhiteSpace(empId))
                return new List<EmployeeLotDetailDto>();

            using IDbConnection db = new SqlConnection(ConnectionString);
            const string sql = @"
                SELECT 
                    vid.EntryID,
                    vid.LotNo,
                    ISNULL(vid.OrderNo, '') AS OrderNo,
                    vid.ItemCode,
                    ISNULL(i.ItemName, vid.ItemCode) AS ItemName,
                    COALESCE(vid.RcvProcessID, vi.ProcessID, 0) AS ProcessID,
                    ISNULL(p.Description, '') AS ProcessName,
                    vi.DT AS IssDT,
                    vid.ReturnDT,
                    ISNULL(vid.IssQty, 0) AS IssQty,
                    ISNULL(vid.RcvdQty, 0) AS RcvdQty,
                    CASE WHEN EXISTS (
                        SELECT 1 FROM VendReceived vr 
                        INNER JOIN VendRcvdDetail vrd ON vr.EntryID = vrd.RefID 
                        WHERE vrd.Issue_RefID = vid.EntryID AND vr.DT > vid.ReturnDT AND vid.ReturnDT IS NOT NULL
                    ) THEN 1 ELSE 0 END AS IsDelayed
                FROM VendIssued vi WITH (NOLOCK)
                INNER JOIN VendIssdDetail vid WITH (NOLOCK) ON vi.EntryID = vid.RefID
                LEFT JOIN Items i WITH (NOLOCK) ON vid.ItemCode = i.ItemID
                LEFT JOIN Processes p WITH (NOLOCK) ON COALESCE(vid.RcvProcessID, vi.ProcessID) = p.ProcessID
                WHERE vi.IssEmpID = @EmpID
                  AND (vi.DT BETWEEN @DtFrom AND @DtTo OR (vid.IssQty > vid.RcvdQty AND vid.LotNo <> '0'))
                ORDER BY vi.DT DESC, vid.LotNo ASC";

            return (await db.QueryAsync<EmployeeLotDetailDto>(sql, new
            {
                EmpID = empId.Trim(),
                DtFrom = dtFrom.Date,
                DtTo = dtTo.Date.AddDays(1).AddTicks(-1)
            })).ToList();
        }
    }
}
