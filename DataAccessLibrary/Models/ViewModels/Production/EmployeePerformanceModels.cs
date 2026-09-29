using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class EmployeePerformanceFilter
    {
        public int DateRangeType { get; set; } = 3; // 0=Today, 1=This Week, 2=Last 15d, 3=This Month, 4=Last 30d, 5=Custom
        public DateTime DtFrom { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        public DateTime DtTo { get; set; } = DateTime.Today;
        public int? DepartmentId { get; set; }
        public int? ProcessId { get; set; }
        public string? EmpId { get; set; }
        public int StatusFilter { get; set; } = 0; // 0=All, 1=Overloaded, 2=Has Overdue, 3=Optimal
        public string? SearchText { get; set; }
    }

    public class EmployeePerformanceSummaryDto
    {
        public string EmpID { get; set; } = string.Empty;
        public string EmployeeName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string DepartmentName { get; set; } = string.Empty;
        public string PrimaryProcessName { get; set; } = string.Empty;

        public int DailyCapacity { get; set; }
        public int OvertimeDailyCapacity { get; set; }
        public int WorkingDays { get; set; }
        public int PeriodCapacity => DailyCapacity * Math.Max(1, WorkingDays);

        public int AssignedQty { get; set; }
        public int AssignedLotsCount { get; set; }
        public int CompletedQty { get; set; }
        public int CompletedLotsCount { get; set; }
        public int ActiveBalanceQty { get; set; }
        public int ActiveLotsCount { get; set; }

        public int DelayedLotsCount { get; set; }
        public int OverdueLotsCount { get; set; }
        public int OverdueQty { get; set; }

        public double UtilizationPct => PeriodCapacity > 0
            ? Math.Min(200.0, Math.Round((CompletedQty * 100.0) / PeriodCapacity, 1))
            : 0;

        public double LoadPct => PeriodCapacity > 0
            ? Math.Min(200.0, Math.Round(((ActiveBalanceQty + AssignedQty) * 100.0) / PeriodCapacity, 1))
            : 0;

        public double OnTimeCompletionRate => CompletedLotsCount > 0
            ? Math.Max(0.0, Math.Round(((CompletedLotsCount - DelayedLotsCount) * 100.0) / CompletedLotsCount, 1))
            : 100.0;

        public string PerformanceStatus
        {
            get
            {
                if (OverdueLotsCount > 0) return "Action Required";
                if (UtilizationPct >= 110 || LoadPct >= 110) return "Overloaded";
                if (UtilizationPct >= 75) return "Optimal";
                if (ActiveBalanceQty > 0 || AssignedQty > 0) return "In Progress";
                return "Under-Utilized";
            }
        }

        public string PerformanceStatusBadgeClass => PerformanceStatus switch
        {
            "Action Required" => "bg-danger text-white",
            "Overloaded" => "bg-warning text-dark",
            "Optimal" => "bg-success text-white",
            "In Progress" => "bg-info text-dark",
            _ => "bg-secondary text-white"
        };
    }

    public class EmployeeLotDetailDto
    {
        public long EntryID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public DateTime? IssDT { get; set; }
        public DateTime? ReturnDT { get; set; }
        public int IssQty { get; set; }
        public int RcvdQty { get; set; }
        public int BalanceQty => Math.Max(0, IssQty - RcvdQty);
        public bool IsOverdue => ReturnDT.HasValue && ReturnDT.Value.Date < DateTime.Today && BalanceQty > 0;
        public int DaysOverdue => ReturnDT.HasValue && IsOverdue ? (DateTime.Today - ReturnDT.Value.Date).Days : 0;
        public bool IsDelayed { get; set; }

        public string Status
        {
            get
            {
                if (BalanceQty == 0 && IssQty > 0)
                {
                    return IsDelayed ? "Completed (Delayed)" : "Completed (On-Time)";
                }
                if (IsOverdue) return $"Overdue by {DaysOverdue}d";
                return "In-Process";
            }
        }

        public string StatusBadgeClass => Status switch
        {
            "Completed (On-Time)" => "bg-success text-white",
            "Completed (Delayed)" => "bg-warning text-dark",
            var s when s.StartsWith("Overdue") => "bg-danger text-white",
            _ => "bg-primary text-white"
        };
    }

    public class EmployeePerformanceDashboardDto
    {
        public int TotalWorkers { get; set; }
        public int TotalAssignedQty { get; set; }
        public int TotalCompletedQty { get; set; }
        public int TotalActiveWipQty { get; set; }
        public int TotalOverdueLots { get; set; }
        public int WorkingDaysInPeriod { get; set; }
        public double OverallUtilizationPct { get; set; }

        public List<EmployeePerformanceSummaryDto> Employees { get; set; } = new();
    }

    public class DepartmentLookupItem
    {
        public int DeptID { get; set; }
        public string DeptName { get; set; } = string.Empty;
    }

    public class ProcessLookupItem
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class PerformanceEmployeeLookupItem
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string DisplayText => !string.IsNullOrEmpty(EmpID) ? $"[{EmpID}] {Name}" : Name;
    }
}
