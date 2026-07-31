using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Filter criteria for querying the Employee Fine Register.
    /// </summary>
    public class EmpFineLedgerFilterDto
    {
        public string DeptID { get; set; } = "0";
        public string EmpID { get; set; } = "0";
        public DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);
        public DateTime ToDate { get; set; } = DateTime.Today;
        public string? SearchText { get; set; }
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Individual fine / clearance ledger transaction model.
    /// </summary>
    public class EmpFineLedgerItemModel
    {
        public long EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string? Designation { get; set; }
        public DateTime DT { get; set; }
        public string Description { get; set; } = string.Empty;
        public decimal Amount { get; set; }
        public decimal RunningBalance { get; set; }
        public bool IsOpeningBalance { get; set; } = false;
    }

    /// <summary>
    /// Summary totals for KPI metric cards.
    /// </summary>
    public class EmpFineLedgerKpiSummaryModel
    {
        public decimal OpeningBalance { get; set; }
        public decimal TotalFinePosted { get; set; }
        public decimal TotalFineCleared { get; set; }
        public decimal ClosingBalance { get; set; }
    }
}
