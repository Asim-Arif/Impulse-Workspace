using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Filter criteria for querying the Loan & Advance Register.
    /// </summary>
    public class AdvanceLedgerFilterDto
    {
        /// <summary>Department ID ("0" = All)</summary>
        public string DeptID { get; set; } = "0";

        /// <summary>Employee ID ("0" = All)</summary>
        public string EmpID { get; set; } = "0";

        /// <summary>-1 = All, 0 = Short Term, 1 = Long Term</summary>
        public int LoanType { get; set; } = -1;

        /// <summary>Start Date filter</summary>
        public DateTime FromDate { get; set; } = DateTime.Today.AddMonths(-1);

        /// <summary>End Date filter</summary>
        public DateTime ToDate { get; set; } = DateTime.Today;

        /// <summary>Quick search text (Name, ID, Voucher #)</summary>
        public string? SearchText { get; set; }

        /// <summary>Contractor mode flag</summary>
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Individual record representing an advance/loan transaction in VAdvances.
    /// </summary>
    public class AdvanceLedgerItemModel
    {
        public long EntryID { get; set; }
        public string EmpID { get; set; } = string.Empty;
        public string EmpName { get; set; } = string.Empty;
        public string DeptID { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public string? Designation { get; set; }
        public DateTime DT { get; set; }

        /// <summary>0 = Short Term, 1 = Long Term</summary>
        public int Type { get; set; }

        public decimal Amount { get; set; }
        public decimal DAmount { get; set; }
        public decimal AmountCleared { get; set; }
        public decimal Balance { get; set; }
        public string? AccVoucherNo { get; set; }

        /// <summary>0 = Uncleared/Active, >0 = Cleared</summary>
        public int Cleared { get; set; }

        public DateTime? DeductionStartDT { get; set; }
    }

    /// <summary>
    /// Summary totals model for KPI metric cards.
    /// </summary>
    public class AdvanceLedgerKpiSummaryModel
    {
        public decimal TotalIssued { get; set; }
        public decimal TotalCleared { get; set; }
        public decimal TotalOutstanding { get; set; }
    }
}
