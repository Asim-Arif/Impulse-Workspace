using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Determines which payment account type is selected, which in turn
    /// dictates the voucher type: Cash→CPV, OtherAccount→JV, Bank→BPV.
    /// </summary>
    public enum AdvanceLongPaymentMode
    {
        Cash = 0,
        OtherAccount = 1,
        Bank = 2
    }

    /// <summary>
    /// DTO passed from the Blazor page to the service layer when saving
    /// a Long-Term Advance entry.
    /// </summary>
    public class AdvanceLongDto
    {
        // ── Employee ──────────────────────────────────────────────────────────
        public string EmpID { get; set; } = string.Empty;

        // ── Advance Details ───────────────────────────────────────────────────
        /// <summary>Issue date of the loan (stored in Advances.DT)</summary>
        public DateTime IssueDate { get; set; } = DateTime.Today;

        /// <summary>
        /// The first month in which monthly deductions begin.
        /// Stored in Advances.DeductionStartDT.
        /// Defaults to 1st of the month following the issue date.
        /// </summary>
        public DateTime DeductionStartDT { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

        /// <summary>Total loan amount (Advances.Amount)</summary>
        public decimal Amount { get; set; }

        /// <summary>Monthly deduction amount (Advances.DAmount)</summary>
        public decimal DAmount { get; set; }

        /// <summary>Description / narration (Advances.Description)</summary>
        public string Description { get; set; } = "Long Term Loan";

        // ── Payment Mode ──────────────────────────────────────────────────────
        public AdvanceLongPaymentMode PaymentMode { get; set; } = AdvanceLongPaymentMode.Cash;

        /// <summary>
        /// AccNo of the payment-side account.
        /// Cash→Cash-in-Hand AccNo, OtherAccount→any AccNo, Bank→Bank AccNo.
        /// </summary>
        public string PaymentAccNo { get; set; } = string.Empty;

        // ── Cheque Details (BPV only — null/empty for Cash & JV) ─────────────
        public int    BankID      { get; set; }
        public string ChequeNo    { get; set; } = string.Empty;
        public int    ChqBookNo   { get; set; }
        public string ChequeType  { get; set; } = string.Empty;   // "Open" or "Payees A/C Only"
        public DateTime ChequeDate { get; set; } = DateTime.Today;

        // ── Contractor flag ──────────────────────────────────────────────────
        /// <summary>True when the form was opened in contractor mode.</summary>
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Read-only summary shown on-screen when an employee is selected.
    /// Fetched from Advances + AmtCleared tables.
    /// </summary>
    public class AdvanceLongSummaryModel
    {
        /// <summary>Outstanding long-term loan balance for the employee.</summary>
        public decimal PreviousBalance { get; set; }

        /// <summary>Monthly deduction from their last advance record.</summary>
        public decimal PreviousDeductionAmount { get; set; }
    }
}
