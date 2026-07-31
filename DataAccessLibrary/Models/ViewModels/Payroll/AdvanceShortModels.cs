using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Determines which payment account type is selected, which in turn
    /// dictates the voucher type: Cash→CPV, OtherAccount→JV, Bank→BPV.
    /// </summary>
    public enum AdvanceShortPaymentMode
    {
        Cash = 0,
        OtherAccount = 1,
        Bank = 2
    }

    /// <summary>
    /// DTO passed from the Blazor page to the service layer when saving
    /// a Short-Term Advance entry.
    /// </summary>
    public class AdvanceShortDto
    {
        // ── Employee ──────────────────────────────────────────────────────────
        public string EmpID { get; set; } = string.Empty;

        // ── Advance Details ───────────────────────────────────────────────────
        /// <summary>Issue date of the advance (stored in Advances.DT)</summary>
        public DateTime IssueDate { get; set; } = DateTime.Today;

        /// <summary>
        /// Deduction start date for the short term advance.
        /// Stored in Advances.DeductionStartDT.
        /// </summary>
        public DateTime DeductionStartDT { get; set; } = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

        /// <summary>Short term advance amount (Advances.Amount)</summary>
        public decimal Amount { get; set; }

        /// <summary>Description / narration (Advances.Description)</summary>
        public string Description { get; set; } = "Short Term";

        // ── Payment Mode ──────────────────────────────────────────────────────
        public AdvanceShortPaymentMode PaymentMode { get; set; } = AdvanceShortPaymentMode.Cash;

        /// <summary>
        /// AccNo of the payment-side account.
        /// Cash→Cash-in-Hand AccNo, OtherAccount→any AccNo, Bank→Bank AccNo.
        /// </summary>
        public string PaymentAccNo { get; set; } = string.Empty;

        // ── Cheque Details (BPV only — null/empty for Cash & JV) ─────────────
        public int      BankID      { get; set; }
        public string   ChequeNo    { get; set; } = string.Empty;
        public int      ChqBookNo   { get; set; }
        public string   ChequeType  { get; set; } = string.Empty;   // "Open" or "Payees A/C Only"
        public DateTime ChequeDate  { get; set; } = DateTime.Today;

        // ── Contractor flag ──────────────────────────────────────────────────
        /// <summary>True when the form was opened in contractor mode.</summary>
        public bool IsContractor { get; set; } = false;
    }

    /// <summary>
    /// Read-only summary shown on-screen when an employee is selected.
    /// Fetched from Advances table for Type = 0 (Short Term).
    /// </summary>
    public class AdvanceShortSummaryModel
    {
        /// <summary>Outstanding short-term advance balance for the employee.</summary>
        public decimal PreviousShortTermBalance { get; set; }
    }
}
