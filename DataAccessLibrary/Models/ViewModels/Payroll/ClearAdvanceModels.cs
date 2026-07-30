using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    /// <summary>
    /// Determines which receipt account type is selected, which in turn
    /// dictates the voucher type: Cash (Cash In Hand)→CRV, OtherAccount→JV.
    /// </summary>
    public enum ClearAdvanceReceiptMode
    {
        Cash = 0,
        OtherAccount = 1
    }

    /// <summary>
    /// DTO passed from the Blazor page to the service layer when saving
    /// a Long-Term Loan clearance (repayment) entry.
    /// </summary>
    public class ClearAdvanceDto
    {
        // ── Employee ──────────────────────────────────────────────────────────
        public string EmpID { get; set; } = string.Empty;

        /// <summary>Display name used in the voucher narration (Name - FName Designation).</summary>
        public string EmpDisplayName { get; set; } = string.Empty;

        // ── Clearance Details ─────────────────────────────────────────────────
        /// <summary>Date the repayment was received (stored in AmtCleared.DT).</summary>
        public DateTime ClearDate { get; set; } = DateTime.Today;

        /// <summary>Amount received from the employee (AmtCleared.AmtClrd).</summary>
        public decimal AmountReceived { get; set; }

        /// <summary>
        /// Optional new monthly deduction. When &gt; 0 and a latest advance exists,
        /// Advances.DAmount is updated for the latest EntryID.
        /// </summary>
        public decimal NewDeduction { get; set; }

        /// <summary>Description / narration (AmtCleared.Description).</summary>
        public string Description { get; set; } = string.Empty;

        // ── Voucher Posting ───────────────────────────────────────────────────
        /// <summary>When true, no voucher is posted; only the AmtCleared row is written.</summary>
        public bool DontPostVoucher { get; set; } = false;

        public ClearAdvanceReceiptMode ReceiptMode { get; set; } = ClearAdvanceReceiptMode.Cash;

        /// <summary>
        /// AccNo of the receiving (debit) account.
        /// Cash→Cash-in-Hand AccNo, OtherAccount→any transactional AccNo.
        /// </summary>
        public string ReceiptAccNo { get; set; } = string.Empty;

        /// <summary>
        /// EntryID of the employee's latest Type=1 advance; null when the
        /// employee has no long-term advance (skips the DAmount update).
        /// </summary>
        public long? LatestAdvanceEntryID { get; set; }
    }

    /// <summary>
    /// Read-only summary shown on-screen when an employee is selected.
    /// Fetched from Advances + AmtCleared tables.
    /// </summary>
    public class ClearAdvanceSummaryModel
    {
        /// <summary>
        /// Outstanding long-term loan balance:
        /// SUM(Advances.Amount, Type=1) − SUM(AmtCleared.AmtClrd).
        /// </summary>
        public decimal OutstandingBalance { get; set; }

        /// <summary>Monthly deduction (DAmount) from the employee's latest Type=1 advance.</summary>
        public decimal CurrentDeduction { get; set; }

        /// <summary>EntryID of the latest Type=1 advance; null when none exists.</summary>
        public long? LatestAdvanceEntryID { get; set; }
    }
}
