using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    // NOTE: The receipt mode is shared with the long-term clearance page via
    // ClearAdvanceReceiptMode (Cash→CRV, OtherAccount→JV) — see ClearAdvanceModels.cs.

    /// <summary>
    /// One open short-term advance row (Advances where Type=0 and Cleared=0)
    /// shown in the checkbox table on the Clear Short Term Loan page.
    /// </summary>
    public class ShortTermAdvanceItem
    {
        /// <summary>Advances.EntryID (bigint).</summary>
        public long EntryID { get; set; }

        /// <summary>Advances.DT.</summary>
        public DateTime Date { get; set; }

        /// <summary>Advances.Amount (money).</summary>
        public decimal Amount { get; set; }

        /// <summary>Advances.Description.</summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>UI-only: whether the row is ticked for clearance.</summary>
        public bool IsSelected { get; set; }
    }

    /// <summary>
    /// DTO passed from the Blazor page to the service layer when saving a
    /// Short-Term Loan clearance (bulk settlement of open advances).
    /// </summary>
    public class ClearShortTermDto
    {
        // ── Employee ──────────────────────────────────────────────────────────
        public string EmpID { get; set; } = string.Empty;

        /// <summary>Display name used in the voucher narration (Name - FName Designation).</summary>
        public string EmpDisplayName { get; set; } = string.Empty;

        // ── Clearance Details ─────────────────────────────────────────────────
        /// <summary>Clearance date (stored in ShortTermAdvanceClearance.DT).</summary>
        public DateTime ClearDate { get; set; } = DateTime.Today;

        /// <summary>Remarks (ShortTermAdvanceClearance.Remarks — varchar(50)).</summary>
        public string Remarks { get; set; } = string.Empty;

        // ── Voucher Posting ───────────────────────────────────────────────────
        /// <summary>When true, no voucher is posted; only the clearance rows are written.</summary>
        public bool DontPostVoucher { get; set; } = false;

        public ClearAdvanceReceiptMode ReceiptMode { get; set; } = ClearAdvanceReceiptMode.Cash;

        /// <summary>
        /// AccNo of the receiving (debit) account.
        /// Cash→Cash-in-Hand AccNo, OtherAccount→any transactional AccNo.
        /// </summary>
        public string ReceiptAccNo { get; set; } = string.Empty;

        /// <summary>The ticked advance rows to clear (all-or-nothing per row).</summary>
        public List<ShortTermAdvanceItem> SelectedAdvances { get; set; } = new();
    }
}
