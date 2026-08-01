using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class MakerLongLoanLedgerEntryModel
    {
        public long EntryID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string Description { get; set; } = string.Empty;
        public string? Remarks { get; set; }
        public decimal Amount { get; set; }
        public decimal DAmount { get; set; }
        public string AccVoucherNo { get; set; } = string.Empty;
        public decimal RunningBalance { get; set; }

        public string FullDescription => !string.IsNullOrWhiteSpace(Remarks)
            ? $"{Description} ({Remarks})"
            : Description;
    }

    public class MakerLongLoanLedgerResultModel
    {
        public string AccNo { get; set; } = string.Empty;
        public string MakerName { get; set; } = string.Empty;
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal TotalAdvances { get; set; }
        public decimal TotalDeductions { get; set; }
        public decimal ClosingBalance => OpeningBalance + TotalAdvances - TotalDeductions;
        public List<MakerLongLoanLedgerEntryModel> Entries { get; set; } = new();
    }
}
