using System;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class MakerLoanBalanceModel
    {
        public string AccNo { get; set; } = string.Empty;
        public decimal LTBalance { get; set; }
        public decimal LTDeductionAmount { get; set; }
        public decimal STBalance { get; set; }
    }

    public class MakerLoanClearanceDto
    {
        public string AccNo { get; set; } = string.Empty;
        public DateTime ClearanceDate { get; set; } = DateTime.Today;
        public decimal LTToClear { get; set; }
        public decimal STToClear { get; set; }
        public decimal NewLTDeductionAmount { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }
}
