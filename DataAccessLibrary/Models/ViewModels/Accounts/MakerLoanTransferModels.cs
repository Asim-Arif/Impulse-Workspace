using System;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class MakerLoanTransferDto
    {
        public string AccNo { get; set; } = string.Empty;
        public DateTime TransferDate { get; set; } = DateTime.Today;
        public decimal LTToClear { get; set; }
        public decimal STToClear { get; set; }
        public decimal NewLTDeductionAmount { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }
}
