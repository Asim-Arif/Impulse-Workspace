using System;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class MakerAccountModel
    {
        public long VendID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string MakerNo { get; set; } = string.Empty;
        public string MakerName { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public decimal PreviousBalance { get; set; }
    }

    public class MakerMasterPoModel
    {
        public string MasterPONo { get; set; } = string.Empty;
        public decimal POValue { get; set; }
        public decimal LTTaken { get; set; }
        public decimal STTaken { get; set; }
        public decimal TotalTaken => LTTaken + STTaken;
        public decimal POLimit => (POValue / 2m) - TotalTaken;
    }

    public enum MakerLoanPaymentMode
    {
        Cash = 0,
        Bank = 1,
        OtherAccount = 2
    }

    public class MakerShortTermLoanDto
    {
        public string AccNo { get; set; } = string.Empty;
        public string MakerNo { get; set; } = string.Empty;
        public string MakerName { get; set; } = string.Empty;
        public DateTime IssueDate { get; set; } = DateTime.Today;
        public decimal Amount { get; set; }
        public string Description { get; set; } = string.Empty;
        public string? MasterPONo { get; set; }
        public decimal POValue { get; set; }
        public decimal LTTaken { get; set; }
        public decimal STTaken { get; set; }
        public MakerLoanPaymentMode PaymentMode { get; set; } = MakerLoanPaymentMode.Cash;
        public string PaymentAccNo { get; set; } = string.Empty;
        public string BankAccNo { get; set; } = string.Empty;
        public string BankId { get; set; } = string.Empty;
        public string ChequeNo { get; set; } = string.Empty;
        public string ChequeType { get; set; } = "Open";
        public DateTime ChequeDate { get; set; } = DateTime.Today;

        public bool DontPostVoucher { get; set; } = false;
    }
}
