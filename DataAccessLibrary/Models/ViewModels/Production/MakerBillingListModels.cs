using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerBillingListItem
    {
        public long EntryID { get; set; }
        public int VendID { get; set; }
        public string VenderName { get; set; } = string.Empty;
        public string VendID1 { get; set; } = string.Empty;
        public long? BillNo { get; set; }
        public string? VchrNo { get; set; }
        public DateTime DT { get; set; }
        public decimal TotalQty { get; set; }
        public decimal TotalValue { get; set; }
        public decimal NetPayable { get; set; }
        public decimal TotalPaidAmount { get; set; }
        public decimal Balance => NetPayable - TotalPaidAmount;
        public decimal STDeduction { get; set; }
        public decimal LTDeduction { get; set; }
        public bool Posted { get; set; }
        public string? Payment_VchrNo { get; set; }
    }

    public class MakerBillingListFilter
    {
        public int VendID { get; set; } = 0;
        public DateTime DtFrom { get; set; } = DateTime.Today.AddDays(-30);
        public DateTime DtTo { get; set; } = DateTime.Today;
        public int DateRangeIndex { get; set; } = 3; // Default 30 days
        public string SearchText { get; set; } = string.Empty;
    }

    public class MakerBillingListSummary
    {
        public int TotalBills { get; set; }
        public decimal TotalQty { get; set; }
        public decimal TotalValue { get; set; }
        public decimal TotalPaid { get; set; }
        public decimal TotalBalance { get; set; }
    }

    public class MakerBillingPaymentVoucherRequest
    {
        public long MPB_RefID { get; set; }
        public int VendID { get; set; }
        public string MakerName { get; set; } = string.Empty;
        public long? BillNo { get; set; }
        public string? BillVoucherNo { get; set; }
        public string DebitAccountNo { get; set; } = string.Empty;
        public string CreditAccountNo { get; set; } = string.Empty;
        public string PaymentMode { get; set; } = "Cash"; // Cash (CPV) or Bank (BPV)
        public decimal Amount { get; set; }
        public DateTime PaymentDate { get; set; } = DateTime.Today;
        public string ChequeNo { get; set; } = string.Empty;
        public DateTime ChequeDate { get; set; } = DateTime.Today;
        public string Remarks { get; set; } = string.Empty;
    }
}
