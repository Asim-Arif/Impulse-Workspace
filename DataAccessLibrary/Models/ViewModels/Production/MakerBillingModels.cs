using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerBillingFilter
    {
        public int VendID { get; set; } = 0;
        public string DebitHeadAccNo { get; set; } = string.Empty;
        public string CatIDs { get; set; } = "0";
        public string MasterPONo { get; set; } = string.Empty;
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=7d, 2=15d, 3=30d, 4=60d, 5=90d, 6=1y, 7=5y, 8=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
        public DateTime PostingDate { get; set; } = DateTime.Today;
        public bool RejectionRelaxation { get; set; } = false;
        public bool ShowPicture { get; set; } = false;
    }

    public class MakerBillingLotItem
    {
        public long EntryID { get; set; }          // VendRcvdDetail.EntryID (VRD_EntryID)
        public long VI_EntryID { get; set; }       // VendIssd.EntryID (Tag)
        public long IssDetail_RefID { get; set; }  // VendIssdDetail.EntryID
        public string RcvReceiptID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string ItemPic { get; set; } = string.Empty;
        public decimal IssQty { get; set; }
        public decimal Rate { get; set; }
        public decimal IssValue => Math.Round(IssQty * Rate, 2);
        public decimal TempAdvance { get; set; }
        public decimal RepairAmt { get; set; }
        public decimal RcvdValue { get; set; }
        public decimal RcvdQty { get; set; }
        public decimal UnAuthValue { get; set; }
        public decimal BalValue { get; set; }
        public decimal PercentRcvd { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string Process { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public bool Audited { get; set; }
        public string InternalRefNo { get; set; } = string.Empty;
        public bool IsSelected { get; set; }

        public decimal Value => RcvdValue != 0 ? RcvdValue : IssValue;
        public decimal Qty => RcvdQty != 0 ? RcvdQty : IssQty;
    }

    public class MakerRepairItem
    {
        public long EntryID { get; set; }
        public string RecieptID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public decimal Rate { get; set; }
        public decimal RepairAmt { get; set; }
        public decimal WastageAmt { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string Process { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
    }

    public class MakerMaterialDedItem
    {
        public long EntryID { get; set; }          // RawMaterialIssuanceDetail.EntryID
        public long MasterEntryID { get; set; }    // RawMaterialIssuance.EntryID
        public string IssNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string RMName { get; set; } = string.Empty;
        public string RawMaterialDisplay => $"({RMID1}) {RMName}";
        public string ItemCode { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public decimal Qty { get; set; }
        public decimal Amount => Math.Round(Qty * Price);
        public bool IsSelected { get; set; }
    }

    public class MakerManualDedItem
    {
        public Guid RowId { get; set; } = Guid.NewGuid();
        public string Description { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    public class MakerHiddenLotItem
    {
        public long EntryID { get; set; }          // MakerBilling_Receivings_Hidden.EntryID
        public long VRD_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Process { get; set; } = string.Empty;
    }

    public class MakerHiddenRMItem
    {
        public long EntryID { get; set; }
        public long MasterEntryID { get; set; }
        public string IssNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string RMName { get; set; } = string.Empty;
        public string RawMaterialDisplay => $"({RMID1}) {RMName}";
        public string ItemCode { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public decimal Qty { get; set; }
        public decimal Amount => Math.Round(Qty * Price);
    }

    public class MakerBillingLookups
    {
        public List<LookupItemInt> Makers { get; set; } = new List<LookupItemInt>();
        public List<LookupItemInt> Categories { get; set; } = new List<LookupItemInt>();
        public List<LookupItemString> DebitHeads { get; set; } = new List<LookupItemString>();
    }

    public class PostMakerBillRequest
    {
        public int VendID { get; set; }
        public string DebitHeadAccNo { get; set; } = string.Empty;
        public DateTime PostingDate { get; set; } = DateTime.Today;
        public DateTime DtFrom { get; set; }
        public DateTime DtTo { get; set; }
        public bool RejectionRelaxation { get; set; }

        public decimal ShortTermDeduction { get; set; }
        public decimal LongTermDeduction { get; set; }
        public decimal RepairAmount { get; set; }
        public decimal MaterialDeductionAmount { get; set; }
        public decimal ManualDeductionAmount { get; set; }

        public List<MakerBillingLotItem> SelectedLots { get; set; } = new List<MakerBillingLotItem>();
        public List<MakerMaterialDedItem> SelectedMaterialDeductions { get; set; } = new List<MakerMaterialDedItem>();
        public List<MakerManualDedItem> ManualDeductions { get; set; } = new List<MakerManualDedItem>();

        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class PostMakerBillResult
    {
        public bool Success { get; set; }
        public long BillEntryID { get; set; }
        public long BillNo { get; set; }
        public string VoucherNo { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }
}
