using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendorBillingListViewModel
    {
        public int EntryID { get; set; }
        public string VchrNo { get; set; }
        public int BillNo { get; set; }
        public DateTime DT { get; set; }
        public string AccTitle { get; set; }
        public decimal TotalBillAmount { get; set; }
        public decimal TotalDedAmount { get; set; }
        
        public decimal NetAmount => TotalBillAmount - TotalDedAmount;
    }
}
