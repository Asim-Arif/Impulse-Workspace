using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendOrderDetailViewModel
    {
        public int EntryID { get; set; }
        public string OrderNo { get; set; }
        public string RMID { get; set; }
        public float? Rate { get; set; }
        public float? Qty { get; set; }
        public int? ForRMID { get; set; }
        public bool? CCItem { get; set; }
        public string ItemRemarks { get; set; }

        // Additional properties for UI display (joined fields)
        public string RMName { get; set; }
        public string RMType { get; set; }
        public string Unit { get; set; }
        public float Amount => (Rate ?? 0) * (Qty ?? 0);
    }
}
