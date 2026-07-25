using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendGateRcvdListViewModel
    {
        public int SNo { get; set; }
        public string PORefNo { get; set; }
        public string RcvID { get; set; }
        public DateTime RcvDate { get; set; }
        public string VenderName { get; set; }
        public decimal TotalQty { get; set; }
        public decimal TotalQtyRejected { get; set; }
    }
}
