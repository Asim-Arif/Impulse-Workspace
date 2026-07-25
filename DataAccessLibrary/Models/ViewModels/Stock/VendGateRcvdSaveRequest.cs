using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendGateRcvdSaveRequest
    {
        public DateTime RcvDate { get; set; } = DateTime.Today;
        public string VendID { get; set; }
        public string OrderNo { get; set; }
        public string Description { get; set; }
        
        public List<VendGateRcvdItemViewModel> ReceivedItems { get; set; } = new List<VendGateRcvdItemViewModel>();
    }
}
