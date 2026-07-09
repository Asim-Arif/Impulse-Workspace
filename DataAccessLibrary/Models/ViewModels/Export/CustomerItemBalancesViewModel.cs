using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerItemBalancesViewModel
    {
        public long ID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public DateTime DeliveryDT { get; set; }
        
        // This is Qty - ShippedQty
        public decimal BalQty { get; set; }
        public decimal Qty { get; set; }
        public decimal ShippedQty { get; set; }

        public int DeliveryStatus { get; set; } // 0=No Date, 1=Regular, 2=Urgent
        
        // Helps track the row's edited state
        public bool IsEditing { get; set; }
        public decimal EditBalQty { get; set; }
    }
}
