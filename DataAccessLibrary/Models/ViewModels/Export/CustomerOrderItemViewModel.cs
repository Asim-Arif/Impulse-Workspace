using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CustomerOrderItemViewModel
    {
        public int ID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public decimal Qty { get; set; }
        public decimal InvQty { get; set; }
        public decimal CustomPrice { get; set; }
        public int SortNo { get; set; }
        public DateTime DeliveryDT { get; set; } = new DateTime(1900, 1, 1);
        public string Stamps { get; set; } = string.Empty;
        public string Quality { get; set; } = string.Empty;
        public int DeliveryStatus { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public decimal Weight { get; set; }
        public int Item_Finishing_Type { get; set; }
        public string IW_OrderNo { get; set; } = string.Empty;
        public string IW_BatchNo { get; set; } = string.Empty;

        // Joined view properties
        public string ItemName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Unit { get; set; } = "Pcs";
        public string Item_Finishing_Type_Text { get; set; } = "Finished";
    }
}
