using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class OrderItemListViewModel
    {
        public int ID { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string CompItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public DateTime DeliveryDT { get; set; }
        public int DeliveryStatus { get; set; }
        public int Qty { get; set; }
        public int ShippedQty { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public int? GroupID { get; set; }
        public string DispatchListNos { get; set; } = string.Empty;
        public int? InnerQty { get; set; }

        // Computed properties
        public int BalanceQty => Qty - ShippedQty;
        public int? DaysOverdue => DeliveryStatus == 0 ? (int?)null : (int)(DateTime.Today - DeliveryDT).TotalDays;
        public string DisplayDeliveryDT => DeliveryStatus == 0 ? "-" : DeliveryDT.ToString("dd-MMM-yyyy");
        public string DisplayRemarks => !string.IsNullOrEmpty(Remarks) ? Remarks : (DeliveryStatus == 2 ? "Urgent" : string.Empty);
    }

    public class OrderItemGroupLookupModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class OrderNoLookupModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;
        public string DisplayText => $"{OrderNo} {{{InternalRefNo}}}";
    }

    public class DeliveryEditModel
    {
        public int ID { get; set; }
        public DateTime NewDeliveryDT { get; set; } = DateTime.Today;
        public string Remarks { get; set; } = string.Empty;
        public int PreviousDeliveryStatus { get; set; }
        public DateTime? PreviousDeliveryDT { get; set; }
    }
}
