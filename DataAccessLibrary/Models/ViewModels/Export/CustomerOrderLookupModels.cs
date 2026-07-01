namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CompanyLookupModel
    {
        public int EntryID { get; set; }
        public string CompanyName { get; set; } = string.Empty;
    }

    public class CustomerLookupModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    public class CustomerOrderListItemModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public decimal OrderAmt { get; set; }
        public string Curr { get; set; } = string.Empty;
        public decimal TotalInvQty { get; set; }
        public decimal TotalOrderQty { get; set; }
        public DateTime DeliveryDT { get; set; }
        public int CompanyRefID { get; set; }
        public string CompanyName { get; set; } = string.Empty;
        public string? UnshippedOrderNo { get; set; }
        public decimal? TotalShippedQty { get; set; }
        public decimal? ExchRate { get; set; }
        public string InternalRefNo { get; set; } = string.Empty;
        public decimal TotalWeight { get; set; }
        public decimal? TotalShippedAmt { get; set; }
        public int Cancelled { get; set; }
        public int TotalArticles { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public int? ShippedItemCount { get; set; }
        public decimal? ShippedItemQty { get; set; }
        public int? PartiallyShippedItemCount { get; set; }
        public decimal? PartiallyShippedItemQty { get; set; }
        public decimal? TotalBalanceQty { get; set; }
        public string OrderType { get; set; } = string.Empty;
        public decimal? TotalPlannedQty { get; set; }
        public bool? OrderPlanApproved { get; set; }
        public bool? LateOrderAlerts { get; set; }

        // Helper derived fields for UI display
        public string OrderStatus { get; set; } = string.Empty;
        public string PurchasePlanStatus { get; set; } = string.Empty;
    }
}
