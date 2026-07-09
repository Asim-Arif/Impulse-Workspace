using System;

namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class ArticleShippedInvoiceModel
    {
        public string CustomInvoice { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public DateTime? GatePassDT { get; set; }
        public long Qty { get; set; }
        public double Price { get; set; }
        public double Amount => Qty * Price;
        public string OrderNo { get; set; } = string.Empty;
        public DateTime? OrderDT { get; set; }
        public DateTime? DeliveryDTItem { get; set; }
        
        public long Days 
        {
            get
            {
                if (OrderDT.HasValue && GatePassDT.HasValue)
                {
                    return (long)(GatePassDT.Value - OrderDT.Value).TotalDays;
                }
                return 0;
            }
        }
    }

    public class ArticleUnshippedOrderModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string Proformas { get; set; } = string.Empty;
        public long Qty { get; set; }
        public long ShippedQty { get; set; }
        public long BalanceQty => Qty - ShippedQty;
        public double Price { get; set; }
        public double Amount => BalanceQty * Price;
    }

    public class ArticleCustomerPriceModel
    {
        public long EntryID { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public double FOB { get; set; }
        public double ExWorks { get; set; }
        public double CnISea { get; set; }
        public double CnIAir { get; set; }
        public double CIFSea { get; set; }
        public double CIFAir { get; set; }
        public double CnFSea { get; set; }
        public double CnFAir { get; set; }

        // UI Helpers
        public bool IsEditing { get; set; }
        public string EditingColumn { get; set; } = string.Empty;
        public double EditPrice { get; set; }
    }

    public class ArticleQuotationModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string QuotationNo { get; set; } = string.Empty;
        public long Qty { get; set; }
        public double Price { get; set; }
        public double Amount => Qty * Price;
    }
}
