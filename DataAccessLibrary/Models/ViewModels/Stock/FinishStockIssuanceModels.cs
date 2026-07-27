using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class FinishStockIssuanceMasterModel
    {
        public string IssNo { get; set; } = string.Empty;
        public DateTime DT { get; set; } = DateTime.Today;
        public string Department { get; set; } = string.Empty;
        public string AttnPerson { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = "Stock";
        public string UserName { get; set; } = "System";
    }

    public class FinishStockIssuanceItemModel
    {
        public int LineNo { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public int Store_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public int Shelf_RefID { get; set; }
        public string LocationDisplay { get; set; } = string.Empty;
        public int RISD_EntryID { get; set; } // EntryID from VItemsWithLotWiseStock / RcvItemsSimpleDetail
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public double AvailableQty { get; set; }
        public double Qty { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }

    public class CustomerLookupModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string DisplayText => string.IsNullOrWhiteSpace(Country) ? CustCode : $"{CustCode} ({Country})";
    }

    public class LocationLotStockModel
    {
        public int EntryID { get; set; } // RISD EntryID
        public string ItemID { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string Batch_No { get; set; } = string.Empty;
        public double NetQty { get; set; }
        public int Shelf_RefID { get; set; }
        public int Store_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;

        public string DisplayText => $"Lot: {LotNo} | Batch: {Batch_No} | Location: {StoreName} [{RackNo}] [{ShelfNo}] ({NetQty:N0})";
    }
}
