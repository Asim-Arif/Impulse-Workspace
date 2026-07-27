using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class FinishStockReceivingMasterModel
    {
        public string RcvNo { get; set; } = string.Empty;
        public DateTime DT { get; set; } = DateTime.Today;
        public string RcvdBy { get; set; } = string.Empty; // Counted By / Dept
        public string RcvdFrom { get; set; } = string.Empty; // Received From (Attention)
        public string OrderNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = "Stock";
        public string UserName { get; set; } = "System";
    }

    public class FinishStockReceivingItemModel
    {
        public int LineNo { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public int Store_RefID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public int Shelf_RefID { get; set; }
        public string LocationDisplay { get; set; } = string.Empty; // Store [Rack] [Shelf]
        public string LotNo_Manual { get; set; } = string.Empty;
        public string BatchNo_Manual { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;
        public double Qty { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }
}
