using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ProductionItemListItem
    {
        public DateTime? StartDT { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int Qty { get; set; }
        public string Description { get; set; } = string.Empty;   // Current process name
        public int LastProcessDays { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string? POIRemarks { get; set; }
        public int TLRunning { get; set; }
        public int TLUnShippedQty { get; set; }
        public int R1 { get; set; }
        public int R2 { get; set; }
        public int PackStock { get; set; }
        public string? UnShipped_OrderNo { get; set; }
        public int OpenPOQty { get; set; }
        public double? Price { get; set; }
        public double? PriceForCost { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Hub_Name { get; set; } = string.Empty;
        public int EntryType { get; set; }
        public int ReWorkLot { get; set; }
        public int ProcessID { get; set; }
        public int VendID { get; set; }
        public int? GroupID { get; set; }
        public string? Country { get; set; }

        // ── Computed display helpers ──────────────────────────────────
        public string DisplayRemarks => !string.IsNullOrEmpty(POIRemarks) ? POIRemarks : Remarks;
        public string ShippingStatus => string.IsNullOrEmpty(UnShipped_OrderNo) ? "Shipped" : "Un-Shipped";
        public double? DisplayPrice => CustCode?.ToUpperInvariant() == "STOCK" ? PriceForCost : Price;
        public int SemiQty => R1 + R2;
        public string TLRunningDisplay => $"{TLRunning}/{TLUnShippedQty}";

        // ── Inline editing & row state (not persisted) ────────────────
        public int SNo { get; set; }
        public bool IsEditingRemarks { get; set; }
        public string EditingRemarksValue { get; set; } = string.Empty;
    }

    public class ProductionItemListFilters
    {
        public string CustCode { get; set; } = "0";
        public string Country { get; set; } = "0";
        /// <summary>0=All, 1=+15d, 2=+30d, 3=+45d, 4=Custom</summary>
        public int DateRangeMode { get; set; } = 0;
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
        public string ItemGroupId { get; set; } = "0";
        public string OrderNo { get; set; } = "0";
        public string ProcessId { get; set; } = "0";
        public string MakerId { get; set; } = "0";
        public string ItemCode { get; set; } = "0";
        public string Hub { get; set; } = "0";
        public string LotNo { get; set; } = "";
        public bool ShowIssuance { get; set; } = true;
        public bool ShowReceiving { get; set; } = true;
        public bool ShowLotsOnly { get; set; } = true;
        public bool ShowNotYetStarted { get; set; } = true;
        public bool ShowRegularLots { get; set; } = true;
        public bool ShowRepairLots { get; set; } = true;
    }

    public class ProductionItemListLookups
    {
        public List<LookupItemString> Customers { get; set; } = new();
        public List<LookupItemString> Countries { get; set; } = new();
        public List<LookupItemString> Orders { get; set; } = new();
        public List<LookupItemInt> ItemGroups { get; set; } = new();
        public List<LookupItemInt> Processes { get; set; } = new();
        public List<LookupItemInt> Makers { get; set; } = new();
        public List<LookupItemString> Items { get; set; } = new();
        public List<string> Hubs { get; set; } = new();
    }
}
