using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerPOListItem
    {
        public long EntryID { get; set; }
        public long DetailEntryID { get; set; }
        public int VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayMaker => !string.IsNullOrEmpty(VendID1) ? $"{{ {VendID1} }} {VenderName}" : VenderName;

        public string RecieptID { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string Description { get; set; } = string.Empty; // Process description

        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string TipSize { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string FullArticle => $"{ItemID} {ItemName} {TipSize} {ItemSize} {SizeUnit}".Trim();

        public decimal TotalValue { get; set; }
        public decimal TotalIssQty { get; set; }
        public bool Authorized { get; set; }
        public string AuthStatus => Authorized ? "Authorized" : "Un-Auth.";

        public string MakerPOReturnDTs { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string MasterPONo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string InternalRefNo { get; set; } = string.Empty;

        public string CatID { get; set; } = string.Empty;
        public string GroupID { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string CustCode { get; set; } = string.Empty;

        public bool? ReWorkLot { get; set; }
        public bool Closed { get; set; }
        public long? BookMarkEntryID { get; set; }
        public string? ComplaintItemID { get; set; }
    }

    public class MakerPOListFilter
    {
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=Last 15d, 2=Last 30d, 3=Last 60d, 4=Last 90d, 5=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;

        public int DueDaysMode { get; set; } = 0; // 0=None, 1=Within 3 Days, 2=Within 7 Days, 3=Custom
        public int DueDaysCustom { get; set; } = 0;

        public HashSet<int> MakerIds { get; set; } = new HashSet<int>();
        public HashSet<string> ItemCatIds { get; set; } = new HashSet<string>();
        public HashSet<int> ItemGroupIds { get; set; } = new HashSet<int>();
        public HashSet<int> ProcessIds { get; set; } = new HashSet<int>();
        public HashSet<string> CustomerCodes { get; set; } = new HashSet<string>();

        public string ItemId { get; set; } = "0";
        public string PurchaserEmpId { get; set; } = "0";

        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string MasterPONo { get; set; } = string.Empty;
        public string InActiveDays { get; set; } = string.Empty;

        public bool ShowMasterPOOnly { get; set; }
        public bool MasterPOOpen { get; set; }
        public bool RepairLots { get; set; }
        public bool RegularLotsOnly { get; set; }
        public bool Bookmarks { get; set; }
        public bool OpenLotsOnly { get; set; }
        public bool RepairIssuance { get; set; }
    }

    public class LookupItemInt
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class LookupItemString
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }
}
