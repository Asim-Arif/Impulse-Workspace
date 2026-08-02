using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class MakerLookupModel
    {
        public long VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayText => string.IsNullOrWhiteSpace(VendID1) ? VenderName : $"[{VendID1}] {VenderName}";
    }

    public class ProcessLookupModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public int SNO { get; set; }
    }

    public class UnassignedItemModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string? ItemSize { get; set; }
        public string? SizeUnit { get; set; }
        public string? FinQuality { get; set; }

        public string DisplayText
        {
            get
            {
                string details = string.Join(" ", new[] { ItemSize, SizeUnit, FinQuality }.Where(s => !string.IsNullOrWhiteSpace(s)));
                return string.IsNullOrWhiteSpace(details)
                    ? $"[{ItemID}] {ItemName}"
                    : $"[{ItemID}] {ItemName} ({details})";
            }
        }
    }

    public class AssignedMakerItemModel
    {
        public long EntryID { get; set; }
        public long VendID { get; set; }
        public int ProcessID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty; // Group
        public decimal Rate { get; set; }
        public decimal RateDozen => Rate * 12m;
        public string AssignedUnit { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public string MakerDescription { get; set; } = string.Empty;

        // Editing state helper for inline edit modal/drawer
        public bool IsEditing { get; set; }
        public decimal EditRate { get; set; }
        public string EditRemarks { get; set; } = string.Empty;
    }

    public class ItemRevisionHistoryModel
    {
        public long EntryID { get; set; }
        public long VAI_RefID { get; set; }
        public decimal Rate { get; set; }
        public decimal RateDozen => Rate * 12m;
        public string UserName { get; set; } = string.Empty;
        public DateTime DTEntry { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }
}
