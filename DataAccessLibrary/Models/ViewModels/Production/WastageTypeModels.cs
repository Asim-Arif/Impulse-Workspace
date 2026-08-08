using System;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class WastageTypeModel
    {
        public int EntryID { get; set; }
        public string WastageName { get; set; } = string.Empty;
        public int WastageType { get; set; } = 0; // 0: Store, 1: Vendor, 2: Originator, 3: Ask
        public bool Closed { get; set; } = false;

        public string WastageTypeTitle => WastageType switch
        {
            0 => "Charge to Store",
            1 => "Charge to Vendor / Maker",
            2 => "Charge to Originator",
            3 => "Charge to Ask",
            _ => "Unknown"
        };

        public string BadgeClass => WastageType switch
        {
            0 => "bg-secondary",
            1 => "bg-danger",
            2 => "bg-warning text-dark",
            3 => "bg-info text-dark",
            _ => "bg-light text-dark"
        };
    }

    public class WastageTypeDto
    {
        public int EntryID { get; set; }
        public string WastageName { get; set; } = string.Empty;
        public int WastageType { get; set; } = 0;
        public bool Closed { get; set; } = false;
    }
}
