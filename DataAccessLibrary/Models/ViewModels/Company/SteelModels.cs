using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class SteelTypeNodeModel
    {
        public int SteelID { get; set; }
        public string SteelType { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int StripeMinW { get; set; }
        public int StripeMaxW { get; set; }
        public bool IsExpanded { get; set; } = true;
        public List<SteelGageNodeModel> Gages { get; set; } = new();
        public List<SteelCompModel> Compositions { get; set; } = new();
    }

    public class SteelGageNodeModel
    {
        public string GageID { get; set; } = string.Empty;
        public int SteelID { get; set; }
        public string Gage { get; set; } = string.Empty;
        public string GageUnit { get; set; } = string.Empty;
    }

    public class SteelCompModel
    {
        public int CompID { get; set; }
        public string Symbol { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        
        // These will map to LVal and UVal in SteelTypeComps
        public float? MinValue { get; set; }
        public float? MaxValue { get; set; }
    }
}
