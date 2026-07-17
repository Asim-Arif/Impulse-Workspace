namespace DataAccessLibrary.Models.ViewModels.Export
{
    /// <summary>Item Group lookup for Export Performance Report (ItemGroups: ID, Description)</summary>
    public class ExportItemGroupLookupModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>Item Category lookup for Export Performance Report (ItemCatagories: CatID, Description)</summary>
    public class ExportItemCategoryLookupModel
    {
        public int CatID { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>Additional / Main Group lookup (ItemGroupsMain: MainGroupID, MainGroupName)</summary>
    public class ExportMainGroupLookupModel
    {
        public int MainGroupID { get; set; }
        public string MainGroupName { get; set; } = string.Empty;
    }
}
