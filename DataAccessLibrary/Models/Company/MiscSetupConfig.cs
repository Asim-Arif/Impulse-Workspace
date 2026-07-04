using System.Collections.Generic;

namespace DataAccessLibrary.Models.Company
{
    public class MiscSetupConfig
    {
        public string TableName { get; }
        public string IdColumn { get; }
        public string DescColumn { get; }
        public string PageTitle { get; }

        public MiscSetupConfig(string tableName, string idColumn, string descColumn, string pageTitle)
        {
            TableName = tableName;
            IdColumn = idColumn;
            DescColumn = descColumn;
            PageTitle = pageTitle;
        }

        // Static registry of all known generic lookup setups
        public static readonly IReadOnlyDictionary<string, MiscSetupConfig> Configurations = new Dictionary<string, MiscSetupConfig>
        {
            // The key is the URL parameter, e.g. /company/misc-setup/AdditionalGroups
            { "AdditionalGroups", new MiscSetupConfig("ItemGroupsMain", "MainGroupID", "MainGroupName", "Additional Groups") },
            { "ItemTypes", new MiscSetupConfig("ItemTypes", "ID", "Description", "Item Types") },
            { "ItemFinishedQuality", new MiscSetupConfig("ItemsFinQuality", "ID", "Description", "Item Finished Quality") }
        };
    }
}
