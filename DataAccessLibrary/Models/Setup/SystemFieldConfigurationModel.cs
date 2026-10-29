namespace DataAccessLibrary.Models.Setup
{
    public class SystemFieldConfigurationModel
    {
        public int SettingID { get; set; }
        public string Module { get; set; } = string.Empty;
        public string FormKey { get; set; } = string.Empty;
        public string FieldKey { get; set; } = string.Empty;
        public bool IsRequired { get; set; } = true;
        public bool IsVisible { get; set; } = true;
        public string? Description { get; set; }
    }
}

