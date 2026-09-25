using System;

namespace DataAccessLibrary.Models.Setup
{
    public class MenuOptionModel
    {
        public string OptionID { get; set; } = string.Empty;
        public string OptionName { get; set; } = string.Empty;
        public string ModuleName { get; set; } = string.Empty;
        public string? FormName { get; set; }

        public bool IsSelected { get; set; }
    }
}
