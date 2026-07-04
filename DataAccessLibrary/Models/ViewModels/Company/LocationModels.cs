using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class CountryNodeModel
    {
        public string CountryName { get; set; } = string.Empty;
        public string OriginalCountryName { get; set; } = string.Empty; // Used for updates since it's the PK
        public bool IsExpanded { get; set; } = true;
        public List<CityNodeModel> Cities { get; set; } = new();
    }

    public class CityNodeModel
    {
        public int CityID { get; set; }
        public string CountryName { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public bool IsExpanded { get; set; } = false;
        public List<PortNodeModel> Ports { get; set; } = new();
    }

    public class PortNodeModel
    {
        public int PortID { get; set; }
        public int CityID { get; set; }
        public string Port { get; set; } = string.Empty;
    }
}
