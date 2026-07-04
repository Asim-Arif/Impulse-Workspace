using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class StoreNodeModel
    {
        public int StoreID { get; set; }
        public string StoreName { get; set; } = string.Empty;
        public bool IsExpanded { get; set; } = true;
        public List<RackNodeModel> Racks { get; set; } = new();
    }

    public class RackNodeModel
    {
        public int RackID { get; set; }
        public int StoreID { get; set; }
        public string RackNo { get; set; } = string.Empty;
        public bool IsExpanded { get; set; } = false;
        public List<BinNodeModel> Bins { get; set; } = new();
    }

    public class BinNodeModel
    {
        public int BinID { get; set; }
        public int RackID { get; set; }
        public string ShelfNo { get; set; } = string.Empty;
    }
}
