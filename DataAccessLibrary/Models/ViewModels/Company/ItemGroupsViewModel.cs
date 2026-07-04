using System;

namespace DataAccessLibrary.Models.ViewModels.Company
{
    public class ItemGroupsViewModel
    {
        public int ID { get; set; }
        public string Description { get; set; } = string.Empty;
        public long GrpColor { get; set; }
        public int? MainGroup_RefID { get; set; }
        public string EmpID_Purchaser { get; set; } = string.Empty;
        public string PurchaserName { get; set; } = string.Empty;
    }

    public class PurchaserLookupModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }
}
