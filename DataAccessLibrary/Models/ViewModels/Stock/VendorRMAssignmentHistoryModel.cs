using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendorRMAssignmentHistoryModel
    {
        public int EntryID { get; set; }
        public DateTime DTEntry { get; set; }
        public string UserName { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }
}
