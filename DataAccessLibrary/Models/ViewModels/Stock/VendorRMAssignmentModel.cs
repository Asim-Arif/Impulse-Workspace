namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class VendorRMAssignmentModel
    {
        public int EntryID { get; set; }
        public string VendID { get; set; } = string.Empty;
        public int RMID { get; set; }
        public string RMID1 { get; set; } = string.Empty;
        public string RMName { get; set; } = string.Empty;
        public string GroupName { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
    }
}
