namespace DataAccessLibrary.Models.ViewModels.Export
{
    public class CompanyLookupModel
    {
        public int EntryID { get; set; }
        public string CompanyName { get; set; } = string.Empty;
    }

    public class CustomerLookupModel
    {
        public string CustCode { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }
}
