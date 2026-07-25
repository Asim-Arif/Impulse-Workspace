namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class LotListViewModel
    {
        public int EntryID { get; set; }
        public string LotNo { get; set; }
        public string ItemID { get; set; }
        public string Batch_No { get; set; }
        public string Mill_Certificate_No { get; set; }
        public string Lot_Remarks { get; set; }
        
        public bool IsModified { get; set; }
    }
}
