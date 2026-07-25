namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMMovementSaveRequest
    {
        public int From_MLS_RefID { get; set; }
        public int To_Shelf_RefID { get; set; }
        public double QtyMoved { get; set; }
        public int? Rcvd_RefID { get; set; }
        public string BatchNo { get; set; }
        public string LotNo { get; set; }
    }
}
