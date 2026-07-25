namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class LocationwiseStatusViewModel
    {
        public int EntryID { get; set; }
        public int? Rcvd_RefID { get; set; }
        public int? Shelf_RefID { get; set; }
        public string StoreName { get; set; }
        public string ShelfNo { get; set; }
        public double AvailableQty { get; set; }
        public string BatchNo { get; set; }
        public string LotNo { get; set; }
        
        public string DisplayText => $"{StoreName} [{ShelfNo}] ({AvailableQty:N2}) {(string.IsNullOrEmpty(BatchNo) ? "" : $"Batch: {BatchNo}")}";
    }
}
