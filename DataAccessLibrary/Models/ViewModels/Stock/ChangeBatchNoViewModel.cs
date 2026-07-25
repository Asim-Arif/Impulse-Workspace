using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class ChangeBatchNoViewModel
    {
        public long EntryID { get; set; }
        
        public string RcvID { get; set; }
        public string VendorName { get; set; }
        public decimal QtyPlaced { get; set; }
        public string StoreName { get; set; }
        public string RackNo { get; set; }
        public string ShelfNo { get; set; }

        public string BatchNo { get; set; }
        public string OriginalBatchNo { get; set; }
        
        public string LotNo { get; set; }
        public string OriginalLotNo { get; set; }
        
        public string MillCertificateNo { get; set; }
        public string OriginalMillCertificateNo { get; set; }

        public bool IsModified => 
            (BatchNo ?? "") != (OriginalBatchNo ?? "") || 
            (LotNo ?? "") != (OriginalLotNo ?? "") || 
            (MillCertificateNo ?? "") != (OriginalMillCertificateNo ?? "");
    }
}
