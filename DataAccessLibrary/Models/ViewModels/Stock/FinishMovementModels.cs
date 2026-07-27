using System;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class FinishFromLocationModel
    {
        public int EntryID { get; set; }
        public int Shelf_RefID { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string Batch_No { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;
        public string StoreName { get; set; } = string.Empty;
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public double NetQty { get; set; }
        public string Remarks { get; set; } = string.Empty;

        public string DisplayText => $"Lot: {LotNo}, Batch No.: {Batch_No}, Location: {StoreName}[{RackNo}] [{ShelfNo}] ({NetQty:N0})";
    }

    public class FinishMovementRequest
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public FinishFromLocationModel SourceLocation { get; set; }
        public int TargetStoreRefId { get; set; }
        public string TargetRackNo { get; set; } = string.Empty;
        public int TargetShelfRefId { get; set; }
        public string TargetShelfNo { get; set; } = string.Empty;
        public string TargetStoreName { get; set; } = string.Empty;
        public double MoveQty { get; set; }
        public string UserName { get; set; } = "System";

        public string TargetLocationDisplay => $"{TargetRackNo} {{{TargetShelfNo}}}";
    }
}
