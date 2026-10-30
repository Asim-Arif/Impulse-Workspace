using System.Collections.Generic;

namespace DataAccessLibrary.Models.QMS
{
    public class TrackBackInvoiceItemDto
    {
        public int RowIndex { get; set; }
        public string CompItemCode { get; set; } = string.Empty;
        public string? ItemName { get; set; }
        public decimal Qty { get; set; }
    }

    public class TrackBackLotDto
    {
        public string LotNo { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public int AttachmentCount { get; set; }
    }

    public class TrackBackMasterPODto
    {
        public string MasterPONo { get; set; } = string.Empty;
        public string? VenderName { get; set; }
    }

    public class TrackBackOrderDto
    {
        public string OrderNo { get; set; } = string.Empty;
        public decimal Qty { get; set; }
    }
}

