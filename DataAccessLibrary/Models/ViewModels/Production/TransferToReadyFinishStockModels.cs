using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ReadyFinishStockAvailableItem
    {
        public long VRD_EntryID { get; set; }       // VendRcvdDetail.EntryID
        public int VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayMaker => !string.IsNullOrEmpty(VendID1)
            ? $"{{ {VendID1} }} {VenderName}"
            : VenderName;

        public string RecieptID { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string Description { get; set; } = string.Empty; // Process
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string ItemSize { get; set; } = string.Empty;
        public string SizeUnit { get; set; } = string.Empty;
        public string FullArticle => $"{ItemCode} - {ItemName} {ItemSize} {SizeUnit}".Trim();

        public decimal RcvdQty { get; set; }
        public decimal Wastage { get; set; }
        public decimal LostQty { get; set; }
        public decimal IssQty { get; set; }
        public decimal ReWorkQty { get; set; }

        public decimal AvailQty => RcvdQty - (Wastage + LostQty) - IssQty - ReWorkQty;
        public decimal TransferQty { get; set; }

        public string OrderNo { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string Batch_No { get; set; } = string.Empty;
        public string Mill_Certificate_No { get; set; } = string.Empty;

        public bool IsSelected { get; set; }
    }

    public class ReadyFinishStockFilter
    {
        public int MakerId { get; set; } = 0;       // 0 = <All Makers>
        public string ArticleId { get; set; } = string.Empty; // empty = <All Articles>
        public string LotNo { get; set; } = string.Empty;
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=15d, 2=30d, 3=60d, 4=90d, 5=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
    }

    public class StoreLookupDto
    {
        public int EntryID { get; set; }
        public string StoreName { get; set; } = string.Empty;
    }

    public class ShelfLookupDto
    {
        public int EntryID { get; set; }
        public int Store_RefID { get; set; }
        public string RackNo { get; set; } = string.Empty;
        public string ShelfNo { get; set; } = string.Empty;
        public string DisplayName => $"{RackNo} {{{ShelfNo}}}";
    }

    public class TransferItemDto
    {
        public long VRD_EntryID { get; set; }
        public decimal TransferQty { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public string MillCertNo { get; set; } = string.Empty;
    }

    public class PostTransferRequest
    {
        public string ArticleId { get; set; } = string.Empty;
        public int SelectedStoreId { get; set; }
        public int SelectedShelfId { get; set; }
        public string ShelfLocationText { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public List<TransferItemDto> Items { get; set; } = new List<TransferItemDto>();
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class PostTransferResult
    {
        public bool Success { get; set; }
        public string RcvNo { get; set; } = string.Empty;
        public long TransferEntryID { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
