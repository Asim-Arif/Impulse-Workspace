using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class FinalizedLotItem
    {
        public long VRD_EntryID { get; set; }       // VendRcvdDetail.EntryID
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string DisplayMaker => !string.IsNullOrEmpty(VendID1)
            ? $"{{ {VendID1} }} {VenderName}"
            : VenderName;

        public string InternalRefNo { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public string Description { get; set; } = string.Empty; // Process
        public int ProcessID { get; set; }
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

        public decimal DispatchQty { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
    }

    public class DispatchListStagedItem
    {
        public Guid RowId { get; set; } = Guid.NewGuid();
        public long InnerEntryID { get; set; }      // DispatchListDetail_Inners.EntryID (if existing)
        public long VRD_EntryID { get; set; }
        public decimal Qty { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public int ProcessID { get; set; }
        public string ProcessDescription { get; set; } = string.Empty;
        public int CartonNo { get; set; } = 0;
        public string LotNo { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public bool IsExisting => InnerEntryID > 0;
        public bool IsSelected { get; set; }
    }

    public class DispatchListFilter
    {
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public int MakerId { get; set; } = 0;       // 0 = <All Makers>
        public string ArticleId { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public string LotNo { get; set; } = string.Empty;
        public int DateRangeIndex { get; set; } = 0; // 0=Today, 1=15d, 2=30d, 3=60d, 4=90d, 5=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today;
        public DateTime DtTo { get; set; } = DateTime.Today;
    }

    public class SaveDispatchListRequest
    {
        public long EntryID { get; set; }           // 0 = New, > 0 = Edit
        public string DispatchListNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public List<DispatchListStagedItem> Items { get; set; } = new List<DispatchListStagedItem>();
        public List<long> DeletedInnerIDs { get; set; } = new List<long>();
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
    }

    public class SaveDispatchListResult
    {
        public bool Success { get; set; }
        public long DispatchListEntryID { get; set; }
        public string DispatchListNo { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    public class LoadedDispatchListForEdit
    {
        public long EntryID { get; set; }
        public string DispatchListNo { get; set; } = string.Empty;
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public bool Finalyzed { get; set; }
        public List<DispatchListStagedItem> Items { get; set; } = new List<DispatchListStagedItem>();
    }
}
