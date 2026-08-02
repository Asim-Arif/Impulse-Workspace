using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class AvailableLotIssuanceItemModel
    {
        public long VendIssdDetailEntryID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public byte[]? ItemPic { get; set; }
        public string? ItemPicBase64 => ItemPic != null && ItemPic.Length > 0 ? $"data:image/jpeg;base64,{Convert.ToBase64String(ItemPic)}" : null;
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public decimal AvailableQty { get; set; }
        public decimal IssuanceQty { get; set; }
        public decimal Rate { get; set; }
        public string Unit { get; set; } = "Pcs";
        public int TargetProcessID { get; set; }
        public string TargetProcessName { get; set; } = string.Empty;
        public int ReturnProcessID { get; set; }
        public string ReturnProcessName { get; set; } = string.Empty;
        public int Priority { get; set; }
        public int ReWorkLot { get; set; }
        public int RepairType { get; set; }
        public int ReqAuth { get; set; }
        public string CountedBy { get; set; } = string.Empty;
    }

    public class LotIssuanceLookupResultModel
    {
        public bool IsFound { get; set; }
        public bool AlreadyIssued { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<AvailableLotIssuanceItemModel> Items { get; set; } = new List<AvailableLotIssuanceItemModel>();
    }

    public class CreateLotIssuanceHeaderModel
    {
        public long VendID { get; set; }
        public int ProcessID { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public string MasterPONo { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public bool SampleProvided { get; set; }
        public bool DrawingProvided { get; set; }
        public bool ForgingProvided { get; set; }
        public bool SteelProvided { get; set; }
        public string IssEmpID { get; set; } = string.Empty;
        public string CountedBy { get; set; } = string.Empty;
    }

    public class CreateLotIssuanceLineModel
    {
        public long IssueRefID { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string OrderNo { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public decimal IssQty { get; set; }
        public string LotNo { get; set; } = string.Empty;
        public string BatchNo { get; set; } = string.Empty;
        public int RcvProcessID { get; set; }
        public DateTime ReturnDT { get; set; } = DateTime.Today;
        public int Priority { get; set; }
        public int ReWorkLot { get; set; }
        public int RepairType { get; set; }
        public string CountedBy { get; set; } = string.Empty;
    }
}
