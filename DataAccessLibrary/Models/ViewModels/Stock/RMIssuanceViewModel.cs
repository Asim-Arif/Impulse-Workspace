using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Stock
{
    public class RMIssuanceSaveRequest
    {
        public string IssNo { get; set; } // Out: generated inside the save transaction
        public DateTime DT { get; set; } = DateTime.Today;
        public int MakerVendID { get; set; }
        public int VI_RefID { get; set; }
        public string IssdToEmpID { get; set; } // Combo display text "{EmpID} Name" (legacy format)
        public string AuthorizedByEmpID { get; set; } // Combo display text "{EmpID} Name" (legacy format)
        public string Remarks { get; set; }

        public List<RMIssuanceDetailRow> Rows { get; set; } = new List<RMIssuanceDetailRow>();

        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    public class RMIssuanceDetailRow
    {
        public int RMID { get; set; } // RM.RMID (int key)
        public string RMID1 { get; set; }
        public string RMName { get; set; }
        public double? RMSize { get; set; }
        public string SizeUnit { get; set; }
        public string Unit { get; set; }
        public string Store { get; set; } // "StoreName [RackNo] [ShelfNo]"
        public string Location { get; set; } // Full location display (incl. qty)
        public int Shelf_RefID { get; set; }
        public double Qty { get; set; }
        public double Rate { get; set; }
        public string BatchNo { get; set; }
        public string LotNo { get; set; }
        public double TotalQtyInStock { get; set; } // Snapshot of RM.QtyInStock at add time
    }

    public class MakerLookupModel
    {
        public int VendID { get; set; }
        public string VenderName { get; set; }
    }

    public class EmployeeLookupModel
    {
        public string EmpID { get; set; }
        public string Name { get; set; }
        public string Display => "{" + EmpID + "} " + Name;
    }

    public class IssuanceReceiptLookupModel
    {
        public int EntryID { get; set; }
        public string RecieptID { get; set; }
        public DateTime? DT { get; set; }
        public string Display => RecieptID;
    }

    public class IssuanceInfoModel
    {
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string Process { get; set; } // Processes.Description
        public double TotalIssQty { get; set; }
    }

    public class RMLookupModel
    {
        public int RMID { get; set; }
        public string RMID1 { get; set; }
        public string RMName { get; set; }
        public double? RMSize { get; set; }
        public string SizeUnit { get; set; }
        public int? GroupID { get; set; }
        public string Display => RMName + " {" + RMID1 + "} " + (RMSize.HasValue ? RMSize.Value.ToString() : "") + " " + SizeUnit;
    }

    public class RMLocationModel
    {
        public int Shelf_RefID { get; set; }
        public string Display { get; set; } // "Store [Rack] [Shelf] (qty)" / "Batch:… Lot:… Store…"
        public string Store { get; set; } // "StoreName [RackNo] [ShelfNo]"
        public double ShelfQty { get; set; }
        public double SheetsQty { get; set; }
        public string BatchNo { get; set; }
        public string LotNo { get; set; }
    }

    public class RMDetailsModel
    {
        public double? MinLevel { get; set; }
        public double? Rate { get; set; }
        public string Unit { get; set; }
        public double? QtyInStock { get; set; }
        public bool IsBatchwise { get; set; }
        public bool IsTwoUnits { get; set; }
    }

    public class RMHistoryRowModel
    {
        public string RMID1 { get; set; }
        public string RMName { get; set; }
        public double Qty { get; set; }
    }
}
