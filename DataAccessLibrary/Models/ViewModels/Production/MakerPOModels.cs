using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class ProcessPOLookupModel
    {
        public int ProcessID { get; set; }
        public string Description { get; set; } = string.Empty;
        public int SNO { get; set; }
    }

    public class MakerPOLookupModel
    {
        public long VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public decimal MakerCapacity { get; set; }
        public decimal Balance { get; set; }
        public string DisplayText => string.IsNullOrWhiteSpace(VendID1) 
            ? $"{VenderName} (Cap: {MakerCapacity}, Bal: {Balance})"
            : $"[{VendID1}] {VenderName} (Cap: {MakerCapacity}, Bal: {Balance})";
    }

    public class AssignedItemLookupModel
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public string Unit { get; set; } = "Pcs";
        public string DisplayText => $"[{ItemID}] {ItemName}";
    }

    public class SteelTypeLookupModel
    {
        public int SteelID { get; set; }
        public string SteelType { get; set; } = string.Empty;
    }

    public class EmployeeLookupModel
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
        public string DisplayText => $"[{EmpID}] {Name} ({Designation})";
    }

    public class IssuableOrderItemModel
    {
        public int ProcessID { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemID { get; set; } = string.Empty;
        public decimal OrderQty { get; set; }
        public decimal IssQty { get; set; } // SP IssQty for subsequent processes
        public int InitialProcessID { get; set; }
        public string CompItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal TotIssdProc { get; set; }
        public decimal AdjustedQty { get; set; }
        public string InternalRefNo { get; set; } = string.Empty;
        public bool IsSelected { get; set; }

        public decimal CalculatedIssQty
        {
            get
            {
                if (ProcessID == InitialProcessID)
                {
                    decimal qty = OrderQty - TotIssdProc - AdjustedQty;
                    return qty > 0 ? qty : 0;
                }
                return IssQty > 0 ? IssQty : (OrderQty - TotIssdProc);
            }
        }
    }

    public class StagedPOItemModel
    {
        public int LineNo { get; set; }
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal IssQty { get; set; }
        public decimal Rate { get; set; }
        public decimal RateDozen => Rate * 12m;
        public decimal Value => IssQty * Rate;
        public DateTime ReturnDT { get; set; }
        public string BatchNo { get; set; } = string.Empty;
        public int ReturnProcessID { get; set; }
        public string ReturnProcessName { get; set; } = string.Empty;
        public string AssignedUnit { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public string IssEmpID { get; set; } = string.Empty;
        public string CountedBy { get; set; } = string.Empty;
    }

    public class CreateMakerPOHeaderModel
    {
        public long VendID { get; set; }
        public DateTime DT { get; set; } = DateTime.Today;
        public int ProcessID { get; set; }
        public string ItemID { get; set; } = string.Empty;
        public string SpecialInstructions { get; set; } = string.Empty;
        public decimal ExcessQtyPercentage { get; set; } = 0m;
        public int MaximumRcvingsAgainstPO { get; set; } = 1;
        public bool SampleProvided { get; set; }
        public bool DrawingProvided { get; set; }
        public bool ForgingProvided { get; set; }
        public bool SteelProvided { get; set; }
        public string IssEmpID { get; set; } = string.Empty;
        public int SteelType_RefID { get; set; }
        public string CheckerEmpID { get; set; } = string.Empty;
    }

    public class CreateMakerPOLineModel
    {
        public string OrderNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public decimal Rate { get; set; }
        public decimal IssQty { get; set; }
        public int ReturnProcessID { get; set; }
        public DateTime ReturnDT { get; set; }
        public int Priority { get; set; }
        public string BatchNo { get; set; } = string.Empty;
        public string CountedBy { get; set; } = string.Empty;
        public string IssEmpID { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
    }

    public class MakerBalanceSummaryModel
    {
        public string ProcessName { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public decimal RemQty { get; set; }
    }

    public class MakerHistorySummaryModel
    {
        public long EntryID { get; set; }
        public string ProcessName { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public DateTime DT { get; set; }
        public decimal IssQty { get; set; }
        public decimal Rate { get; set; }
        public decimal IssValue { get; set; }
        public DateTime ReturnDT { get; set; }
        public int Priority { get; set; }
    }
}
