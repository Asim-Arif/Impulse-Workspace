using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Production
{
    public class RejectionAuthItemModel
    {
        public int EntryID { get; set; }
        public string RecieptID { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int Qty { get; set; }
        public int Rate { get; set; }
        public int RepairAmt { get; set; }
        public decimal WastageAmt => (decimal)Qty * Rate;
        public string LotNo { get; set; } = string.Empty;
        public int RcvdQty { get; set; }
        public int? VendID { get; set; }
        public int? FromVendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;
        public string? EmpID { get; set; }
        public string? CatID { get; set; }
        public bool? DeductFromBilling { get; set; }

        public string StatusText => DeductFromBilling == true
            ? "Deduct"
            : (DeductFromBilling == false ? "Ignored" : "Pending");
    }

    public class RejectionAuthMakerLookup
    {
        public int VendID { get; set; }
        public string VendID1 { get; set; } = string.Empty;
        public string VenderName { get; set; } = string.Empty;

        public string DisplayName => VendID == 0
            ? "<Show All>"
            : (string.IsNullOrWhiteSpace(VendID1) ? VenderName : $"{{ {VendID1} }} {VenderName}");
    }

    public class RejectionAuthEmployeeLookup
    {
        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string FName { get; set; } = string.Empty;
        public string Designation { get; set; } = string.Empty;
    }

    public class RejectionAuthCategoryLookup
    {
        public string CatID { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class RejectionAuthFilterModel
    {
        public bool ShowEmployees { get; set; }
        public int VendID { get; set; }
        public string EmpID { get; set; } = "0";
        public int DateRangeIndex { get; set; } = 1; // 0=Today, 1=Last 15 Days, 2=Last 30 Days, 3=Last 60 Days, 4=Last 90 Days, 5=Custom
        public DateTime DtFrom { get; set; } = DateTime.Today.AddDays(-15);
        public DateTime DtTo { get; set; } = DateTime.Today;
        public string LotNo { get; set; } = string.Empty;
        public List<string> CategoryIds { get; set; } = new();
    }
}
