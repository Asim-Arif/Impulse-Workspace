using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public class GatePassReturnableListRowModel
{
    public int EntryID { get; set; }
    public string GatePassNo { get; set; } = string.Empty;
    public DateTime? DT { get; set; }
    public string? PersonName { get; set; }
    public string? VehicleNo { get; set; }
    public string? Description { get; set; }
    public string? UserName { get; set; }
    public string? MachineName { get; set; }
    public DateTime? DTEntry { get; set; }
    public string? SenderName { get; set; }
    public bool Returnable { get; set; }
    public string? IssdTo { get; set; }
    public string? GatePassType { get; set; }
    public DateTime? ReturnDT { get; set; }
    public string? DeptID_From { get; set; }
    public string? DeptID_To { get; set; }
    public string? DeptName_From { get; set; }
    public string? DeptName_To { get; set; }
    public int? VendID { get; set; }
    public string? VM_AccNo { get; set; }
    public string? Name { get; set; } // Vendor/Maker AccTitle
    public string? ApprovedBy_Name { get; set; }
    public string? ApprovedBy_Designation { get; set; }
    public double TotalQty { get; set; }
    public double TotalQtyRcvd { get; set; }

    public string Status
    {
        get
        {
            if (!Returnable) return "Non-Returnable";
            if (TotalQtyRcvd <= 0) return "Open";
            if (TotalQtyRcvd >= TotalQty && TotalQty > 0) return "Closed";
            return "Partially";
        }
    }

    public string StatusBadgeClass
    {
        get
        {
            return Status switch
            {
                "Open" => "bg-warning text-dark",
                "Partially" => "bg-info text-dark",
                "Closed" => "bg-success",
                _ => "bg-secondary"
            };
        }
    }
}

public class GatePassReturnableFilterModel
{
    public int DateRangePreset { get; set; } = 0; // 0=15 days, 1=30 days, 2=60 days, 3=90 days, 4=180 days, 5=Custom
    public DateTime DtFrom { get; set; } = DateTime.Today.AddDays(-15);
    public DateTime DtTo { get; set; } = DateTime.Today;
    public string GatePassType { get; set; } = string.Empty;
    public string ReturnableFilter { get; set; } = "All"; // All, Yes, No
    public string VM_AccNo { get; set; } = string.Empty;
    public string SearchText { get; set; } = string.Empty;
}

public class GatePassReturnableDetailItemModel
{
    public int EntryID { get; set; }
    public int MEGO_RefID { get; set; }

    [Required(ErrorMessage = "Item description / detail is required.")]
    public string Detail { get; set; } = string.Empty;

    public string Unit { get; set; } = "Pcs";
    public string? Packing { get; set; }

    [Range(0.001, double.MaxValue, ErrorMessage = "Quantity must be greater than zero.")]
    public double Qty { get; set; } = 1;
}

public class GatePassReturnableEntryModel
{
    public int EntryID { get; set; }

    [Required(ErrorMessage = "Gate pass number is required.")]
    public string GatePassNo { get; set; } = string.Empty;

    public DateTime DT { get; set; } = DateTime.Now;

    public bool Returnable { get; set; } = true;
    public DateTime? ReturnDT { get; set; } = DateTime.Today.AddDays(7);

    [Required(ErrorMessage = "Gate pass type is required.")]
    public string GatePassType { get; set; } = string.Empty;

    [Required(ErrorMessage = "Issuance from department is required.")]
    public string DeptID_From { get; set; } = string.Empty;

    public string? DeptID_To { get; set; } // Employee ID or Department
    public string? IssdTo { get; set; } // Party / Person / Driver issued to
    public string? PersonName { get; set; } // Carrying Person
    public string? SenderName { get; set; } // Dispatcher / Sender
    public string? VM_AccNo { get; set; } // Vendor / Maker Account No
    public string? VendName { get; set; }

    public string? ApprovedBy_Name { get; set; } // Employee ID
    public string? ApprovedBy_Designation { get; set; }

    public bool PrintSlipOnSave { get; set; }

    public List<GatePassReturnableDetailItemModel> Items { get; set; } = new();
}

public class GatePassReceiveDetailItemModel
{
    public int DetailRefID { get; set; }
    public string Detail { get; set; } = string.Empty;
    public string Unit { get; set; } = "Pcs";
    public string? Packing { get; set; }
    public double IssuedQty { get; set; }
    public double PreviouslyReceivedQty { get; set; }
    public double BalanceQty => Math.Max(0, IssuedQty - PreviouslyReceivedQty);
    public double QtyReceivedNow { get; set; }
}

public class GatePassReceiveModel
{
    public int RefEntryID { get; set; }
    public string GatePassNo { get; set; } = string.Empty;
    public DateTime DT { get; set; }
    public string? IssdTo { get; set; }
    public string? VendorName { get; set; }
    public DateTime RcvDT { get; set; } = DateTime.Today;
    public bool PrintSlipOnSave { get; set; }
    public List<GatePassReceiveDetailItemModel> Items { get; set; } = new();
}

