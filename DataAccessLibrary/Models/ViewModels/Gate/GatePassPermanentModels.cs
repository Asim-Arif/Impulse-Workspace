using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public static class GatePassPermanentRangeKeys
{
    public const string Last15Days = "15";
    public const string Last30Days = "30";
    public const string Last60Days = "60";
    public const string Last90Days = "90";
    public const string Last180Days = "180";
    public const string Custom = "custom";
}

public static class GatePassPermanentTypes
{
    public const int FOC = 0;
    public const int ExportSale = 1;
    public const int Sale = 2;

    public static string GetCaption(int? type) => type switch
    {
        FOC => "FOC",
        ExportSale => "Export Sale",
        Sale => "Sale",
        _ => "General"
    };

    public static string GetBadgeClass(int? type) => type switch
    {
        FOC => "badge bg-info text-dark",
        ExportSale => "badge bg-primary",
        Sale => "badge bg-success",
        _ => "badge bg-secondary"
    };
}

public sealed class GatePassPermanentFilterModel
{
    public string RangeKey { get; set; } = GatePassPermanentRangeKeys.Last15Days;
    public DateTime FromDate { get; set; } = DateTime.Today.AddDays(-15);
    public DateTime ToDate { get; set; } = DateTime.Today;
    public string? DeptID { get; set; }
    public int? GatePassType { get; set; }
    public string? VM_AccNo { get; set; }
    public string VendorName { get; set; } = "<All Vendors>";
    public string SearchText { get; set; } = "";
}

public sealed class GatePassPermanentListRowModel
{
    public int EntryID { get; set; }
    public string? GatePassNo { get; set; }
    public DateTime? DT { get; set; }
    public string? PersonName { get; set; }
    public string? VehicleNo { get; set; }
    public string? Description { get; set; }
    public string? UserName { get; set; }
    public string? MachineName { get; set; }
    public DateTime? DTEntry { get; set; }
    public DateTime? OnlyDT { get; set; }
    public string? EmpID { get; set; }
    public int? GatePassType { get; set; }
    public string? DeptID { get; set; }
    public int? VendID { get; set; }
    public string? Name { get; set; }
    public string? VendNameManual { get; set; }
    public string? VM_AccNo { get; set; }
    public bool Approved_HOD { get; set; }
    public bool Approved { get; set; }
    public bool Approved_Finance { get; set; }

    public string EffectiveVendor => !string.IsNullOrWhiteSpace(Name)
        ? Name
        : (!string.IsNullOrWhiteSpace(VendNameManual) ? VendNameManual : "-");

    public string TypeCaption => GatePassPermanentTypes.GetCaption(GatePassType);
    public string TypeBadgeClass => GatePassPermanentTypes.GetBadgeClass(GatePassType);

    public bool IsFullyApproved => Approved_HOD && Approved && Approved_Finance;
}

public sealed class GatePassPermanentDetailItemModel
{
    public int EntryID { get; set; }
    public int MEGO_RefID { get; set; }

    [Required(ErrorMessage = "Item detail/description is required.")]
    [MaxLength(255)]
    public string Detail { get; set; } = "";

    [MaxLength(50)]
    public string Unit { get; set; } = "Pcs";

    [Range(0.001, double.MaxValue, ErrorMessage = "Qty must be greater than 0.")]
    public float Qty { get; set; } = 1;

    [Range(0, double.MaxValue, ErrorMessage = "Rate cannot be negative.")]
    public float Rate { get; set; } = 0;

    public float Amount => Qty * Rate;
}

public sealed class GatePassPermanentEntryModel
{
    public int EntryID { get; set; }

    [Required(ErrorMessage = "Gate pass # is required.")]
    [MaxLength(50)]
    public string GatePassNo { get; set; } = "";

    [Required(ErrorMessage = "Date & time are required.")]
    public DateTime DT { get; set; } = DateTime.Now;

    [Required(ErrorMessage = "Please select a department.")]
    [MaxLength(50)]
    public string DeptID { get; set; } = "";

    public int GatePassType { get; set; } = GatePassPermanentTypes.FOC;

    [MaxLength(255)]
    public string? PersonName { get; set; }

    [MaxLength(50)]
    public string? VehicleNo { get; set; }

    [MaxLength(50)]
    public string? EmpID { get; set; }

    [MaxLength(50)]
    public string? VM_AccNo { get; set; }

    [MaxLength(50)]
    public string? VendNameManual { get; set; }

    [MaxLength(255)]
    public string? Description { get; set; }

    public bool PrintSlipOnSave { get; set; } = false;

    public List<GatePassPermanentDetailItemModel> Items { get; set; } = new();

    public float TotalQty => Items.Sum(x => x.Qty);
    public float TotalAmount => Items.Sum(x => x.Amount);
}

