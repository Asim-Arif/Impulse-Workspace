using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public static class GateInRangeKeys
{
    public const string Last15Days = "15";
    public const string Last30Days = "30";
    public const string Last60Days = "60";
    public const string Last90Days = "90";
    public const string Last180Days = "180";
    public const string Custom = "custom";
}

public static class GateInPassTypes
{
    public const int Foc = 0;
    public const int ExportSale = 1;
    public const int Sales = 2;
    public const int All = 99;

    public static string Caption(int value) => value switch
    {
        Foc => "FOC",
        ExportSale => "Export Sales",
        Sales => "Sales",
        _ => "<All Types>"
    };
}

public static class GateInPassKinds
{
    public const int Permanent = 0;
    public const int Temporary = 1;

    public static string Caption(int value) => value == Temporary
        ? "Temporary Gate In"
        : "Permanent Gate In";
}

public static class GateReturnTypes
{
    public const string Returnable = "Returnable";
    public const string NonReturnable = "Non-Returnable";

    public static bool IsReturnable(string? value) =>
        string.Equals(value?.Trim(), Returnable, StringComparison.OrdinalIgnoreCase);
}

public sealed class GateInFilterModel
{
    public string DeptID { get; set; } = string.Empty;
    public string DepartmentId { get => DeptID; set => DeptID = value; }

    public int GatePassType { get; set; } = GateInPassTypes.All;
    public string GatePassTypeName { get; set; } = "<All Types>";

    public string RangeKey { get; set; } = GateInRangeKeys.Last15Days;
    public DateTime FromDate { get; set; } = DateTime.Today.AddDays(-15);
    public DateTime ToDate { get; set; } = DateTime.Today;

    public string VM_AccNo { get; set; } = string.Empty;
    public string VendorName { get; set; } = string.Empty;
    public string SearchText { get; set; } = string.Empty;
}

public sealed class GateInListRowModel
{
    public int EntryId { get; set; }
    public int EntryID { get => EntryId; set => EntryId = value; }

    public DateTime? DT { get; set; }
    public DateTime? DateTime { get => DT; set => DT = value; }

    public string GatePassNo { get; set; } = string.Empty;
    public string PersonName { get; set; } = string.Empty;
    public string VehicleNo { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;

    public string DeptID { get; set; } = string.Empty;
    public string DepartmentId { get => DeptID; set => DeptID = value; }
    public string DeptName { get; set; } = string.Empty;

    public string EmpID { get; set; } = string.Empty;
    public string EmpName { get; set; } = string.Empty;

    public int GatePassType { get; set; }
    public string GatePassTypeName => GateInPassTypes.Caption(GatePassType);

    public int GatePassType2 { get; set; }
    public int GatePassKind { get => GatePassType2; set => GatePassType2 = value; }
    public string GatePassKindName => GateInPassKinds.Caption(GatePassType2);

    public string? VM_AccNo { get; set; }
    public string? Name { get; set; } // Vendor from VVendersAndMakers
    public string? VendNameManual { get; set; }
    public string DisplayVendorName => !string.IsNullOrWhiteSpace(Name) ? Name : (VendNameManual ?? string.Empty);
    public string VendorName { get => DisplayVendorName; set => Name = value; }

    public bool Approved_HOD { get; set; }
    public bool Approved { get; set; }
    public bool Approved_Finance { get; set; }
    public bool IsFinanceApproved { get => Approved_Finance; set => Approved_Finance = value; }

    public bool HasReturnableItems { get; set; }
    public bool IsReceived { get; set; }
    public DateTime? ReceivedDate { get; set; }

    public string ApprovalStatus
    {
        get
        {
            if (Approved_Finance) return "Finance Approved";
            if (Approved) return "Admin Approved";
            if (Approved_HOD) return "HOD Approved";
            return "Pending";
        }
    }

    public string ApprovalBadgeClass
    {
        get
        {
            if (Approved_Finance) return "bg-success";
            if (Approved) return "bg-info text-dark";
            if (Approved_HOD) return "bg-primary";
            return "bg-secondary";
        }
    }
}

public sealed class GateInEntryModel
{
    public int EntryId { get; set; }
    public int EntryID { get => EntryId; set => EntryId = value; }

    public DateTime DT { get; set; } = DateTime.Now;
    public DateTime DateTime { get => DT; set => DT = value; }

    [Required(ErrorMessage = "Please select Department.")]
    public string DeptID { get; set; } = string.Empty;
    public string DepartmentId { get => DeptID; set => DeptID = value; }
    public string DeptName { get; set; } = string.Empty;
    public string DepartmentName { get => DeptName; set => DeptName = value; }

    public string EmpID { get; set; } = string.Empty;
    public string EmployeeId { get => EmpID; set => EmpID = value; }
    public string EmpName { get; set; } = string.Empty;
    public string EmployeeText { get => EmpName; set => EmpName = value; }

    [Required(ErrorMessage = "Please enter Gate Pass No.")]
    [StringLength(50)]
    public string GatePassNo { get; set; } = string.Empty;

    [StringLength(255)]
    public string PersonName { get; set; } = string.Empty;

    [StringLength(50)]
    public string VehicleNo { get; set; } = string.Empty;

    [StringLength(255)]
    public string Description { get; set; } = string.Empty;

    public int GatePassType { get; set; } = GateInPassTypes.Foc;
    public string GatePassTypeName { get; set; } = string.Empty;

    public int GatePassType2 { get; set; } = GateInPassKinds.Permanent;
    public int GatePassKind { get => GatePassType2; set => GatePassType2 = value; }

    public int VendorId { get; set; }
    public string? VM_AccNo { get; set; }
    public string? VendorName { get; set; }
    public string VendorText { get => VendorName ?? string.Empty; set => VendorName = value; }

    [StringLength(500)]
    public string? VendNameManual { get; set; }

    public bool Approved_HOD { get; set; }
    public bool Approved { get; set; }
    public bool Approved_Finance { get; set; }
    public bool IsFinanceApproved { get => Approved_Finance; set => Approved_Finance = value; }

    public bool PrintSlipOnSave { get; set; }

    public List<GateInDetailModel> Details { get; set; } = new();

    public bool IsFullyReceived => Details.Any(x => x.IsReturnable)
        && Details.Where(x => x.IsReturnable).All(x => x.BalanceQty <= 0);
}

public sealed class GateInDetailModel
{
    public int EntryId { get; set; }
    public int EntryID { get => EntryId; set => EntryId = value; }
    public int MEGO_RefID { get; set; }

    [Required(ErrorMessage = "Please enter item detail/description.")]
    [StringLength(255)]
    public string Detail { get; set; } = string.Empty;

    [StringLength(50)]
    public string Unit { get; set; } = "Pcs";

    [Range(0.001, double.MaxValue, ErrorMessage = "Qty must be greater than zero.")]
    public decimal Qty { get; set; } = 1;

    public decimal Rate { get; set; } = 0;
    public decimal Amount => Qty * Rate;

    public string ReturnType { get; set; } = GateReturnTypes.NonReturnable;
    public DateTime? ReturnDate { get; set; }
    public decimal ReceivedQty { get; set; }
    public DateTime? ReceivedDate { get; set; }
    public string ReceivedBy { get; set; } = string.Empty;
    public string ReceivedRemarks { get; set; } = string.Empty;
    public decimal BalanceQty => Math.Max(0, Qty - ReceivedQty);
    public bool IsReturnable => GateReturnTypes.IsReturnable(ReturnType);
}

public sealed class GateInEmployeeLookupModel
{
    public string EmpID { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string? FName { get; set; }
    public string? Designation { get; set; }
    public string? DeptID { get; set; }
    public string? DeptName { get; set; }
    public string DisplayText => $"[{EmpID}] {Name}" + (!string.IsNullOrWhiteSpace(Designation) ? $" - {Designation}" : "");
}

public sealed class GateInVendorLookupModel
{
    public string AccNo { get; set; } = string.Empty;
    public string AccTitle { get; set; } = string.Empty;
    public string DisplayText => $"{AccTitle} {{{AccNo}}}";
}

public sealed class GateLookupOptionModel
{
    public string Id { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}

public sealed class GateReceiveModel
{
    public int EntryId { get; set; }
    public string GatePassNo { get; set; } = string.Empty;
    public DateTime DateTime { get; set; } = DateTime.Now;
    public string GatePassTypeName { get; set; } = string.Empty;
    public string PersonName { get; set; } = string.Empty;
    public string VendorName { get; set; } = string.Empty;
    public DateTime ReceiveDate { get; set; } = DateTime.Today;
    public string Remarks { get; set; } = string.Empty;
    public List<GateReceiveLineModel> Lines { get; set; } = new();
}

public sealed class GatePassTypeModel
{
    public int EntryId { get; set; }

    [Required(ErrorMessage = "Please enter gate-pass type.")]
    [StringLength(50)]
    public string GatePassType { get; set; } = string.Empty;

    public bool Active { get; set; } = true;
}

public sealed class GateReceiveLineModel
{
    public int DetailId { get; set; }
    public string Detail { get; set; } = string.Empty;
    public string Unit { get; set; } = string.Empty;
    public decimal Qty { get; set; }
    public decimal AlreadyReceivedQty { get; set; }
    public decimal BalanceQty { get; set; }
    public decimal ReceiveQty { get; set; }
}
