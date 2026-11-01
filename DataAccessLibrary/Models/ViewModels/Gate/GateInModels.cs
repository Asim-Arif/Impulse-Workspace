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
        _ => "<All>"
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
    public string DepartmentId { get; set; } = "";
    public int? GatePassType { get; set; }
    public string GatePassTypeName { get; set; } = "<All Gatepass Types>";
    public string RangeKey { get; set; } = GateInRangeKeys.Last15Days;
    public DateTime FromDate { get; set; } = DateTime.Today.AddDays(-15);
    public DateTime ToDate { get; set; } = DateTime.Today;
    public string VendorName { get; set; } = "<All Vendors>";
}

public sealed class GateInListRowModel
{
    public int EntryId { get; set; }
    public DateTime? DateTime { get; set; }
    public string GatePassNo { get; set; } = "";
    public string PersonName { get; set; } = "";
    public string VehicleNo { get; set; } = "";
    public string Description { get; set; } = "";
    public string VendorName { get; set; } = "";
    public string DepartmentId { get; set; } = "";
    public int GatePassType { get; set; }
    public string GatePassTypeName { get; set; } = "";
    public int GatePassKind { get; set; }
    public bool IsFinanceApproved { get; set; }
    public bool HasReturnableItems { get; set; }
    public bool IsReceived { get; set; }
    public DateTime? ReceivedDate { get; set; }
}

public sealed class GateInEntryModel
{
    public int EntryId { get; set; }
    public DateTime DateTime { get; set; } = DateTime.Now;

    [Required(ErrorMessage = "Please select department.")]
    public string DepartmentId { get; set; } = "";
    public string DepartmentName { get; set; } = "";

    public string EmployeeId { get; set; } = "";
    public string EmployeeText { get; set; } = "";

    [Required(ErrorMessage = "Please enter Gate Pass No.")]
    [StringLength(50)]
    public string GatePassNo { get; set; } = "";

    [StringLength(255)]
    public string PersonName { get; set; } = "";

    [StringLength(50)]
    public string VehicleNo { get; set; } = "";

    [StringLength(255)]
    public string Description { get; set; } = "";

    public int VendorId { get; set; }
    public string VendorText { get; set; } = "<Select Vendor>";
    public string GatePassTypeName { get; set; } = "";
    public int GatePassType { get; set; } = GateInPassTypes.Foc;
    public int GatePassKind { get; set; } = GateInPassKinds.Permanent;
    public bool IsFinanceApproved { get; set; }
    public List<GateInDetailModel> Details { get; set; } = new();
    public bool IsFullyReceived => Details.Any(x => x.IsReturnable)
        && Details.Where(x => x.IsReturnable).All(x => x.BalanceQty <= 0);
}

public sealed class GateInDetailModel
{
    public int EntryId { get; set; }

    [StringLength(255)]
    public string Detail { get; set; } = "";

    [StringLength(50)]
    public string Unit { get; set; } = "";

    public decimal Qty { get; set; }
    public decimal Rate { get; set; }
    public string ReturnType { get; set; } = GateReturnTypes.NonReturnable;
    public DateTime? ReturnDate { get; set; }
    public decimal ReceivedQty { get; set; }
    public DateTime? ReceivedDate { get; set; }
    public string ReceivedBy { get; set; } = "";
    public string ReceivedRemarks { get; set; } = "";
    public decimal BalanceQty => Math.Max(0, Qty - ReceivedQty);
    public bool IsReturnable => GateReturnTypes.IsReturnable(ReturnType);
    public decimal Amount => Qty * Rate;
}

public sealed class GatePassTypeModel
{
    public int EntryId { get; set; }

    [Required(ErrorMessage = "Please enter gate-pass type.")]
    [StringLength(50)]
    public string GatePassType { get; set; } = "";

    public bool Active { get; set; } = true;
}

public sealed class GateReceiveModel
{
    public int EntryId { get; set; }
    public string GatePassNo { get; set; } = "";
    public DateTime DateTime { get; set; } = DateTime.Now;
    public string GatePassTypeName { get; set; } = "";
    public string PersonName { get; set; } = "";
    public string VendorName { get; set; } = "";
    public DateTime ReceiveDate { get; set; } = DateTime.Today;
    public string Remarks { get; set; } = "";
    public List<GateReceiveLineModel> Lines { get; set; } = new();
}

public sealed class GateReceiveLineModel
{
    public int DetailId { get; set; }
    public string Detail { get; set; } = "";
    public string Unit { get; set; } = "";
    public decimal Qty { get; set; }
    public decimal AlreadyReceivedQty { get; set; }
    public decimal BalanceQty { get; set; }
    public decimal ReceiveQty { get; set; }
}

public sealed class GateLookupOptionModel
{
    public string Id { get; set; } = "";
    public string Text { get; set; } = "";
}

