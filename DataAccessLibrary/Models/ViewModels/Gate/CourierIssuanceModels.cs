using System;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public static class CourierIssuanceRangeKeys
{
    public const string Last15Days = "15";
    public const string Last30Days = "30";
    public const string Last60Days = "60";
    public const string Last90Days = "90";
    public const string Last180Days = "180";
    public const string Custom = "custom";
}

public sealed class CourierIssuanceFilterModel
{
    public string RangeKey { get; set; } = CourierIssuanceRangeKeys.Last15Days;
    public DateTime FromDate { get; set; } = DateTime.Today.AddDays(-15);
    public DateTime ToDate { get; set; } = DateTime.Today;
    public string? VendID { get; set; }
    public string VendorName { get; set; } = "<All Vendors>";
    public string SearchText { get; set; } = "";
}

public sealed class CourierIssuanceListRowModel
{
    public int EntryID { get; set; }
    public DateTime? DT { get; set; }
    public string? TrackingNo { get; set; }
    public string? Description { get; set; }
    public string? VendorName { get; set; }
    public string? VendorAddress { get; set; }
    public string? IssuedBy { get; set; }
    public string? ConcernedPerson { get; set; }
    public string? ReceivedBy { get; set; }
    public bool? Self { get; set; }
    public bool? Paid { get; set; }
    public string? VendID { get; set; }
    public string? Weight { get; set; }
    public string? Dimension { get; set; }
    public bool? CourierStatus { get; set; }
    public string? MachineName { get; set; }
    public string? UserName { get; set; }
    public float? Parcel_Qty { get; set; }
    public string? Vehicle_No { get; set; }

    public string PaymentTypeCaption => Self == true ? "Self" : (Paid == true ? "Paid" : "-");
    public bool HasTracking => !string.IsNullOrWhiteSpace(TrackingNo);
}

public sealed class CourierIssuanceEntryModel
{
    public int EntryID { get; set; }

    [Required(ErrorMessage = "Date is required.")]
    public DateTime DT { get; set; } = DateTime.Now;

    [MaxLength(50)]
    public string? TrackingNo { get; set; }

    [MaxLength(255)]
    public string? Description { get; set; }

    public string? VendID { get; set; }

    [MaxLength(50)]
    public string? VendorName { get; set; }

    [MaxLength(255)]
    public string? VendorAddress { get; set; }

    [MaxLength(50)]
    public string? IssuedBy { get; set; }

    [MaxLength(50)]
    public string? ConcernedPerson { get; set; }

    [MaxLength(50)]
    public string? ReceivedBy { get; set; }

    public string PaymentMode { get; set; } = "Self"; // "Self" or "Paid"

    public bool Self => string.Equals(PaymentMode, "Self", StringComparison.OrdinalIgnoreCase);
    public bool Paid => string.Equals(PaymentMode, "Paid", StringComparison.OrdinalIgnoreCase);

    [MaxLength(50)]
    public string? Weight { get; set; }

    [MaxLength(50)]
    public string? Dimension { get; set; }

    public float? Parcel_Qty { get; set; }

    [MaxLength(50)]
    public string? Vehicle_No { get; set; }

    public bool PrintSlipOnSave { get; set; } = false;
}

public sealed class CourierVendorLookupModel
{
    public string AccNo { get; set; } = string.Empty;
    public string AccTitle { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string DisplayText => string.IsNullOrWhiteSpace(AccNo) ? AccTitle : $"[{AccNo}] {AccTitle}";
}

public sealed class CourierEmployeeLookupModel
{
    public string EmpID { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string FName { get; set; } = string.Empty;
    public string Designation { get; set; } = string.Empty;
    public string DisplayText => !string.IsNullOrWhiteSpace(EmpID) && EmpID != "0" ? $"[{EmpID}] {Name}" : Name;
}

