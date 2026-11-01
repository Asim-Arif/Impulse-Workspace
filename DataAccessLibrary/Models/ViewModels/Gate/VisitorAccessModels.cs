using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public static class MiscGateInRangeKeys
{
    public const string Today = "today";
    public const string Last15Days = "15";
    public const string Last30Days = "30";
    public const string Last60Days = "60";
    public const string Last90Days = "90";
    public const string Last180Days = "180";
    public const string Custom = "custom";
}

public sealed class MiscGateInFilterModel
{
    public string RangeKey { get; set; } = MiscGateInRangeKeys.Today;
    public DateTime FromDate { get; set; } = DateTime.Today;
    public DateTime ToDate { get; set; } = DateTime.Today;
    public string EmployeeName { get; set; } = "";
    public string VendorName { get; set; } = "";
}

public sealed class MiscGateInListRowModel
{
    public int EntryId { get; set; }
    public DateTime? DateTime { get; set; }
    public string VisitorName { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string ConcernedPerson { get; set; } = "";
    public string Purpose { get; set; } = "";
    public DateTime? TimeOut { get; set; }
    public string VendorName { get; set; } = "";
}

public sealed class MiscGateInEntryModel
{
    public int EntryId { get; set; }

    [Required(ErrorMessage = "Date is required.")]
    public DateTime DateTime { get; set; } = DateTime.Now;

    [Required(ErrorMessage = "Visitor name is required.")]
    [StringLength(50)]
    public string VisitorName { get; set; } = "";

    [StringLength(50)]
    public string CnicNo { get; set; } = "";

    [StringLength(50)]
    public string CompanyName { get; set; } = "";

    [StringLength(255)]
    public string Purpose { get; set; } = "";

    [StringLength(50)]
    public string MobileNo { get; set; } = "";

    [StringLength(50)]
    public string CardNo { get; set; } = "";

    [StringLength(150)]
    [EmailAddress(ErrorMessage = "Enter a valid email address.")]
    public string ConcernedPersonEmail { get; set; } = "";

    [StringLength(50)]
    public string ConcernedPersonWhatsApp { get; set; } = "";

    public bool NotifyByEmail { get; set; }
    public bool NotifyByWhatsApp { get; set; }

    [StringLength(50)]
    public string VisitCode { get; set; } = "";

    public string? QrSvgDataUrl { get; set; }
    public string NotificationStatus { get; set; } = "";

    [StringLength(50)]
    public string EmployeeId { get; set; } = "";

    [StringLength(50)]
    public string EmployeeText { get; set; } = "";

    [StringLength(255)]
    public string Remarks { get; set; } = "";

    public int VendorId { get; set; }

    [StringLength(50)]
    public string VendorText { get; set; } = "<Select Vendor>";

    public bool HasTimeOut { get; set; }
    public DateTime? TimeOut { get; set; }
}

public sealed class MiscGateOutModel
{
    public int EntryId { get; set; }
    public DateTime DateTime { get; set; } = DateTime.Now;
    public DateTime TimeOut { get; set; } = DateTime.Now;
    public string VisitorName { get; set; } = "";
    public string CnicNo { get; set; } = "";
    public string CardNo { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string Purpose { get; set; } = "";
    public string MobileNo { get; set; } = "";
    public string ConcernedPerson { get; set; } = "";
    public string VendorName { get; set; } = "";
    public string Remarks { get; set; } = "";
    public bool IsAlreadyOut { get; set; }
    public string VisitCode { get; set; } = "";
    public string VerificationInput { get; set; } = "";
    public string? QrSvgDataUrl { get; set; }
    public string OutVerifiedBy { get; set; } = "";
    public DateTime? OutVerifiedAt { get; set; }
    public TimeSpan StayDuration => TimeOut >= DateTime ? TimeOut - DateTime : TimeSpan.Zero;
}

public sealed class ConcernedPersonContactModel
{
    public string EmployeeId { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string EmailAddress { get; set; } = "";
    public string WhatsAppNo { get; set; } = "";
}

public sealed class VisitorPassCardModel
{
    public int EntryId { get; set; }
    public string VisitorName { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string ConcernedPerson { get; set; } = "";
    public string CardNo { get; set; } = "";
    public string VisitCode { get; set; } = "";
    public DateTime EntryTime { get; set; }
    public string QrSvgDataUrl { get; set; } = "";
}
