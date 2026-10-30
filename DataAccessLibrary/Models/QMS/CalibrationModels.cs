using System;

namespace DataAccessLibrary.Models.QMS
{
    public class CalibrationListItemModel
    {
        public int EntryID { get; set; }
        public string AssetID { get; set; } = string.Empty;
        public string? Description { get; set; }
        public DateTime? DT { get; set; }
        public string? CertificateNo { get; set; }
        public string? MasterReference { get; set; }
        public string? AccNo { get; set; }
        public string? AccTitle { get; set; }
        public string? Remarks { get; set; }
        public DateTime? NextDueDT { get; set; }
        public string? UserName { get; set; }
        public DateTime? DTEntry { get; set; }
        public int AttachmentCount { get; set; }

        public string AssetDisplay => string.IsNullOrWhiteSpace(Description) ? AssetID : $"({AssetID}) {Description}";
        public bool IsOverdue => NextDueDT.HasValue && NextDueDT.Value.Date < DateTime.Today;
        public bool IsDueSoon => NextDueDT.HasValue && NextDueDT.Value.Date >= DateTime.Today && NextDueDT.Value.Date <= DateTime.Today.AddDays(7);
    }

    public class CalibrationModel
    {
        public int EntryID { get; set; }
        public string AssetID { get; set; } = string.Empty;
        public string? AssetDescription { get; set; }
        public DateTime? DT { get; set; } = DateTime.Today;
        public DateTime? NextDueDT { get; set; } = DateTime.Today.AddDays(7);
        public int? CalibrationFrequency { get; set; } = 0; // 0: Monthly, 1: Yearly
        public string? CertificateNo { get; set; }
        public string? MasterReference { get; set; }
        public string? AccNo { get; set; }
        public string? AccTitle { get; set; }
        public string? Remarks { get; set; }
        public string? UserName { get; set; }
        public string? MachineName { get; set; }
        public DateTime? DTEntry { get; set; }
    }

    public class FixedAssetLookupDto
    {
        public string AssetID { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool? RequiresCalibration { get; set; }
        public int? CalibrationFrequency { get; set; }
        public string DisplayText => $"({AssetID}) {Description}";
    }

    public class VendorLookupDto
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string DisplayText => $"{AccTitle} ({AccNo})";
    }

    public class CalibrationFilterDto
    {
        public string? AssetID { get; set; } = "0";
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string? SearchTerm { get; set; }
    }
}

