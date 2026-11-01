using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.FixedAssets
{
    public sealed class FixedAssetProfileModel
    {
        public int EntryId { get; set; }
        [Required] public string AssetId { get; set; } = string.Empty;
        [Required] public string Description { get; set; } = string.Empty;
        public int? MasterAssetRefId { get; set; }
        public string OldCode { get; set; } = string.Empty;
        public DateTime? AcquiredDate { get; set; } = DateTime.Today;
        public string Location { get; set; } = string.Empty;
        public string ModelNo { get; set; } = string.Empty;
        public string SerialNo { get; set; } = string.Empty;
        public string Manufacturer { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
        public string EmployeeId { get; set; } = string.Empty;
        public string DepartmentId { get; set; } = string.Empty;
        public string VendorAccountNo { get; set; } = string.Empty;
        public byte Status { get; set; }
        public DateTime? PlacedInServiceDate { get; set; } = DateTime.Today;
        public decimal AcquisitionCost { get; set; }
        public decimal AccumulatedDepreciation { get; set; }
        public decimal BookValue { get; set; }
        public bool RequiresCalibration { get; set; }
        public int? CalibrationFrequency { get; set; }
        public string AccumulatedAccountNo { get; set; } = string.Empty;
        public string DepreciationAccountNo { get; set; } = string.Empty;
        public byte? DepreciationMethodId { get; set; }
        public string Convention { get; set; } = string.Empty;
        public decimal DepreciationRate { get; set; }
        public DateTime? DepreciationBeginMonth { get; set; }
        public int? LifeInYears { get; set; }
        public bool InActive { get; set; }
    }

    public sealed class FixedAssetLookupOptions
    {
        public IReadOnlyList<FixedAssetLookupItem> Masters { get; set; } = Array.Empty<FixedAssetLookupItem>();
        public IReadOnlyList<FixedAssetLookupItem> Departments { get; set; } = Array.Empty<FixedAssetLookupItem>();
        public IReadOnlyList<FixedAssetLookupItem> Employees { get; set; } = Array.Empty<FixedAssetLookupItem>();
        public IReadOnlyList<FixedAssetLookupItem> Accounts { get; set; } = Array.Empty<FixedAssetLookupItem>();
        public IReadOnlyList<string> Conventions { get; set; } = Array.Empty<string>();
    }

    public sealed class FixedAssetLookupItem
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public decimal PercentRate { get; set; }
        public string Display => string.IsNullOrWhiteSpace(Name) ? Id : $"{Id} - {Name}";
    }

    public sealed class FixedAssetListRow
    {
        public int EntryId { get; set; }
        public string AssetId { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string OldCode { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public DateTime? AcquiredDate { get; set; }
        public decimal Cost { get; set; }
        public decimal AccumulatedDepreciation { get; set; }
        public bool AutoEntry { get; set; }
        public string VoucherNo { get; set; } = string.Empty;
        public string GatePassStatus { get; set; } = string.Empty;
        public string ReceiptNoteStatus { get; set; } = string.Empty;
        public bool InActive { get; set; }
        public bool Disposed { get; set; }
        public int DisposedType { get; set; }
        public string Category { get; set; } = string.Empty;
    }

    public sealed class FixedAssetRepairLogRow
    {
        public int SerialNo { get; set; }
        public DateTime? Date { get; set; }
        public string Remarks { get; set; } = string.Empty;
        public int MeterReading { get; set; }
        public DateTime? ScheduleDate { get; set; }
        public decimal Cost { get; set; }
        public string BillReference { get; set; } = string.Empty;
    }

    public sealed class FixedAssetAdjustmentModel
    {
        public DateTime AdjustmentDate { get; set; } = DateTime.Today;
        public string AssetId { get; set; } = string.Empty;
        public DateTime? CapitalizationDate { get; set; }
        public decimal AcquisitionCost { get; set; }
        public decimal PreviousAdjustment { get; set; }
        public decimal AdjustmentAmount { get; set; }
        public decimal TotalCost => AcquisitionCost + PreviousAdjustment + AdjustmentAmount;
        public string Remarks { get; set; } = string.Empty;
        public string PdfPath { get; set; } = string.Empty;
    }

    public sealed class FixedAssetMoveModel
    {
        public int? CurrentHeadId { get; set; }
        public string AssetId { get; set; } = string.Empty;
        public string OriginalHeadName { get; set; } = string.Empty;
        public int? NewHeadId { get; set; }
        public string Remarks { get; set; } = string.Empty;
    }

    public sealed class FixedAssetDisposalModel
    {
        public string DisposalType { get; set; } = string.Empty;
        public string AssetId { get; set; } = string.Empty;
        public DateTime DisposalDate { get; set; } = DateTime.Today;
        public decimal Wdv { get; set; }
        public decimal SalePrice { get; set; }
        public decimal GainLoss => Wdv - SalePrice;
        public string Remarks { get; set; } = string.Empty;
    }
}

