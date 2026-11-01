using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.FixedAssets
{
    public sealed class FixedAssetRegisterRow
    {
        public string AssetId { get; set; } = string.Empty;
        public DateTime? AcquiredDate { get; set; }
        public string AcquiredDateText => AcquiredDate.HasValue ? AcquiredDate.Value.ToString("dd-MMM-yyyy") : "-";
        public string Particulars { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public decimal Cost { get; set; }
        public decimal BookValue { get; set; }
        public string DisposalDateText { get; set; } = "-";
        public DateTime? DisposalDate { get; set; }
        public decimal DisposedValue { get; set; }
        public decimal OpeningBalance { get; set; }
        public decimal AccountingYear { get; set; }
        public decimal AccumulatedBalance { get; set; }
        public decimal WrittenDownValue { get; set; }
        public decimal Rate { get; set; }
        public decimal AdjustmentOpening { get; set; }
        public decimal AdjustedAmount { get; set; }
        public string DeptId { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public int MasterId { get; set; }
        public string MasterDescription { get; set; } = string.Empty;
        public bool Disposed { get; set; }
        public int DisposedType { get; set; }
        public bool InActive { get; set; }
        public DateTime? InActiveDate { get; set; }
        public DateTime? RetirementDate { get; set; }
    }

    public sealed class FixedAssetRegisterPeriod
    {
        public string DisplayText { get; set; } = string.Empty;
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
    }

    public sealed class FixedAssetRegisterFilterOptions
    {
        public List<FixedAssetLookupItem> Departments { get; set; } = new();
        public List<FixedAssetLookupItem> Categories { get; set; } = new();
        public List<FixedAssetRegisterPeriod> Periods { get; set; } = new();
        public DateTime DefaultDateFrom { get; set; }
        public DateTime DefaultDateTo { get; set; }
        public DateTime? LastSavedDateFrom { get; set; }
        public DateTime? LastSavedDateTo { get; set; }
    }

    public sealed class FixedAssetRegisterSaveResult
    {
        public bool Success { get; set; }
        public string? Message { get; set; }
        public int EntryId { get; set; }
    }
}

