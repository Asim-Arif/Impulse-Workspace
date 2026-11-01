using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.FixedAssets
{
    public sealed class FixedAssetDashboardFilter
    {
        public string FiscalYear { get; set; } = GetCurrentFinancialYear();
        public int QuarterIndex { get; set; } = GetCurrentFiscalQuarter(); // 0=All Quarters, 1=Q1(Jul-Sep), 2=Q2(Oct-Dec), 3=Q3(Jan-Mar), 4=Q4(Apr-Jun)
        public string? DepartmentId { get; set; }
        public int? MasterId { get; set; }
        public string? SearchText { get; set; }
        public string Status { get; set; } = "All";

        public static string GetCurrentFinancialYear()
        {
            var today = DateTime.Today;
            var startYear = today.Month >= 7 ? today.Year : today.Year - 1;
            return $"{startYear}-{startYear + 1}";
        }

        public static int GetCurrentFiscalQuarter()
        {
            var monthIndex = DateTime.Today.Month switch
            {
                7 => 1,
                8 => 2,
                9 => 3,
                10 => 4,
                11 => 5,
                12 => 6,
                1 => 7,
                2 => 8,
                3 => 9,
                4 => 10,
                5 => 11,
                _ => 12
            };

            return ((monthIndex - 1) / 3) + 1;
        }
    }

    public sealed class FixedAssetDashboardSummary
    {
        public int TotalAssets { get; set; }
        public int ActiveAssets { get; set; }
        public int DisposedAssets { get; set; }
        public int InactiveAssets { get; set; }
        public decimal TotalAcquisitionCost { get; set; }
        public decimal TotalAdjustments { get; set; }
        public decimal TotalGrossValue => TotalAcquisitionCost + TotalAdjustments;
        public decimal TotalAccumulatedDepreciation { get; set; }
        public decimal TotalNetBookValue => Math.Max(0, TotalGrossValue - TotalAccumulatedDepreciation);
        public decimal TotalDisposalProceeds { get; set; }
        public decimal TotalDisposalGainLoss { get; set; }
        public double DepreciationPercentage => TotalGrossValue > 0 ? (double)(TotalAccumulatedDepreciation / TotalGrossValue * 100) : 0;
        public double ActiveRate => TotalAssets > 0 ? (double)ActiveAssets / TotalAssets * 100 : 0;
        public int CurrentYearAdditionsCount { get; set; }
        public decimal CurrentYearAdditionsCost { get; set; }
        public decimal CurrentYearDepreciation { get; set; }
    }

    public sealed class FixedAssetCategoryCard
    {
        public int MasterId { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public int AssetCount { get; set; }
        public decimal TotalCost { get; set; }
        public decimal AccumulatedDepreciation { get; set; }
        public decimal NetBookValue { get; set; }
        public decimal DepreciationRate { get; set; }
        public double CostPercentage { get; set; }
        public string BadgeColor { get; set; } = "primary";
    }

    public sealed class FixedAssetDepartmentCard
    {
        public string DeptId { get; set; } = string.Empty;
        public string DeptName { get; set; } = string.Empty;
        public int AssetCount { get; set; }
        public decimal TotalCost { get; set; }
        public decimal NetBookValue { get; set; }
        public double CostPercentage { get; set; }
    }

    public sealed class FixedAssetYearTrend
    {
        public string FiscalYear { get; set; } = string.Empty;
        public decimal AcquisitionCost { get; set; }
        public decimal Depreciation { get; set; }
        public int AssetCount { get; set; }
    }

    public sealed class FixedAssetStatusBreakdown
    {
        public string StatusName { get; set; } = string.Empty;
        public int Count { get; set; }
        public decimal TotalCost { get; set; }
        public double Percentage { get; set; }
        public string ColorHex { get; set; } = "#3b82f6";
        public string BadgeClass { get; set; } = "bg-primary";
        public string Icon { get; set; } = "bi-check-circle";
    }

    public sealed class FixedAssetTableRow
    {
        public int EntryId { get; set; }
        public string AssetId { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public DateTime? AcquiredDate { get; set; }
        public decimal Cost { get; set; }
        public decimal AccumulatedDepreciation { get; set; }
        public decimal BookValue { get; set; }
        public string Status { get; set; } = "Active";
        public bool Disposed { get; set; }
        public bool InActive { get; set; }
    }

    public sealed class FixedAssetDisposalRow
    {
        public int EntryId { get; set; }
        public string AssetId { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public DateTime? DisposalDate { get; set; }
        public string DisposalType { get; set; } = string.Empty;
        public decimal Cost { get; set; }
        public decimal SalePrice { get; set; }
        public decimal GainLoss { get; set; }
        public string Buyer { get; set; } = string.Empty;
    }

    public sealed class FixedAssetDashboardResult
    {
        public FixedAssetDashboardSummary Summary { get; set; } = new();
        public List<FixedAssetCategoryCard> Categories { get; set; } = new();
        public List<FixedAssetDepartmentCard> Departments { get; set; } = new();
        public List<FixedAssetYearTrend> YearlyTrends { get; set; } = new();
        public List<FixedAssetStatusBreakdown> StatusBreakdown { get; set; } = new();
        public List<FixedAssetTableRow> RecentAssets { get; set; } = new();
        public List<FixedAssetTableRow> TopValueAssets { get; set; } = new();
        public List<FixedAssetDisposalRow> RecentDisposals { get; set; } = new();
        public List<string> FiscalYears { get; set; } = new();
        public List<FixedAssetLookupItem> LookupCategories { get; set; } = new();
        public List<FixedAssetLookupItem> LookupDepartments { get; set; } = new();
    }
}

