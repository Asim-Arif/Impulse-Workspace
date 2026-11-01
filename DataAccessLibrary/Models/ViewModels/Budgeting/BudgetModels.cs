using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Budgeting
{
    public class FinancialYearOption
    {
        public string Code { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
        public int StartYear { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }

    public class MasterHeadOption
    {
        public string Code { get; set; } = string.Empty;
        public string Head { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public bool BudgetAccountGlance { get; set; }
        public string DisplayText => string.IsNullOrEmpty(Code) || Code == "0" ? Head : $"{Code} - {Head}";
    }

    public class MonthwiseBudgetRow
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string HeadName { get; set; } = string.Empty;

        public decimal MonthJuly { get; set; }
        public decimal MonthAugust { get; set; }
        public decimal MonthSeptember { get; set; }
        public decimal MonthOctober { get; set; }
        public decimal MonthNovember { get; set; }
        public decimal MonthDecember { get; set; }
        public decimal MonthJanuary { get; set; }
        public decimal MonthFebruary { get; set; }
        public decimal MonthMarch { get; set; }
        public decimal MonthApril { get; set; }
        public decimal MonthMay { get; set; }
        public decimal MonthJune { get; set; }

        public decimal RowTotal => MonthJuly + MonthAugust + MonthSeptember +
                                   MonthOctober + MonthNovember + MonthDecember +
                                   MonthJanuary + MonthFebruary + MonthMarch +
                                   MonthApril + MonthMay + MonthJune;
    }

    public class ExpenseBudgetRow
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string HeadName { get; set; } = string.Empty;

        public decimal Proposed { get; set; }
        public decimal Revised { get; set; }
    }

    public class MonthwiseBudgetTotals
    {
        public decimal July { get; set; }
        public decimal August { get; set; }
        public decimal September { get; set; }
        public decimal October { get; set; }
        public decimal November { get; set; }
        public decimal December { get; set; }
        public decimal January { get; set; }
        public decimal February { get; set; }
        public decimal March { get; set; }
        public decimal April { get; set; }
        public decimal May { get; set; }
        public decimal June { get; set; }
        public decimal GrandTotal => July + August + September + October + November + December +
                                     January + February + March + April + May + June;
    }

    public class ExpenseBudgetTotals
    {
        public decimal TotalProposed { get; set; }
        public decimal TotalRevised { get; set; }
    }

    public class BudgetReportFilter
    {
        public string FinancialYear { get; set; } = string.Empty;
        public string HeadCode { get; set; } = "0";
        public string HeadName { get; set; } = "<All Heads>";
        public int QuarterNo { get; set; } // 1, 2, 3, 4, or 0
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public bool IsExpenses { get; set; }
        public bool IsRevision { get; set; }
        public string ReportType { get; set; } = "quarter"; // "quarter", "date-range", "glance"
    }

    public class BudgetReportRow
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string HeadCode { get; set; } = string.Empty;
        public string HeadName { get; set; } = string.Empty;
        public decimal ProposedAmount { get; set; }
        public decimal RevisedAmount { get; set; }
        public decimal ActualAmount { get; set; }
        public decimal Variance => ActualAmount - ProposedAmount;
        public decimal VariancePercent => ProposedAmount != 0 ? Math.Round((Variance / ProposedAmount) * 100, 2) : 0;
    }
}

