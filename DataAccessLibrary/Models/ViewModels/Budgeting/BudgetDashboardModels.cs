using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Budgeting
{
    public class BudgetDashboardFilter
    {
        public string FinancialYear { get; set; } = GetCurrentFinancialYear();
        public int QuarterIndex { get; set; } = GetCurrentFiscalQuarter(); // 1=Jul-Sep, 2=Oct-Dec, 3=Jan-Mar, 4=Apr-Jun
        public int MonthIndex { get; set; } = 0; // 0 = All Year, 1..12 = Jul..Jun
        public string Mode { get; set; } = "Expenses"; // "Expenses", "Income"
        public string HeadCode { get; set; } = "0";
        public string? SearchText { get; set; }

        public static string GetCurrentFinancialYear()
        {
            var today = DateTime.Today;
            var startYear = today.Month >= 7 ? today.Year : today.Year - 1;
            return $"{startYear} - {startYear + 1}";
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

    public class BudgetDashboardSummary
    {
        public decimal TotalBudget { get; set; }
        public decimal TotalActual { get; set; }
        public decimal Variance => TotalBudget - TotalActual;
        public double PercentageUsed => TotalBudget > 0 ? (double)(TotalActual / TotalBudget * 100) : 0;
        public decimal BudgetProposed { get; set; }
        public decimal BudgetRevised { get; set; }
        public int HeadCount { get; set; }
        public int AccountCount { get; set; }
    }

    public class BudgetCategoryStat
    {
        public string CategoryCode { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public decimal BudgetAmount { get; set; }
        public decimal ActualAmount { get; set; }
        public decimal Variance => BudgetAmount - ActualAmount;
        public double PercentageUsed => BudgetAmount > 0 ? (double)(ActualAmount / BudgetAmount * 100) : 0;
        public double PercentageOfTotalBudget { get; set; }
        public string ColorHex { get; set; } = "#3b82f6";
        public string BadgeClass { get; set; } = "primary";
        public bool IsOverBudget => ActualAmount > BudgetAmount && BudgetAmount > 0;
    }

    public class BudgetMonthlyTrend
    {
        public int MonthIndex { get; set; }
        public string MonthName { get; set; } = string.Empty;
        public string MonthCode { get; set; } = string.Empty;
        public decimal BudgetAmount { get; set; }
        public decimal ActualAmount { get; set; }
        public decimal Variance => BudgetAmount - ActualAmount;
        public double PercentageUsed => BudgetAmount > 0 ? (double)(ActualAmount / BudgetAmount * 100) : 0;
    }

    public class BudgetAccountRow
    {
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public decimal BudgetAmount { get; set; }
        public decimal ActualAmount { get; set; }
        public decimal Variance => BudgetAmount - ActualAmount;
        public double PercentageUsed => BudgetAmount > 0 ? (double)(ActualAmount / BudgetAmount * 100) : 0;
        public bool IsOverBudget => ActualAmount > BudgetAmount && BudgetAmount > 0;
    }

    public class BudgetDashboardResult
    {
        public BudgetDashboardSummary Summary { get; set; } = new();
        public List<BudgetCategoryStat> Categories { get; set; } = new();
        public List<BudgetMonthlyTrend> MonthlyTrends { get; set; } = new();
        public List<BudgetAccountRow> TopAccounts { get; set; } = new();
        public List<FinancialYearOption> FinancialYears { get; set; } = new();
        public List<MasterHeadOption> MasterHeads { get; set; } = new();
    }
}

