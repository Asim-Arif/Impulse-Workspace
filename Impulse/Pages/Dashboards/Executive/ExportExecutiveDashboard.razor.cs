using Microsoft.AspNetCore.Components;
using System.Collections.Generic;

namespace Impulse.Pages.Dashboards.Executive
{
    public partial class ExportExecutiveDashboard : ComponentBase
    {
        // ── Filter & Navigation State ─────────────────────────────
        protected string SelectedPeriod { get; set; } = "FY 2025-2026";
        protected bool IsLoading { get; set; } = false;

        // ── Left Rail KPI Metrics ─────────────────────────────────
        protected int SalesActual { get; set; } = 16461;
        protected int SalesTarget { get; set; } = 15907;
        protected int TargetAchievementPercent { get; set; } = 103;
        protected int SalesLastYear { get; set; } = 11545;
        protected int GrossProfit { get; set; } = 5868;
        protected int GrossProfitMarginPercent { get; set; } = 36;
        protected int YoyGrowthPercent { get; set; } = 43;

        // ── Visual Data Models ────────────────────────────────────
        public class ProductCategorySales
        {
            public string ProductName { get; set; } = "";
            public int ActualSales { get; set; }
            public int TargetSales { get; set; }
            public bool IsAboveTarget => ActualSales >= TargetSales;
        }

        public class MonthlyTargetComparison
        {
            public string Month { get; set; } = "";
            public int Actual { get; set; }
            public int Target { get; set; }
            public double AchievementRatio => Target > 0 ? (double)Actual / Target : 0;
            public string StatusColor => AchievementRatio >= 1.0 ? "#00A88F" : (AchievementRatio >= 0.90 ? "#E5A823" : "#E74C3C");
        }

        public class RegionalSalesItem
        {
            public string Region { get; set; } = "";
            public int Actual { get; set; }
            public int Target { get; set; }
        }

        public class DonutSliceItem
        {
            public string Label { get; set; } = "";
            public int Amount { get; set; }
            public int Percentage { get; set; }
            public string Color { get; set; } = "";
        }

        public class MonthlyGrowthItem
        {
            public string Month { get; set; } = "";
            public int GrowthPercent { get; set; }
            public bool IsPositive => GrowthPercent >= 0;
        }

        public class ItemRankingItem
        {
            public string RegionOrGroup { get; set; } = "";
            public int Qty { get; set; }
            public int Value { get; set; }
        }

        public class TimeSeriesPoint
        {
            public string MonthLabel { get; set; } = "";
            public int DemandQty { get; set; }
            public int SupplyQty { get; set; }
            public int DemandValue { get; set; }
            public int SupplyValue { get; set; }
        }

        // ── Collections ───────────────────────────────────────────
        protected List<ProductCategorySales> ProductSalesList { get; set; } = new();
        protected List<MonthlyTargetComparison> MonthlySalesList { get; set; } = new();
        protected List<RegionalSalesItem> RegionalSalesList { get; set; } = new();
        protected List<DonutSliceItem> MainCategoriesList { get; set; } = new();
        protected List<DonutSliceItem> CustomerSegmentsList { get; set; } = new();
        protected List<MonthlyGrowthItem> MonthlyGrowthList { get; set; } = new();
        protected List<ItemRankingItem> TopItemsList { get; set; } = new();
        protected List<TimeSeriesPoint> DemandSupplyTimeline { get; set; } = new();

        protected override void OnInitialized()
        {
            LoadMockData();
        }

        protected void LoadMockData()
        {
            // 1. Sales by Product/Category
            ProductSalesList = new List<ProductCategorySales>
            {
                new ProductCategorySales { ProductName = "Product 1", ActualSales = 5597, TargetSales = 5000 },
                new ProductCategorySales { ProductName = "Product 2", ActualSales = 4263, TargetSales = 4000 },
                new ProductCategorySales { ProductName = "Product 3", ActualSales = 3405, TargetSales = 3200 },
                new ProductCategorySales { ProductName = "Product 4", ActualSales = 2379, TargetSales = 2600 },
                new ProductCategorySales { ProductName = "Product 5", ActualSales = 817, TargetSales = 1000 }
            };

            // 2. Actual Sales vs Target (12 Months)
            MonthlySalesList = new List<MonthlyTargetComparison>
            {
                new MonthlyTargetComparison { Month = "JUL", Actual = 1210, Target = 1200 },
                new MonthlyTargetComparison { Month = "AUG", Actual = 1550, Target = 1750 },
                new MonthlyTargetComparison { Month = "SEP", Actual = 1676, Target = 1650 },
                new MonthlyTargetComparison { Month = "OCT", Actual = 1527, Target = 1500 },
                new MonthlyTargetComparison { Month = "NOV", Actual = 962, Target = 1300 },
                new MonthlyTargetComparison { Month = "DEC", Actual = 957, Target = 1550 },
                new MonthlyTargetComparison { Month = "JAN", Actual = 1190, Target = 1300 },
                new MonthlyTargetComparison { Month = "FEB", Actual = 1778, Target = 1700 },
                new MonthlyTargetComparison { Month = "MAR", Actual = 1388, Target = 1350 },
                new MonthlyTargetComparison { Month = "APR", Actual = 1450, Target = 1430 },
                new MonthlyTargetComparison { Month = "MAY", Actual = 1550, Target = 1800 },
                new MonthlyTargetComparison { Month = "JUN", Actual = 1223, Target = 1200 }
            };

            // 3. Sales by Region
            RegionalSalesList = new List<RegionalSalesItem>
            {
                new RegionalSalesItem { Region = "East", Actual = 5567, Target = 5761 },
                new RegionalSalesItem { Region = "West", Actual = 3977, Target = 4115 },
                new RegionalSalesItem { Region = "South", Actual = 4772, Target = 4938 },
                new RegionalSalesItem { Region = "North", Actual = 1591, Target = 1646 }
            };

            // 4. Main Category Donut
            MainCategoriesList = new List<DonutSliceItem>
            {
                new DonutSliceItem { Label = "Distributor", Amount = 5761, Percentage = 35, Color = "#1E88E5" },
                new DonutSliceItem { Label = "Dealer", Amount = 4115, Percentage = 25, Color = "#42A5F5" },
                new DonutSliceItem { Label = "Retail", Amount = 3292, Percentage = 20, Color = "#FFB300" },
                new DonutSliceItem { Label = "Corporate", Amount = 2469, Percentage = 15, Color = "#FB8C00" },
                new DonutSliceItem { Label = "Online", Amount = 823, Percentage = 5, Color = "#00897B" }
            };

            // 5. Sales by Customers Donut
            CustomerSegmentsList = new List<DonutSliceItem>
            {
                new DonutSliceItem { Label = "Distributor", Amount = 5761, Percentage = 35, Color = "#00A88F" },
                new DonutSliceItem { Label = "Dealer", Amount = 4115, Percentage = 25, Color = "#1E88E5" },
                new DonutSliceItem { Label = "Retail", Amount = 3292, Percentage = 20, Color = "#5E35B1" },
                new DonutSliceItem { Label = "Corporate", Amount = 2469, Percentage = 15, Color = "#FB8C00" },
                new DonutSliceItem { Label = "Online", Amount = 823, Percentage = 5, Color = "#43A047" }
            };

            // 6. Sales Growth (% by month)
            MonthlyGrowthList = new List<MonthlyGrowthItem>
            {
                new MonthlyGrowthItem { Month = "JUL", GrowthPercent = 5 },
                new MonthlyGrowthItem { Month = "AUG", GrowthPercent = 28 },
                new MonthlyGrowthItem { Month = "SEP", GrowthPercent = 8 },
                new MonthlyGrowthItem { Month = "OCT", GrowthPercent = -9 },
                new MonthlyGrowthItem { Month = "NOV", GrowthPercent = -37 },
                new MonthlyGrowthItem { Month = "DEC", GrowthPercent = -1 },
                new MonthlyGrowthItem { Month = "JAN", GrowthPercent = 24 },
                new MonthlyGrowthItem { Month = "FEB", GrowthPercent = 49 },
                new MonthlyGrowthItem { Month = "MAR", GrowthPercent = -22 },
                new MonthlyGrowthItem { Month = "APR", GrowthPercent = 4 },
                new MonthlyGrowthItem { Month = "MAY", GrowthPercent = 7 },
                new MonthlyGrowthItem { Month = "JUN", GrowthPercent = -21 }
            };

            // 7. Top Items Rankings
            TopItemsList = new List<ItemRankingItem>
            {
                new ItemRankingItem { RegionOrGroup = "East", Qty = 5761, Value = 5567 },
                new ItemRankingItem { RegionOrGroup = "West", Qty = 4115, Value = 3977 },
                new ItemRankingItem { RegionOrGroup = "South", Qty = 4938, Value = 4772 },
                new ItemRankingItem { RegionOrGroup = "North", Qty = 1646, Value = 1591 }
            };

            // 8. Demand & Supply Timeline (Aug 2025 to Aug 2026)
            DemandSupplyTimeline = new List<TimeSeriesPoint>
            {
                new TimeSeriesPoint { MonthLabel = "Aug 2025", DemandQty = 500, SupplyQty = 480, DemandValue = 500, SupplyValue = 480 },
                new TimeSeriesPoint { MonthLabel = "Sep 2025", DemandQty = 400, SupplyQty = 120, DemandValue = 400, SupplyValue = 120 },
                new TimeSeriesPoint { MonthLabel = "Oct 2025", DemandQty = 500, SupplyQty = 120, DemandValue = 500, SupplyValue = 120 },
                new TimeSeriesPoint { MonthLabel = "Nov 2025", DemandQty = 600, SupplyQty = 400, DemandValue = 600, SupplyValue = 400 },
                new TimeSeriesPoint { MonthLabel = "Dec 2025", DemandQty = 700, SupplyQty = 440, DemandValue = 700, SupplyValue = 440 },
                new TimeSeriesPoint { MonthLabel = "Jan 2026", DemandQty = 800, SupplyQty = 480, DemandValue = 800, SupplyValue = 480 },
                new TimeSeriesPoint { MonthLabel = "Feb 2026", DemandQty = 400, SupplyQty = 200, DemandValue = 400, SupplyValue = 200 },
                new TimeSeriesPoint { MonthLabel = "Mar 2026", DemandQty = 200, SupplyQty = 80, DemandValue = 200, SupplyValue = 80 },
                new TimeSeriesPoint { MonthLabel = "Apr 2026", DemandQty = 100, SupplyQty = 90, DemandValue = 100, SupplyValue = 90 },
                new TimeSeriesPoint { MonthLabel = "May 2026", DemandQty = 700, SupplyQty = 360, DemandValue = 700, SupplyValue = 360 },
                new TimeSeriesPoint { MonthLabel = "Jun 2026", DemandQty = 500, SupplyQty = 540, DemandValue = 500, SupplyValue = 540 },
                new TimeSeriesPoint { MonthLabel = "Jul 2026", DemandQty = 700, SupplyQty = 420, DemandValue = 700, SupplyValue = 420 },
                new TimeSeriesPoint { MonthLabel = "Aug 2026", DemandQty = 500, SupplyQty = 300, DemandValue = 500, SupplyValue = 300 }
            };
        }

        protected void RefreshData()
        {
            IsLoading = true;
            LoadMockData();
            IsLoading = false;
        }
    }
}
