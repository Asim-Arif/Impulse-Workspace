using Microsoft.AspNetCore.Components;
using System.Collections.Generic;

namespace Impulse.Pages.Dashboards.Executive
{
    public partial class ExecutiveDashboardsHub : ComponentBase
    {
        public class ExecutiveDashboardCard
        {
            public string Title { get; set; } = "";
            public string Subtitle { get; set; } = "";
            public string Route { get; set; } = "";
            public string IconClass { get; set; } = "";
            public string AccentColor { get; set; } = "#8BAE26";
            public string Status { get; set; } = "Active"; // Active, Planned, InDevelopment
            public List<string> Highlights { get; set; } = new();
            public List<string> Tags { get; set; } = new();
        }

        protected List<ExecutiveDashboardCard> DashboardsList { get; set; } = new();
        protected string ActiveFilter { get; set; } = "All";

        protected override void OnInitialized()
        {
            DashboardsList = new List<ExecutiveDashboardCard>
            {
                new ExecutiveDashboardCard
                {
                    Title = "Product Manufacturing Dashboard",
                    Subtitle = "Production rate, overall productivity, scrap/lost units causes, rework distribution, machine productivity, and maker balance summary.",
                    Route = "/dashboards/manufacturing-executive",
                    IconClass = "fas fa-industry",
                    AccentColor = "#8BAE26",
                    Status = "Active",
                    Highlights = new List<string> { "4.28M Units YTD", "77.12% Plant Productivity", "Concentric Radial Gauges", "Maker Process Reconciliations" },
                    Tags = new List<string> { "Power BI", "Shop Floor", "Quality & Scrap", "Maker Balance" }
                },
                new ExecutiveDashboardCard
                {
                    Title = "Financial & Treasury Executive",
                    Subtitle = "Cash in bank, 6-month burn rate, salary vs expenses trend, solvency runway, debtor aging, vendor disbursements, and maker bills.",
                    Route = "/dashboards/finance-executive",
                    IconClass = "fas fa-vault",
                    AccentColor = "#107C41",
                    Status = "Active",
                    Highlights = new List<string> { "$616.5K Cash in Bank", "10.8mth Solvency Runway", "Top 20 Vendor & Maker Bills", "Debtors Aging Risk" },
                    Tags = new List<string> { "Finance", "Treasury", "Power BI", "Runway" }
                },
                new ExecutiveDashboardCard
                {
                    Title = "Export & Commercial Dashboard",
                    Subtitle = "Sales vs target, 12-month achievement thresholds, regional performance, category & customer distributions, MoM growth, and demand/supply trends.",
                    Route = "/dashboards/export-executive",
                    IconClass = "fas fa-plane-departure",
                    AccentColor = "#00A88F",
                    Status = "Active",
                    Highlights = new List<string> { "16,461 Sales (YoY +43%)", "103% Target Achievement", "Regional & Category Donut", "Demand & Supply (800K)" },
                    Tags = new List<string> { "Exports", "Power BI", "Sales Target", "Commercial" }
                },
                new ExecutiveDashboardCard
                {
                    Title = "Workforce & Employee Executive",
                    Subtitle = "Department compensation, multi-year headcounts, salary by location heatmap matrix, bonus distributions, and zero-absence top performers.",
                    Route = "/dashboards/employee-executive",
                    IconClass = "fas fa-users-gear",
                    AccentColor = "#2E7D32",
                    Status = "Active",
                    Highlights = new List<string> { "$584K Base Salary", "Location Matrix Matrix", "Compensation Stacks", "Top 20 Zero-Leave Stars" },
                    Tags = new List<string> { "HR & Payroll", "Power BI", "Headcount", "Salary Heatmap" }
                }
            };
        }
    }
}
