using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;

namespace Impulse.Pages.Dashboards.Executive
{
    public partial class EmployeeExecutiveDashboard : ComponentBase
    {
        protected string SelectedPeriod { get; set; } = "All Years (2000-2008)";
        protected bool IsLoading { get; set; } = false;

        // ── 1. Salary and Bonus by Department Table ───────────────
        public class DeptSalaryBonusItem
        {
            public string Department { get; set; } = "";
            public int BaseSalary { get; set; }
            public int Bonus { get; set; }
            public string BaseColor { get; set; } = "";
            public string BonusColor { get; set; } = "";
        }

        // ── 2. Employee Count by Year (Stacked) ───────────────────
        public class YearDeptCountItem
        {
            public string Year { get; set; } = "";
            public int Accounting { get; set; }
            public int Administration { get; set; }
            public int CustomerSupport { get; set; }
            public int Finance { get; set; }
            public int HumanResources { get; set; }
            public int IT { get; set; }
            public int Marketing { get; set; }
            public int RD { get; set; }
            public int Sales { get; set; }
            public int Total => Accounting + Administration + CustomerSupport + Finance + HumanResources + IT + Marketing + RD + Sales;
        }

        // ── 3. Compensation Distribution ──────────────────────────
        public class DeptCompensationItem
        {
            public string Department { get; set; } = "";
            public double BaseSalaryMil { get; set; }
            public double BonusMil { get; set; }
            public double CommissionMil { get; set; }
            public double OvertimeMil { get; set; }
            public double TotalMil => BaseSalaryMil + BonusMil + CommissionMil + OvertimeMil;
        }

        // ── 4. Location Salary Matrix ─────────────────────────────
        public class LocationSalaryMatrixRow
        {
            public string Location { get; set; } = "";
            public Dictionary<string, int> YearValues { get; set; } = new();
        }

        // ── 5. Donut Slice ────────────────────────────────────────
        public class DeptDonutSlice
        {
            public string Department { get; set; } = "";
            public int Count { get; set; }
            public string Color { get; set; } = "";
        }

        // ── 6. Bonus by State/Dept and Year Heatmap ───────────────
        public class StateBonusHeatmapRow
        {
            public string StateCode { get; set; } = "";
            public Dictionary<string, int> YearBonus { get; set; } = new();
        }

        // ── 7. Outstanding Liabilities ────────────────────────────
        public class DeptOutstandingItem
        {
            public string Department { get; set; } = "";
            public double ShortTerm { get; set; }
            public double LongTerm { get; set; }
        }

        // ── Collections ───────────────────────────────────────────
        protected List<DeptSalaryBonusItem> DeptSalaryBonusList { get; set; } = new();
        protected int TotalBaseSalary => 584958;
        protected int TotalBonus => 49104;

        protected List<YearDeptCountItem> YearlyDeptCounts { get; set; } = new();
        protected List<DeptCompensationItem> CompensationDistributionList { get; set; } = new();
        protected List<LocationSalaryMatrixRow> LocationSalaryRows { get; set; } = new();
        protected Dictionary<string, int> MatrixGrandTotals { get; set; } = new();
        protected List<DeptDonutSlice> EmployeesByDeptDonut { get; set; } = new();
        protected List<DeptCompensationItem> TopAttendanceEmployees { get; set; } = new();
        protected List<StateBonusHeatmapRow> StateBonusHeatmapRows { get; set; } = new();
        protected List<DeptOutstandingItem> DeptOutstandingList { get; set; } = new();

        protected override void OnInitialized()
        {
            LoadMockData();
        }

        protected void LoadMockData()
        {
            // 1. Salary and Bonus by Department
            DeptSalaryBonusList = new List<DeptSalaryBonusItem>
            {
                new DeptSalaryBonusItem { Department = "Customer Support", BaseSalary = 72489, Bonus = 6851, BaseColor = "#2E7D32", BonusColor = "#1565C0" },
                new DeptSalaryBonusItem { Department = "IT", BaseSalary = 72810, Bonus = 6091, BaseColor = "#2E7D32", BonusColor = "#1976D2" },
                new DeptSalaryBonusItem { Department = "Administration", BaseSalary = 65731, Bonus = 5299, BaseColor = "#43A047", BonusColor = "#C0CA33" },
                new DeptSalaryBonusItem { Department = "R&D", BaseSalary = 65411, Bonus = 5001, BaseColor = "#43A047", BonusColor = "#D4E157" },
                new DeptSalaryBonusItem { Department = "Human Resources", BaseSalary = 62468, Bonus = 5055, BaseColor = "#66BB6A", BonusColor = "#FDD835" },
                new DeptSalaryBonusItem { Department = "Sales", BaseSalary = 62189, Bonus = 5167, BaseColor = "#66BB6A", BonusColor = "#FDD835" },
                new DeptSalaryBonusItem { Department = "Accounting", BaseSalary = 61953, Bonus = 4928, BaseColor = "#66BB6A", BonusColor = "#FFEE58" },
                new DeptSalaryBonusItem { Department = "Marketing", BaseSalary = 61317, Bonus = 5034, BaseColor = "#81C784", BonusColor = "#FFEE58" },
                new DeptSalaryBonusItem { Department = "Finance", BaseSalary = 59588, Bonus = 5088, BaseColor = "#A5D6A7", BonusColor = "#FFF59D" }
            };

            // 2. Multi-year Employee Count by Department (2000 - 2008)
            YearlyDeptCounts = new List<YearDeptCountItem>
            {
                new YearDeptCountItem { Year = "2000", Accounting = 2, Administration = 1, CustomerSupport = 1, Finance = 1, HumanResources = 1, IT = 1, Marketing = 1, RD = 1, Sales = 1 },
                new YearDeptCountItem { Year = "2001", Accounting = 3, Administration = 2, CustomerSupport = 3, Finance = 2, HumanResources = 2, IT = 3, Marketing = 2, RD = 2, Sales = 2 },
                new YearDeptCountItem { Year = "2002", Accounting = 4, Administration = 3, CustomerSupport = 4, Finance = 2, HumanResources = 3, IT = 4, Marketing = 2, RD = 3, Sales = 3 },
                new YearDeptCountItem { Year = "2003", Accounting = 5, Administration = 4, CustomerSupport = 5, Finance = 3, HumanResources = 4, IT = 5, Marketing = 3, RD = 4, Sales = 4 },
                new YearDeptCountItem { Year = "2004", Accounting = 6, Administration = 5, CustomerSupport = 6, Finance = 4, HumanResources = 5, IT = 6, Marketing = 4, RD = 5, Sales = 5 },
                new YearDeptCountItem { Year = "2005", Accounting = 7, Administration = 6, CustomerSupport = 7, Finance = 5, HumanResources = 6, IT = 8, Marketing = 5, RD = 6, Sales = 6 },
                new YearDeptCountItem { Year = "2006", Accounting = 8, Administration = 7, CustomerSupport = 9, Finance = 6, HumanResources = 7, IT = 10, Marketing = 6, RD = 8, Sales = 7 },
                new YearDeptCountItem { Year = "2007", Accounting = 9, Administration = 9, CustomerSupport = 11, Finance = 7, HumanResources = 8, IT = 12, Marketing = 7, RD = 9, Sales = 8 },
                new YearDeptCountItem { Year = "2008", Accounting = 10, Administration = 10, CustomerSupport = 12, Finance = 8, HumanResources = 9, IT = 14, Marketing = 8, RD = 10, Sales = 9 }
            };

            // 3. Compensation Distribution by Department
            CompensationDistributionList = new List<DeptCompensationItem>
            {
                new DeptCompensationItem { Department = "Accounting", BaseSalaryMil = 1.9, BonusMil = 0.2, CommissionMil = 0.1, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Administration", BaseSalaryMil = 2.4, BonusMil = 0.25, CommissionMil = 0.0, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Customer Support", BaseSalaryMil = 0.9, BonusMil = 0.15, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Finance", BaseSalaryMil = 1.6, BonusMil = 0.2, CommissionMil = 0.1, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Human Resources", BaseSalaryMil = 1.2, BonusMil = 0.15, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "IT", BaseSalaryMil = 3.2, BonusMil = 0.35, CommissionMil = 0.0, OvertimeMil = 0.1 },
                new DeptCompensationItem { Department = "Marketing", BaseSalaryMil = 1.5, BonusMil = 0.2, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "R&D", BaseSalaryMil = 2.9, BonusMil = 0.35, CommissionMil = 0.0, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Sales", BaseSalaryMil = 2.0, BonusMil = 0.3, CommissionMil = 1.8, OvertimeMil = 0.05 }
            };

            // 4. Salary by Location and Year
            LocationSalaryRows = new List<LocationSalaryMatrixRow>
            {
                new LocationSalaryMatrixRow
                {
                    Location = "Boston",
                    YearValues = new Dictionary<string, int>
                    {
                        { "2000", 352000 }, { "2001", 853000 }, { "2002", 1056000 }, { "2003", 1367000 },
                        { "2004", 1558000 }, { "2005", 1788000 }, { "2006", 2430990 }, { "2007", 2753519 }, { "2008", 3199200 }
                    }
                },
                new LocationSalaryMatrixRow
                {
                    Location = "Chicago",
                    YearValues = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 277000 }, { "2002", 285000 }, { "2003", 421000 },
                        { "2004", 356000 }, { "2005", 432000 }, { "2006", 416920 }, { "2007", 548816 }, { "2008", 698400 }
                    }
                },
                new LocationSalaryMatrixRow
                {
                    Location = "Los Angeles",
                    YearValues = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 129000 }, { "2002", 139000 }, { "2003", 347000 },
                        { "2004", 430000 }, { "2005", 452000 }, { "2006", 539260 }, { "2007", 704700 }, { "2008", 803900 }
                    }
                },
                new LocationSalaryMatrixRow
                {
                    Location = "Miami",
                    YearValues = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 0 }, { "2002", 0 }, { "2003", 154000 },
                        { "2004", 212500 }, { "2005", 260500 }, { "2006", 315350 }, { "2007", 421241 }, { "2008", 566800 }
                    }
                },
                new LocationSalaryMatrixRow
                {
                    Location = "New York",
                    YearValues = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 0 }, { "2002", 77000 }, { "2003", 190000 },
                        { "2004", 259000 }, { "2005", 305000 }, { "2006", 363170 }, { "2007", 358293 }, { "2008", 427700 }
                    }
                }
            };

            MatrixGrandTotals = new Dictionary<string, int>
            {
                { "2000", 352000 }, { "2001", 1259000 }, { "2002", 1557000 }, { "2003", 2479000 },
                { "2004", 2815500 }, { "2005", 3237500 }, { "2006", 4065700 }, { "2007", 4786571 }, { "2008", 5696000 }
            };

            // 5. Employees by Department Donut
            EmployeesByDeptDonut = new List<DeptDonutSlice>
            {
                new DeptDonutSlice { Department = "Accounting", Count = 50, Color = "#8BC34A" },
                new DeptDonutSlice { Department = "Administration", Count = 48, Color = "#03A9F4" },
                new DeptDonutSlice { Department = "Customer Support", Count = 56, Color = "#00BCD4" },
                new DeptDonutSlice { Department = "Finance", Count = 21, Color = "#37474F" },
                new DeptDonutSlice { Department = "Human Resources", Count = 41, Color = "#8D6E63" },
                new DeptDonutSlice { Department = "IT", Count = 29, Color = "#D84315" },
                new DeptDonutSlice { Department = "Marketing", Count = 59, Color = "#FF8F00" },
                new DeptDonutSlice { Department = "R&D", Count = 37, Color = "#00796B" },
                new DeptDonutSlice { Department = "Sales", Count = 61, Color = "#546E7A" }
            };

            // 6. Top 20 Employees with no absent/leave
            TopAttendanceEmployees = new List<DeptCompensationItem>
            {
                new DeptCompensationItem { Department = "Emp #1042 (Forging)", BaseSalaryMil = 1.8, BonusMil = 0.25, CommissionMil = 0.0, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Emp #1088 (CNC Milling)", BaseSalaryMil = 2.4, BonusMil = 0.3, CommissionMil = 0.0, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Emp #1115 (Assembly)", BaseSalaryMil = 0.85, BonusMil = 0.15, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Emp #1142 (Grinding)", BaseSalaryMil = 1.1, BonusMil = 0.2, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Emp #1201 (QC Lead)", BaseSalaryMil = 1.3, BonusMil = 0.2, CommissionMil = 0.05, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Emp #1255 (Heat Treat)", BaseSalaryMil = 2.2, BonusMil = 0.25, CommissionMil = 0.35, OvertimeMil = 0.05 },
                new DeptCompensationItem { Department = "Emp #1290 (Polishing)", BaseSalaryMil = 1.2, BonusMil = 0.15, CommissionMil = 0.0, OvertimeMil = 0.02 },
                new DeptCompensationItem { Department = "Emp #1334 (Tool Maker)", BaseSalaryMil = 2.3, BonusMil = 0.3, CommissionMil = 0.0, OvertimeMil = 0.1 },
                new DeptCompensationItem { Department = "Emp #1380 (Line Supv)", BaseSalaryMil = 1.6, BonusMil = 0.5, CommissionMil = 1.8, OvertimeMil = 0.05 }
            };

            // 7. Bonus by State and Year Heatmap
            StateBonusHeatmapRows = new List<StateBonusHeatmapRow>
            {
                new StateBonusHeatmapRow
                {
                    StateCode = "CA",
                    YearBonus = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 4613 }, { "2002", 3590 }, { "2003", 5920 },
                        { "2004", 6102 }, { "2005", 6632 }, { "2006", 5408 }, { "2007", 6858 }, { "2008", 7490 }
                    }
                },
                new StateBonusHeatmapRow
                {
                    StateCode = "FL",
                    YearBonus = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 0 }, { "2002", 0 }, { "2003", 2290 },
                        { "2004", 2623 }, { "2005", 3023 }, { "2006", 4484 }, { "2007", 4416 }, { "2008", 5405 }
                    }
                },
                new StateBonusHeatmapRow
                {
                    StateCode = "IL",
                    YearBonus = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 5155 }, { "2002", 5613 }, { "2003", 6217 },
                        { "2004", 5235 }, { "2005", 5954 }, { "2006", 6009 }, { "2007", 4361 }, { "2008", 5539 }
                    }
                },
                new StateBonusHeatmapRow
                {
                    StateCode = "MA",
                    YearBonus = new Dictionary<string, int>
                    {
                        { "2000", 5133 }, { "2001", 5452 }, { "2002", 5435 }, { "2003", 5200 },
                        { "2004", 5738 }, { "2005", 5146 }, { "2006", 5716 }, { "2007", 6007 }, { "2008", 5552 }
                    }
                },
                new StateBonusHeatmapRow
                {
                    StateCode = "NY",
                    YearBonus = new Dictionary<string, int>
                    {
                        { "2000", 0 }, { "2001", 0 }, { "2002", 4820 }, { "2003", 4855 },
                        { "2004", 4512 }, { "2005", 4608 }, { "2006", 4329 }, { "2007", 4643 }, { "2008", 5435 }
                    }
                }
            };

            // 8. Department wise outstanding liabilities
            DeptOutstandingList = new List<DeptOutstandingItem>
            {
                new DeptOutstandingItem { Department = "Accounting", ShortTerm = 2.1, LongTerm = 0.2 },
                new DeptOutstandingItem { Department = "Administration", ShortTerm = 2.5, LongTerm = 0.3 },
                new DeptOutstandingItem { Department = "Customer Support", ShortTerm = 0.9, LongTerm = 0.1 },
                new DeptOutstandingItem { Department = "Finance", ShortTerm = 1.7, LongTerm = 0.2 },
                new DeptOutstandingItem { Department = "Human Resources", ShortTerm = 1.3, LongTerm = 0.1 },
                new DeptOutstandingItem { Department = "IT", ShortTerm = 2.8, LongTerm = 0.35 },
                new DeptOutstandingItem { Department = "Marketing", ShortTerm = 1.4, LongTerm = 0.15 },
                new DeptOutstandingItem { Department = "R&D", ShortTerm = 2.6, LongTerm = 0.3 },
                new DeptOutstandingItem { Department = "Sales", ShortTerm = 1.8, LongTerm = 2.2 }
            };
        }

        protected string GetHeatmapBg(int val)
        {
            if (val == 0) return "#ffffff";
            if (val < 2500) return "#C8E6C9";
            if (val < 4000) return "#A5D6A7";
            if (val < 5000) return "#81C784";
            if (val < 6000) return "#4CAF50";
            if (val < 7000) return "#009688";
            return "#00796B";
        }

        protected string GetHeatmapTextColor(int val)
        {
            return val >= 4500 ? "#ffffff" : "#2c3e50";
        }

        protected void RefreshData()
        {
            IsLoading = true;
            LoadMockData();
            IsLoading = false;
        }
    }
}
