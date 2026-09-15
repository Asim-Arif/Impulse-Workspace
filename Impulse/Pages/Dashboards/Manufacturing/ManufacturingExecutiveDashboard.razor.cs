using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Impulse.Pages.Dashboards.Manufacturing
{
    public partial class ManufacturingExecutiveDashboard : ComponentBase
    {
        // ── Filter State ──────────────────────────────────────────
        protected string SelectedYear { get; set; } = "2024";
        protected string SelectedDepartment { get; set; } = "All";
        protected string SelectedPeriod { get; set; } = "YTD";
        protected bool IsLoading { get; set; } = false;

        // ── Top KPI Metrics ───────────────────────────────────────
        protected decimal UnitsYtdActual { get; set; } = 4282974m;
        protected decimal UnitsYtdTarget { get; set; } = 4500000m;
        protected decimal UnitsYtdVariancePercent { get; set; } = -6m;

        protected int UnitsLostTotal { get; set; } = 476;
        protected decimal PlantProductivityActual { get; set; } = 77.12m;
        protected decimal PlantProductivityTarget { get; set; } = 75.00m;
        protected decimal PlantProductivityVariance { get; set; } = 4.0m;

        protected int UnitsReworkTotal { get; set; } = 476;

        // ── Visual Data Models ────────────────────────────────────
        public class HorizontalBarMetric
        {
            public string Category { get; set; } = "";
            public double MachineA { get; set; }
            public double MachineB { get; set; }
            public double MachineC { get; set; }
            public double Total => MachineA + MachineB + MachineC;
        }

        public class RadialCauseItem
        {
            public string Label { get; set; } = "";
            public int Value { get; set; }
            public double Percentage { get; set; }
            public string Color { get; set; } = "";
            public string PositionClass { get; set; } = "";
        }

        public class MachineProductivityItem
        {
            public string Machine { get; set; } = "";
            public double OutputQty { get; set; }
            public double EfficiencyRate { get; set; }
            public double CapacityUtil { get; set; }
        }

        public class MakerProcessItem
        {
            public string ProcessName { get; set; } = "";
            public double IssuedQty { get; set; }
            public double PtcClosedQty { get; set; }
            public double BalancePending { get; set; }
        }

        public class DepartmentSummaryItem
        {
            public string DepartmentName { get; set; } = "";
            public string IssuedProcess { get; set; } = "";
            public int ActiveWorkers { get; set; }
            public double Efficiency { get; set; }
            public string Status { get; set; } = "Normal";
        }

        // ── Dataset Collections ───────────────────────────────────
        protected List<HorizontalBarMetric> OverallProductivityYtdList { get; set; } = new();
        protected List<HorizontalBarMetric> SecondaryProductivityList { get; set; } = new();
        protected List<RadialCauseItem> LostCausesList { get; set; } = new();
        protected List<RadialCauseItem> ReworkCausesList { get; set; } = new();
        protected List<MachineProductivityItem> MachineProductivityList { get; set; } = new();
        protected List<MakerProcessItem> MakerProcessList { get; set; } = new();
        protected List<DepartmentSummaryItem> DepartmentSummaryList { get; set; } = new();

        protected override void OnInitialized()
        {
            LoadMockData();
        }

        protected void LoadMockData()
        {
            // 1. Overall Productivity YTD (Horizontal bar comparison)
            OverallProductivityYtdList = new List<HorizontalBarMetric>
            {
                new HorizontalBarMetric { Category = "Line A (Forging)", MachineA = 94000, MachineB = 112000, MachineC = 125000 },
                new HorizontalBarMetric { Category = "Line B (Machining)", MachineA = 88000, MachineB = 98000, MachineC = 118000 },
                new HorizontalBarMetric { Category = "Line C (Grinding)", MachineA = 91000, MachineB = 105000, MachineC = 122000 }
            };

            SecondaryProductivityList = new List<HorizontalBarMetric>
            {
                new HorizontalBarMetric { Category = "Stitching & Assembly", MachineA = 78, MachineB = 84, MachineC = 92 },
                new HorizontalBarMetric { Category = "Surface Treatment", MachineA = 65, MachineB = 72, MachineC = 88 },
                new HorizontalBarMetric { Category = "Final Inspection", MachineA = 82, MachineB = 89, MachineC = 95 }
            };

            // 2. Concentric Radial Causes for Lost Units (476 units)
            LostCausesList = new List<RadialCauseItem>
            {
                new RadialCauseItem { Label = "Tooling Error", Value = 168, Percentage = 35.3, Color = "#8BAE26", PositionClass = "top-left" },
                new RadialCauseItem { Label = "Opener / Operator Damage", Value = 142, Percentage = 29.8, Color = "#00A88F", PositionClass = "top-right" },
                new RadialCauseItem { Label = "Physical Damage", Value = 98, Percentage = 20.6, Color = "#20645B", PositionClass = "bottom-left" },
                new RadialCauseItem { Label = "Other Causes", Value = 68, Percentage = 14.3, Color = "#5C946E", PositionClass = "bottom-right" }
            };

            // 3. Concentric Radial Causes for Rework Units (476 units)
            ReworkCausesList = new List<RadialCauseItem>
            {
                new RadialCauseItem { Label = "Dimensional Tolerance", Value = 185, Percentage = 38.9, Color = "#8BAE26", PositionClass = "top-left" },
                new RadialCauseItem { Label = "Surface Finish Flaw", Value = 135, Percentage = 28.4, Color = "#00A88F", PositionClass = "top-right" },
                new RadialCauseItem { Label = "Heat Treatment Re-run", Value = 89, Percentage = 18.7, Color = "#20645B", PositionClass = "bottom-left" },
                new RadialCauseItem { Label = "Coating / Deburr", Value = 67, Percentage = 14.0, Color = "#5C946E", PositionClass = "bottom-right" }
            };

            // 4. Productivity by Machine / Workstation (Grouped vertical bars)
            MachineProductivityList = new List<MachineProductivityItem>
            {
                new MachineProductivityItem { Machine = "M-01", OutputQty = 92, EfficiencyRate = 88, CapacityUtil = 84 },
                new MachineProductivityItem { Machine = "M-02", OutputQty = 85, EfficiencyRate = 79, CapacityUtil = 81 },
                new MachineProductivityItem { Machine = "M-03", OutputQty = 96, EfficiencyRate = 92, CapacityUtil = 89 },
                new MachineProductivityItem { Machine = "M-04", OutputQty = 89, EfficiencyRate = 83, CapacityUtil = 86 },
                new MachineProductivityItem { Machine = "M-05", OutputQty = 94, EfficiencyRate = 91, CapacityUtil = 87 },
                new MachineProductivityItem { Machine = "M-06", OutputQty = 91, EfficiencyRate = 86, CapacityUtil = 85 },
                new MachineProductivityItem { Machine = "M-07", OutputQty = 88, EfficiencyRate = 82, CapacityUtil = 80 },
                new MachineProductivityItem { Machine = "M-08", OutputQty = 95, EfficiencyRate = 93, CapacityUtil = 90 }
            };

            // 5. Maker Balance & PTC Closed Process (Purchase & Process Closed)
            MakerProcessList = new List<MakerProcessItem>
            {
                new MakerProcessItem { ProcessName = "Milling", IssuedQty = 48, PtcClosedQty = 42, BalancePending = 6 },
                new MakerProcessItem { ProcessName = "Turning", IssuedQty = 35, PtcClosedQty = 31, BalancePending = 4 },
                new MakerProcessItem { ProcessName = "Broaching", IssuedQty = 62, PtcClosedQty = 54, BalancePending = 8 },
                new MakerProcessItem { ProcessName = "Grinding", IssuedQty = 28, PtcClosedQty = 25, BalancePending = 3 },
                new MakerProcessItem { ProcessName = "Hardening", IssuedQty = 51, PtcClosedQty = 46, BalancePending = 5 },
                new MakerProcessItem { ProcessName = "Polishing", IssuedQty = 39, PtcClosedQty = 35, BalancePending = 4 },
                new MakerProcessItem { ProcessName = "Ultrasonic Clean", IssuedQty = 44, PtcClosedQty = 41, BalancePending = 3 },
                new MakerProcessItem { ProcessName = "Passivation", IssuedQty = 30, PtcClosedQty = 27, BalancePending = 3 }
            };

            // 6. Department-wise productivity & Maker Performance summary
            DepartmentSummaryList = new List<DepartmentSummaryItem>
            {
                new DepartmentSummaryItem { DepartmentName = "Forging & Stamping", IssuedProcess = "Hot Forging / Die Cut", ActiveWorkers = 24, Efficiency = 88.5, Status = "Optimal" },
                new DepartmentSummaryItem { DepartmentName = "CNC Precision Machining", IssuedProcess = "Milling & Turning", ActiveWorkers = 38, Efficiency = 91.2, Status = "Optimal" },
                new DepartmentSummaryItem { DepartmentName = "Heat Treatment & Hardening", IssuedProcess = "Vacuum Furnace", ActiveWorkers = 14, Efficiency = 76.8, Status = "Attention" },
                new DepartmentSummaryItem { DepartmentName = "Finishing & Deburring", IssuedProcess = "Electro-polish / Tumbling", ActiveWorkers = 29, Efficiency = 82.4, Status = "Normal" },
                new DepartmentSummaryItem { DepartmentName = "Maker Vendor Network", IssuedProcess = "Jobber Outwork / PTC", ActiveWorkers = 45, Efficiency = 84.1, Status = "Normal" }
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
