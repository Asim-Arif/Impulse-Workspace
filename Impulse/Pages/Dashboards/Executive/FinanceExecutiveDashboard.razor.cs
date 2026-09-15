using Microsoft.AspNetCore.Components;
using System.Collections.Generic;

namespace Impulse.Pages.Dashboards.Executive
{
    public partial class FinanceExecutiveDashboard : ComponentBase
    {
        protected string SelectedPeriod { get; set; } = "Past 6 Months";
        protected bool IsLoading { get; set; } = false;

        // ── Top Macro KPI Metrics ─────────────────────────────────
        protected decimal CashInBankCurrent { get; set; } = 616.5m; // In Thousands
        protected decimal BurnRateAvg { get; set; } = 63.9m; // In Thousands
        protected decimal SolvencyRunwayMonths { get; set; } = 10.8m;
        protected decimal DebtorsTotal { get; set; } = 345.0m;
        protected decimal DebtorsOver45Days { get; set; } = 125.0m;

        // ── Models ────────────────────────────────────────────────
        public class MonthlyCashBar
        {
            public string Month { get; set; } = "";
            public double CashAmountK { get; set; }
        }

        public class MonthlyBurnPoint
        {
            public string Month { get; set; } = "";
            public double BurnAmountK { get; set; }
        }

        public class MonthlyExpenseLine
        {
            public string Month { get; set; } = "";
            public double SalaryK { get; set; }
            public double ExpensesK { get; set; }
            public double FixedCostsK { get; set; }
        }

        public class SolvencyIncomeOutgoing
        {
            public string Month { get; set; } = "";
            public double IncomeK { get; set; }
            public double OutgoingsK { get; set; }
        }

        public class MonthlyDebtorBar
        {
            public string Month { get; set; } = "";
            public double DebtorsK { get; set; }
        }

        public class ExpenseShareSlice
        {
            public string Category { get; set; } = "";
            public double Percentage { get; set; }
            public string Color { get; set; } = "";
        }

        public class BankBalanceSlice
        {
            public string BankName { get; set; } = "";
            public double BalanceK { get; set; }
            public double Percentage { get; set; }
            public string Color { get; set; } = "";
        }

        public class VendorPaymentItem
        {
            public string VendorName { get; set; } = "";
            public string Category { get; set; } = "";
            public decimal PaidAmount { get; set; }
            public decimal PendingAmount { get; set; }
        }

        public class MakerPaymentItem
        {
            public string MakerName { get; set; } = "";
            public string ProcessName { get; set; } = "";
            public decimal BilledAmount { get; set; }
            public decimal PaidAmount { get; set; }
            public decimal BalanceDue { get; set; }
        }

        public class ReceivableAgingItem
        {
            public string CustomerName { get; set; } = "";
            public string Country { get; set; } = "";
            public decimal TotalReceivable { get; set; }
            public int OverdueDays { get; set; }
            public bool IsOver45Days => OverdueDays > 45;
        }

        // ── Collections ───────────────────────────────────────────
        protected List<MonthlyCashBar> CashHistoryList { get; set; } = new();
        protected List<MonthlyBurnPoint> BurnHistoryList { get; set; } = new();
        protected List<MonthlyExpenseLine> ExpensesTrendList { get; set; } = new();
        protected List<SolvencyIncomeOutgoing> SolvencyHistoryList { get; set; } = new();
        protected List<MonthlyDebtorBar> DebtorsTrendList { get; set; } = new();
        protected List<ExpenseShareSlice> ExpenseShareList { get; set; } = new();
        protected List<BankBalanceSlice> BankAccountsList { get; set; } = new();
        protected List<VendorPaymentItem> TopVendorsList { get; set; } = new();
        protected List<MakerPaymentItem> TopMakersList { get; set; } = new();
        protected List<ReceivableAgingItem> TopReceivablesList { get; set; } = new();

        protected override void OnInitialized()
        {
            LoadMockData();
        }

        protected void LoadMockData()
        {
            // 1. Cash In Bank (Jan - June)
            CashHistoryList = new List<MonthlyCashBar>
            {
                new MonthlyCashBar { Month = "Jan", CashAmountK = 720 },
                new MonthlyCashBar { Month = "Feb", CashAmountK = 680 },
                new MonthlyCashBar { Month = "March", CashAmountK = 540 },
                new MonthlyCashBar { Month = "April", CashAmountK = 510 },
                new MonthlyCashBar { Month = "May", CashAmountK = 460 },
                new MonthlyCashBar { Month = "June", CashAmountK = 616.5 }
            };

            // 2. Burn Rate (Jan - June)
            BurnHistoryList = new List<MonthlyBurnPoint>
            {
                new MonthlyBurnPoint { Month = "Jan", BurnAmountK = 38 },
                new MonthlyBurnPoint { Month = "Feb", BurnAmountK = 44 },
                new MonthlyBurnPoint { Month = "March", BurnAmountK = 112 },
                new MonthlyBurnPoint { Month = "April", BurnAmountK = 28 },
                new MonthlyBurnPoint { Month = "May", BurnAmountK = 58 },
                new MonthlyBurnPoint { Month = "June", BurnAmountK = 80 }
            };

            // 3. Expenses Trend (Salary, Expenses, Fixed)
            ExpensesTrendList = new List<MonthlyExpenseLine>
            {
                new MonthlyExpenseLine { Month = "Jan", SalaryK = 195, FixedCostsK = 24, ExpensesK = 90 },
                new MonthlyExpenseLine { Month = "Feb", SalaryK = 195, FixedCostsK = 24, ExpensesK = 90 },
                new MonthlyExpenseLine { Month = "March", SalaryK = 215, FixedCostsK = 25, ExpensesK = 145 },
                new MonthlyExpenseLine { Month = "April", SalaryK = 215, FixedCostsK = 25, ExpensesK = 70 },
                new MonthlyExpenseLine { Month = "May", SalaryK = 228, FixedCostsK = 27, ExpensesK = 70 },
                new MonthlyExpenseLine { Month = "June", SalaryK = 228, FixedCostsK = 27, ExpensesK = 115 }
            };

            // 4. Solvency Income vs Outgoings
            SolvencyHistoryList = new List<SolvencyIncomeOutgoing>
            {
                new SolvencyIncomeOutgoing { Month = "Jan", IncomeK = 210, OutgoingsK = 245 },
                new SolvencyIncomeOutgoing { Month = "Feb", IncomeK = 215, OutgoingsK = 255 },
                new SolvencyIncomeOutgoing { Month = "March", IncomeK = 225, OutgoingsK = 310 },
                new SolvencyIncomeOutgoing { Month = "April", IncomeK = 265, OutgoingsK = 270 },
                new SolvencyIncomeOutgoing { Month = "May", IncomeK = 220, OutgoingsK = 275 },
                new SolvencyIncomeOutgoing { Month = "June", IncomeK = 240, OutgoingsK = 295 }
            };

            // 5. Debtors Trend
            DebtorsTrendList = new List<MonthlyDebtorBar>
            {
                new MonthlyDebtorBar { Month = "Jan", DebtorsK = 230 },
                new MonthlyDebtorBar { Month = "Feb", DebtorsK = 230 },
                new MonthlyDebtorBar { Month = "March", DebtorsK = 215 },
                new MonthlyDebtorBar { Month = "April", DebtorsK = 235 },
                new MonthlyDebtorBar { Month = "May", DebtorsK = 290 },
                new MonthlyDebtorBar { Month = "June", DebtorsK = 345 }
            };

            // 6. Expense Share Donut
            ExpenseShareList = new List<ExpenseShareSlice>
            {
                new ExpenseShareSlice { Category = "Coastal Shipping & Freight", Percentage = 26.0, Color = "#00897B" },
                new ExpenseShareSlice { Category = "City Construction & Plant", Percentage = 19.3, Color = "#1E88E5" },
                new ExpenseShareSlice { Category = "TechAdvantage Software / ERP", Percentage = 15.1, Color = "#5E35B1" },
                new ExpenseShareSlice { Category = "Urban Apparel & Safety Uniforms", Percentage = 12.7, Color = "#00ACC1" },
                new ExpenseShareSlice { Category = "Global Exports Forwarding", Percentage = 10.6, Color = "#FB8C00" },
                new ExpenseShareSlice { Category = "Green Gardens & Facilities", Percentage = 6.3, Color = "#7CB342" },
                new ExpenseShareSlice { Category = "Innovative Tech & Tooling", Percentage = 6.2, Color = "#E53935" },
                new ExpenseShareSlice { Category = "Solar Solutions Power", Percentage = 3.8, Color = "#FDD835" }
            };

            // 7. Bank Accounts Balance Donut
            BankAccountsList = new List<BankBalanceSlice>
            {
                new BankBalanceSlice { BankName = "Standard Chartered (Export)", BalanceK = 258.9, Percentage = 42.0, Color = "#1E88E5" },
                new BankBalanceSlice { BankName = "Habib Bank Limited (HBL Ops)", BalanceK = 172.6, Percentage = 28.0, Color = "#00897B" },
                new BankBalanceSlice { BankName = "Meezan Islamic Bank (FCY)", BalanceK = 110.9, Percentage = 18.0, Color = "#43A047" },
                new BankBalanceSlice { BankName = "MCB Bank (Disbursements)", BalanceK = 74.1, Percentage = 12.0, Color = "#FB8C00" }
            };

            // 8. Top 20 Vendors Names & Payment
            TopVendorsList = new List<VendorPaymentItem>
            {
                new VendorPaymentItem { VendorName = "Apex Steel & Alloy Mills", Category = "Raw Material (Steel)", PaidAmount = 145000, PendingAmount = 18500 },
                new VendorPaymentItem { VendorName = "National Gas & Vacuum Furnace", Category = "Utilities / Energy", PaidAmount = 88500, PendingAmount = 9200 },
                new VendorPaymentItem { VendorName = "Precision Die & Mold Works", Category = "Tooling & Consumables", PaidAmount = 74200, PendingAmount = 6400 },
                new VendorPaymentItem { VendorName = "Al-Madina Corrugated Packaging", Category = "Export Cartons", PaidAmount = 61800, PendingAmount = 4500 },
                new VendorPaymentItem { VendorName = "Surgical Grade Chemical Supply", Category = "Electro-polishing", PaidAmount = 52400, PendingAmount = 3800 },
                new VendorPaymentItem { VendorName = "Express Cargo & Freight Lines", Category = "Logistics", PaidAmount = 48900, PendingAmount = 7100 },
                new VendorPaymentItem { VendorName = "United Hardware & Fasteners", Category = "Spare Parts", PaidAmount = 36500, PendingAmount = 2900 },
                new VendorPaymentItem { VendorName = "Lubricants & Industrial Oils Ltd", Category = "Maintenance", PaidAmount = 29400, PendingAmount = 1800 }
            };

            // 9. Top 20 Makers Name & Payment (Outwork Jobbers)
            TopMakersList = new List<MakerPaymentItem>
            {
                new MakerPaymentItem { MakerName = "Master Forging Workshop (Maker-01)", ProcessName = "Hot Forging / Drop Hammer", BilledAmount = 168000, PaidAmount = 152000, BalanceDue = 16000 },
                new MakerPaymentItem { MakerName = "Iqbal Precision Grinding (Maker-04)", ProcessName = "Centerless Grinding", BilledAmount = 124500, PaidAmount = 114000, BalanceDue = 10500 },
                new MakerPaymentItem { MakerName = "Rehman Turning & CNC (Maker-07)", ProcessName = "Milling & Turning", BilledAmount = 118000, PaidAmount = 108000, BalanceDue = 10000 },
                new MakerPaymentItem { MakerName = "Naseer Heat Treatment (Maker-12)", ProcessName = "Annealing & Hardening", BilledAmount = 94200, PaidAmount = 88000, BalanceDue = 6200 },
                new MakerPaymentItem { MakerName = "Bismillah Mirror Polishing (Maker-15)", ProcessName = "Final Mirror Finishing", BilledAmount = 88500, PaidAmount = 82000, BalanceDue = 6500 },
                new MakerPaymentItem { MakerName = "Al-Razi Filing & Filing (Maker-19)", ProcessName = "Hand Serration & Filing", BilledAmount = 76000, PaidAmount = 71500, BalanceDue = 4500 },
                new MakerPaymentItem { MakerName = "Tariq Ultrasonic Cleaning (Maker-22)", ProcessName = "Degreasing & Ultrasonic", BilledAmount = 52000, PaidAmount = 49000, BalanceDue = 3000 }
            };

            // 10. Receivable Names & Overdue Amount
            TopReceivablesList = new List<ReceivableAgingItem>
            {
                new ReceivableAgingItem { CustomerName = "Apex Dental & Surgical Instruments", Country = "USA", TotalReceivable = 92400, OverdueDays = 54 },
                new ReceivableAgingItem { CustomerName = "MedTech International Italia S.R.L.", Country = "Italy", TotalReceivable = 78500, OverdueDays = 62 },
                new ReceivableAgingItem { CustomerName = "Central America Dental Supplies S.A.", Country = "Panama", TotalReceivable = 54100, OverdueDays = 49 },
                new ReceivableAgingItem { CustomerName = "Vanguard Medical Innovations LLC", Country = "USA", TotalReceivable = 46200, OverdueDays = 28 },
                new ReceivableAgingItem { CustomerName = "Lombardia Precision Tools S.p.A.", Country = "Italy", TotalReceivable = 38900, OverdueDays = 18 },
                new ReceivableAgingItem { CustomerName = "Equine Care & Dental Equipment", Country = "France", TotalReceivable = 34900, OverdueDays = 58 }
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
