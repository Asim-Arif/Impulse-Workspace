using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Dashboard
{
    public class HRPillarData
    {
        public int TotalActiveEmployees { get; set; }
        public decimal TotalMonthlyPayroll { get; set; }
        public int TotalPresentEmployees { get; set; }
        public int TotalAbsentEmployees => Math.Max(0, TotalActiveEmployees - TotalPresentEmployees);
        public double AttendancePercentage => TotalActiveEmployees > 0 ? Math.Round(((double)TotalPresentEmployees / TotalActiveEmployees) * 100.0, 1) : 0;
        public decimal LongTermLoansBalance { get; set; }
        public decimal ShortTermAdvancesBalance { get; set; }
    }

    public class StockPillarData
    {
        public long FinishedStockQty { get; set; }
        public decimal FinishedStockValue { get; set; }

        public long SFStockQty { get; set; }
        public decimal SFStockValue { get; set; }

        public long ForgingStockQty { get; set; }
        public decimal ForgingStockValue { get; set; }

        public long OtherStockQty { get; set; }
        public decimal OtherStockValue { get; set; }

        public int TotalReWorkLots { get; set; }
        public long TotalReWorkQty { get; set; }

        public long TotalInventoryQty => FinishedStockQty + SFStockQty + ForgingStockQty + OtherStockQty;
        public decimal TotalInventoryValue => FinishedStockValue + SFStockValue + ForgingStockValue + OtherStockValue;
    }

    public class FinancialPillarData
    {
        public decimal CashInHand { get; set; }
        public decimal BankBalance { get; set; }
        public decimal AccountsReceivable { get; set; }
        public decimal AccountsPayable { get; set; }
        public decimal FixedAssetsCost { get; set; }
        public decimal NetLiquidCash => CashInHand + BankBalance;
    }

    public class MonthlyExportSalesTrendItem
    {
        public string MonthLabel { get; set; } = string.Empty;
        public int Year { get; set; }
        public int Month { get; set; }
        public int InvoiceCount { get; set; }
        public long TotalShippedQty { get; set; }
        public decimal TotalAmountPKR { get; set; }
        public decimal TotalAmountForeign { get; set; }
    }

    public class ExportPillarData
    {
        public long InHandQty { get; set; }
        public decimal InHandValue { get; set; }
        public decimal TotalSalesCurrentFY { get; set; }
        public decimal LifetimeTotalSales { get; set; }
        public decimal RebateReceivable { get; set; }
        public List<MonthlyExportSalesTrendItem> MonthlyTrend { get; set; } = new();
    }

    public class ProductionPillarData
    {
        public int RunningLotsCount { get; set; }
        public long RunningLotsQty { get; set; }

        public int MakerOrdersCount { get; set; }
        public long MakerOrdersQty { get; set; }

        public int ForgingOrdersCount { get; set; }
        public long ForgingOrdersQty { get; set; }

        // Critical Bottleneck Stations
        public long HeatTreatmentProcessID { get; set; }
        public int LotsAtTemperCount { get; set; }
        public long LotsAtTemperQty { get; set; }

        public long PolishingProcessID { get; set; }
        public int LotsAtPolishingCount { get; set; }
        public long LotsAtPolishingQty { get; set; }

        public long QCProcessID { get; set; }
        public int LotsAtQCCount { get; set; }
        public long LotsAtQCQty { get; set; }

        public long PackingProcessID { get; set; }
        public int LotsAtPackedCount { get; set; }
        public long LotsAtPackedQty { get; set; }
    }

    public class StockGroupSummaryItem
    {
        public int GroupID { get; set; }
        public string GroupName { get; set; } = string.Empty;
        public long TotalStock { get; set; }
        public decimal TotalValue { get; set; }
    }

    public class StockGroupItemDetail
    {
        public string ItemID { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public long TotalStock { get; set; }
        public decimal TotalValue { get; set; }
    }

    public class CommandCenterDashboardData
    {
        public DateTime AsOfDate { get; set; } = DateTime.Now;
        public HRPillarData HR { get; set; } = new();
        public StockPillarData Stock { get; set; } = new();
        public FinancialPillarData Financial { get; set; } = new();
        public ExportPillarData Export { get; set; } = new();
        public ProductionPillarData Production { get; set; } = new();
    }
}
