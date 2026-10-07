using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.IntraOffice
{
    public class DashboardMetricsDto
    {
        public int TotalLeads { get; set; } = 42;
        public int NewInquiries { get; set; } = 18;
        public int ActiveCustomers { get; set; } = 126;
        public int PendingQuotations { get; set; } = 14;
        public int QuotationsWon { get; set; } = 38;
        public int QuotationsLost { get; set; } = 6;
        public int ConfirmedSalesOrders { get; set; } = 41;
        public int OrdersInProduction { get; set; } = 18;
        public int OrdersReadyForShipment { get; set; } = 5;
        public int ShipmentsInTransit { get; set; } = 12;
        public decimal OutstandingReceivables { get; set; } = 384000;
        public decimal OverduePayments { get; set; } = 42500;
        public decimal MonthlySalesTotal { get; set; } = 3890200;
        public decimal MonthlyExportSalesTotal { get; set; } = 2450000;
        public List<CountrySalesDto> SalesByCountry { get; set; } = new();
        public List<ProductSalesDto> SalesByProduct { get; set; } = new();
        public List<SalespersonStatDto> SalespersonStats { get; set; } = new();
        public List<CrmRecentOrderDto> RecentOrders { get; set; } = new();
        public List<CrmRecentShipmentDto> RecentShipments { get; set; } = new();
    }

    public class CountrySalesDto
    {
        public string Country { get; set; } = string.Empty;
        public decimal TotalAmount { get; set; }
        public int OrderCount { get; set; }
        public double Percentage { get; set; }
    }

    public class ProductSalesDto
    {
        public string SKU { get; set; } = "";
        public string ProductName { get; set; } = "";
        public int TotalQuantity { get; set; }
        public decimal TotalRevenue { get; set; }
    }

    public class SalespersonStatDto
    {
        public string SalespersonName { get; set; } = "";
        public int LeadsCount { get; set; }
        public int QuotationsCount { get; set; }
        public int OrdersCount { get; set; }
        public decimal ClosedRevenue { get; set; }
    }

    public class ARAgingSummaryDto
    {
        public decimal CurrentNotDue { get; set; }
        public decimal Days1To30 { get; set; }
        public decimal Days31To60 { get; set; }
        public decimal Days61To90 { get; set; }
        public decimal Days90Plus { get; set; }
        public decimal TotalOutstanding => CurrentNotDue + Days1To30 + Days31To60 + Days61To90 + Days90Plus;
    }

    public class CrmRecentOrderDto
    {
        public string OrderNumber { get; set; } = "";
        public string CustomerName { get; set; } = "";
        public string Country { get; set; } = "";
        public decimal TotalAmount { get; set; }
        public string Status { get; set; } = "InProduction";
        public DateTime OrderDate { get; set; } = DateTime.UtcNow;
    }

    public class CrmRecentShipmentDto
    {
        public string TrackingNumber { get; set; } = "";
        public string DestinationCountry { get; set; } = "";
        public string Carrier { get; set; } = "Maersk Line";
        public string Status { get; set; } = "InTransit";
        public DateTime EstimatedDelivery { get; set; } = DateTime.UtcNow.AddDays(7);
    }
}

