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
