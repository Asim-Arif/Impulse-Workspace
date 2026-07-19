using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Accounts
{
    public class CustomInvoiceAuthItemModel
    {
        // Columns returned by dbo.VCustomInvoiceAuth
        public string Curr { get; set; } = string.Empty;
        public string CustomInvoice { get; set; } = string.Empty;
        public DateTime? DT { get; set; }
        public string CustCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public double TotalAmt { get; set; }
        public int PackingInfo { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public double TotalCustomAmt { get; set; }

        // Helper derived fields for UI display (legacy ListView checkbox / inline exchange-rate editor state)
        public bool IsChecked { get; set; }
        public bool IsEditingRate { get; set; }
        public double ExchRate { get; set; }
        public double EditingExchRate { get; set; }
        public double AmtInPakRs { get; set; }
    }

    public class CustomInvoiceAdvanceAdjustmentModel
    {
        public string CustomInvoice { get; set; } = string.Empty;
        public double Amount { get; set; }
        public double ExchRate { get; set; }
    }

    public class CustomInvoiceAuthSaveModel
    {
        public DateTime VoucherDate { get; set; }
        public string SalesAccNo { get; set; } = string.Empty;
        public string ExchDiffAccNo { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public List<CustomInvoiceAuthItemModel> Items { get; set; } = new();
    }
}
