using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Radzen;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels;
using BlazorContextMenu;
using Impulse.Utilities;
using Impulse.Services;

namespace Impulse.Pages.Export.CustomInvoices
{
    public partial class CustomPaymentStatus : ComponentBase
    {
        [Inject] protected Radzen.NotificationService NotificationService { get; set; } = null!;
        [Inject] protected IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;

        public bool IsLoading { get; set; } = false;
        public List<CustomPaymentStatusModel> PaymentStatuses { get; set; } = new();
        public List<AllCustomersBalanceModel> AllCustomersBalances { get; set; } = new();
        
        public CustomerLookupModel? SelectedCustomer { get; set; }
        public int StatusIndex { get; set; } = 0; // 0=UnPaid, 1=Paid, 2=All
        public bool ShowAllCustomers { get; set; } = false;

        public decimal TotalAmount { get; set; } = 0;
        public string AmountInWords { get; set; } = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        public async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            var allCustomers = await CustomInvoiceService.GetCustomersAsync();
            return allCustomers.Where(c => c.CustCode != "HICO" && 
                (string.IsNullOrEmpty(searchText) || 
                 c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                 c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task OnCustomerChanged(CustomerLookupModel cust)
        {
            SelectedCustomer = cust;
            await LoadDataAsync();
        }

        public async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                string? custCode = SelectedCustomer?.CustCode; // if null, it's <All Customers>
                PaymentStatuses = await CustomPaymentService.GetCustomPaymentStatusesAsync(custCode, StatusIndex);
                CalculateTotals();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading payment statuses: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void CalculateTotals()
        {
            TotalAmount = PaymentStatuses.Sum(x => x.CustomAmtToRcv);
            
            // Assuming USD by default or extracting from the first available payment status model
            // For now just format to default or no currency string
            AmountInWords = TotalAmount > 0 ? NumberToWordsConverter.Convert((decimal)TotalAmount, "") : ""; 
        }

        public async Task ToggleAllCustomersView()
        {
            ShowAllCustomers = !ShowAllCustomers;
            if (ShowAllCustomers)
            {
                IsLoading = true;
                AllCustomersBalances = await CustomPaymentService.GetAllCustomersBalanceAsync();
                IsLoading = false;
            }
        }

        public async Task ShowContextMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e, CustomPaymentStatusModel data, string menuId)
        {
            SelectedPaymentData = data;
            await BlazorContextMenuService.ShowMenu(menuId, (int)e.ClientX, (int)e.ClientY);
        }

        private CustomPaymentStatusModel? SelectedPaymentData { get; set; }

        public void OnReceivePaymentClick(ItemClickEventArgs e)
        {
            var data = e.Data as CustomPaymentStatusModel ?? SelectedPaymentData;
            if (data != null && !string.IsNullOrEmpty(data.CustomInvoice))
            {
                NavManager.NavigateTo($"/export/receive-custom-payment/{data.CustomInvoice}/{data.CustCode}");
            }
        }

        public async Task OnRefreshClick(ItemClickEventArgs e)
        {
            await LoadDataAsync();
        }

        public async Task OnPrintBankLetterClick(ItemClickEventArgs e)
        {
            var data = e.Data as CustomPaymentStatusModel ?? SelectedPaymentData;
            if (data == null) return;

            if (data.EntryType != 1)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Invalid Selection", "Bank letters can only be printed for payment receipts, not invoices.");
                return;
            }

            var request = new ReportRequest
            {
                ReportName = "rptLetterToBank.rpt",
                SelectionFormula = $"{{VPaymentsRcvd.EntryID}}={data.EntryID}"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        public async Task OnPrintPaymentStatusForThisInvoiceClick(ItemClickEventArgs e)
        {
            var data = e.Data as CustomPaymentStatusModel ?? SelectedPaymentData;
            if (data == null) return;

            var request = new ReportRequest
            {
                ReportName = "rptPaymentStatus.rpt",
                SelectionFormula = $"{{CustomInvoice.CustomInvoice}}='{data.CustomInvoice}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        public async Task OnPrintPaymentStatusOfAllClick(ItemClickEventArgs e)
        {
            if (PaymentStatuses == null || !PaymentStatuses.Any()) return;

            var invoiceNumbers = PaymentStatuses
                .Select(p => $"'{p.CustomInvoice}'")
                .Distinct();

            var request = new ReportRequest
            {
                ReportName = "rptPaymentStatus.rpt",
                SelectionFormula = $"{{CustomInvoice.CustomInvoice}} IN [{string.Join(", ", invoiceNumbers)}]"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        public async Task OnPrintStatusCustomerwiseClick(ItemClickEventArgs e)
        {
            if (PaymentStatuses == null || !PaymentStatuses.Any()) return;

            var invoiceNumbers = PaymentStatuses
                .Select(p => $"'{p.CustomInvoice}'")
                .Distinct();

            var request = new ReportRequest
            {
                ReportName = "PaymentStatusCustomerWise.rpt",
                SelectionFormula = $"{{CustomInvoice.CustomInvoice}} IN [{string.Join(", ", invoiceNumbers)}]"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        public void OnPostPrcClick(ItemClickEventArgs e)
        {
            if (SelectedPaymentData != null && SelectedPaymentData.EntryID > 0)
            {
                NavManager.NavigateTo($"/export/post-prc/{SelectedPaymentData.EntryID}");
            }
            else
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", "Payment record not selected.");
            }
        }
    }
}
