using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Impulse.Utilities;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export
{
    public partial class TotalDemand : ComponentBase
    {
        [Inject] private ITotalDemandService TotalDemandService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Dropdown & Filter bindings
        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;
        private string selectedRange = "90Days";
        private DateTime dateFrom;
        private DateTime dateTo;
        private bool showCustomAmount = false;

        // Data & Summary bindings
        private List<TotalDemandModel> demandData = new();
        private double totalAmount = 0;
        private string amountInWords = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var fetched = await TotalDemandService.GetActiveCustomersAsync();

                // Add default <All Customers> option
                customers.Add(new CustomerLookupModel { CustCode = "<All Customers>", Name = "All Customers" });
                customers.AddRange(fetched);

                selectedCustomer = customers[0];

                // Compute dates based on default range "90Days"
                SetDatesForRange();

                // Initial data load
                await RefreshDataInternal();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Load Error", "Failed to initialize page data: " + ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(searchText))
                return customers;
            return customers.Where(c => c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            await RefreshDataInternal();
        }

        private async Task OnRangeChanged(ChangeEventArgs e)
        {
            selectedRange = e.Value?.ToString() ?? "90Days";
            SetDatesForRange();
            await RefreshDataInternal();
        }

        private void SetDatesForRange()
        {
            DateTime today = DateTime.Today;
            dateTo = today;

            switch (selectedRange)
            {
                case "Today":
                    dateFrom = today;
                    break;
                case "15Days":
                    dateFrom = today.AddDays(-15);
                    break;
                case "30Days":
                    dateFrom = today.AddDays(-30);
                    break;
                case "60Days":
                    dateFrom = today.AddDays(-60);
                    break;
                case "90Days":
                    dateFrom = today.AddDays(-90);
                    break;
                case "Custom":
                    break;
            }
        }

        private async Task RefreshData()
        {
            isLoading = true;
            try
            {
                await RefreshDataInternal();
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshDataInternal()
        {
            string custCodeVal = (selectedCustomer != null && selectedCustomer.CustCode != "<All Customers>")
                ? selectedCustomer.CustCode
                : "0";

            try
            {
                demandData = await TotalDemandService.GetTotalDemandDataAsync(custCodeVal, dateFrom, dateTo);
                
                // Recalculate summary totals
                CalculateSummaryTotals();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Query Error", "Failed to query demand details: " + ex.Message);
            }
        }

        private void CalculateSummaryTotals()
        {
            totalAmount = demandData.Sum(row => row.OrderAmt);
            
            // Format number to words using utilities helper
            try
            {
                amountInWords = NumberToWordsConverter.Convert((decimal)totalAmount, "");
            }
            catch
            {
                amountInWords = "Conversion Error";
            }
        }

        private async Task PrintReport()
        {
            if (demandData.Count == 0)
                return;

            try
            {
                string selectionFormula = $"{{VTotalExport.DT}} in Date({dateFrom.Year}, {dateFrom.Month}, {dateFrom.Day}) to Date({dateTo.Year}, {dateTo.Month}, {dateTo.Day})";

                if (selectedCustomer != null && selectedCustomer.CustCode != "<All Customers>")
                {
                    selectionFormula += $" AND {{VTotalExport.CustCode}}='{selectedCustomer.CustCode}'";
                }

                var request = new ReportRequest
                {
                    ReportName = "TotalDemand.rpt",
                    SelectionFormula = selectionFormula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "CustomAmt", showCustomAmount ? "True" : "False" },
                        { "SubHeading", $"'{dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy}'" }
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Total demand report dispatched.");
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", "Failed to generate report: " + ex.Message);
            }
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo("/export/statistics");
        }
    }
}
