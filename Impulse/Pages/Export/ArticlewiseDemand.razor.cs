using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Company;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Impulse.Utilities;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export
{
    public partial class ArticlewiseDemand : ComponentBase
    {
        [Inject] private IArticlewiseDemandService ArticlewiseDemandService { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Selection bindings
        private List<ItemLookupModel> articles = new();
        private ItemLookupModel? selectedArticle;

        private List<CustomerLookupModel> customers = new();
        private CustomerLookupModel? selectedCustomer;

        private List<string> countries = new();
        private string selectedCountry = "0";

        private bool useDateFilter = false;
        private DateTime dateFrom = DateTime.Today;
        private DateTime dateTo = DateTime.Today;

        // Data & Summary bindings
        private List<ArticlewiseDemandModel> demandData = new();
        private double totalQuantity = 0;
        private string quantityInWords = string.Empty;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                var fetchedArticles = await ArticlewiseDemandService.GetArticlesAsync();
                articles.Add(new ItemLookupModel { ItemID = "0", ItemName = "<All Articles>" });
                articles.AddRange(fetchedArticles);
                selectedArticle = articles[0];

                var fetchedCustomers = await ArticlewiseDemandService.GetActiveCustomersAsync();
                customers.Add(new CustomerLookupModel { CustCode = "0", Name = "All Customers" });
                customers.AddRange(fetchedCustomers);
                selectedCustomer = customers[0];

                await LoadCountriesAsync();
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

        private async Task<IEnumerable<ItemLookupModel>> SearchArticles(string searchText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(searchText))
                return articles;
            return articles.Where(a => a.ItemName.Contains(searchText, StringComparison.OrdinalIgnoreCase) || a.ItemID.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            await Task.Yield();
            if (string.IsNullOrWhiteSpace(searchText))
                return customers;
            return customers.Where(c => c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task OnArticleChanged(ItemLookupModel? art)
        {
            selectedArticle = art;
            await RefreshDataInternal();
        }

        private async Task OnCustomerChanged(CustomerLookupModel? cust)
        {
            selectedCustomer = cust;
            await LoadCountriesAsync();
            await RefreshDataInternal();
        }

        private async Task LoadCountriesAsync()
        {
            string custCodeVal = (selectedCustomer != null) ? selectedCustomer.CustCode : "0";
            
            try
            {
                var fetchedCountries = await ArticlewiseDemandService.GetCountriesAsync(custCodeVal);
                countries.Clear();
                countries.Add("<All Countries>");
                countries.AddRange(fetchedCountries);
                selectedCountry = "<All Countries>";
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Data Error", "Failed to load countries: " + ex.Message);
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? "<All Countries>";
            await RefreshDataInternal();
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
            string artId = (selectedArticle != null) ? selectedArticle.ItemID : "0";
            string custCode = (selectedCustomer != null) ? selectedCustomer.CustCode : "0";
            string country = (selectedCountry != "<All Countries>") ? selectedCountry : "0";

            try
            {
                demandData = await ArticlewiseDemandService.GetArticlewiseDemandDataAsync(artId, custCode, country, useDateFilter, dateFrom, dateTo);
                CalculateSummaryTotals();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Query Error", "Failed to query demand data: " + ex.Message);
            }
        }

        private void CalculateSummaryTotals()
        {
            totalQuantity = demandData.Sum(row => row.Qty);
            try
            {
                quantityInWords = NumberToWordsConverter.Convert((decimal)totalQuantity, "");
            }
            catch
            {
                quantityInWords = "Conversion Error";
            }
        }

        private async Task PrintReport(bool withValue)
        {
            if (demandData.Count == 0)
                return;

            try
            {
                string reportName = withValue ? "rptArticlewiseOrders_Value.rpt" : "rptArticlewiseOrders.rpt";

                var constraints = new List<string>();

                if (selectedArticle != null && selectedArticle.ItemID != "0")
                {
                    constraints.Add($"{{VArticlewiseOrders.ItemID}}='{selectedArticle.ItemID}'");
                }

                if (selectedCustomer != null && selectedCustomer.CustCode != "0")
                {
                    constraints.Add($"{{VArticlewiseOrders.CustCode}}='{selectedCustomer.CustCode}'");
                }

                if (selectedCountry != "<All Countries>")
                {
                    constraints.Add($"{{VArticlewiseOrders.Country}}='{selectedCountry}'");
                }

                if (useDateFilter)
                {
                    constraints.Add($"{{VArticlewiseOrders.DT}} in Date({dateFrom.Year}, {dateFrom.Month}, {dateFrom.Day}) to Date({dateTo.Year}, {dateTo.Month}, {dateTo.Day})");
                }

                string selectionFormula = constraints.Count > 0 ? string.Join(" AND ", constraints) : "1=1";
                string strTitle = useDateFilter ? $"{dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy}" : "No Date Range";

                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = selectionFormula,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "SubHeading", $"'{strTitle}'" },
                        { "DateRange", $"'{strTitle}'" },
                        { "Title", $"'{strTitle}'" }
                    }
                };

                await ReportNavigation.PrintReportAsync(request);
                NotificationServiceManager.ShowSuccess("Success", "Articlewise demand report dispatched.");
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
