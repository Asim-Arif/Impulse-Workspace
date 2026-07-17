using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Export;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Export
{
    public partial class ItemwiseExportPerformance : ComponentBase
    {
        [Inject] private IExportPerformanceReportService ReportService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService Notification { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        private bool isLoading = true;

        // Lookup data
        private List<CustomerLookupModel> customers = new();
        private List<ExportItemGroupLookupModel> itemGroups = new();
        private List<ExportItemCategoryLookupModel> categories = new();
        private List<ExportMainGroupLookupModel> mainGroups = new();

        // Selected values (null = All)
        private CustomerLookupModel? selectedCustomer;
        private ExportItemGroupLookupModel? selectedItemGroup;
        private ExportItemCategoryLookupModel? selectedCategory;
        private ExportMainGroupLookupModel? selectedMainGroup;

        // Date range — defaults to last 6 months, same as legacy DtFrom = DateAdd("m", -6, Date)
        private DateTime dateFrom = DateTime.Today.AddMonths(-6);
        private DateTime dateTo = DateTime.Today;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                customers = await ReportService.GetCustomersAsync();
                itemGroups = await ReportService.GetItemGroupsAsync();
                categories = await ReportService.GetItemCategoriesAsync();
                mainGroups = await ReportService.GetMainGroupsAsync();
            }
            catch (Exception ex)
            {
                Notification.ShowError("Load Error", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        // --- Typeahead search methods ---

        private Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? customers
                : customers.FindAll(c => c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult<IEnumerable<CustomerLookupModel>>(results);
        }

        private Task<IEnumerable<ExportItemGroupLookupModel>> SearchItemGroups(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? itemGroups
                : itemGroups.FindAll(g => g.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult<IEnumerable<ExportItemGroupLookupModel>>(results);
        }

        private Task<IEnumerable<ExportItemCategoryLookupModel>> SearchCategories(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? categories
                : categories.FindAll(c => c.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult<IEnumerable<ExportItemCategoryLookupModel>>(results);
        }

        private Task<IEnumerable<ExportMainGroupLookupModel>> SearchMainGroups(string searchText)
        {
            var results = string.IsNullOrWhiteSpace(searchText)
                ? mainGroups
                : mainGroups.FindAll(mg => mg.MainGroupName.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult<IEnumerable<ExportMainGroupLookupModel>>(results);
        }

        // --- Value Changed handlers ---

        private void OnCustomerChanged(CustomerLookupModel? value) => selectedCustomer = value;
        private void OnItemGroupChanged(ExportItemGroupLookupModel? value) => selectedItemGroup = value;
        private void OnCategoryChanged(ExportItemCategoryLookupModel? value) => selectedCategory = value;
        private void OnMainGroupChanged(ExportMainGroupLookupModel? value) => selectedMainGroup = value;

        // --- Print ---

        private async Task PrintReport()
        {
            try
            {
                // Build parameters matching legacy rpt.ParameterFields:
                // @DTFrom, @DTTo, @ItemGroup (int, 0=all), @CatID (string, ""=all),
                // @AdditionalGroup (int, 0=all), @Customers (string, ""=all)
                var parameters = new Dictionary<string, object>
                {
                    { "@DTFrom",          dateFrom },
                    { "@DTTo",            dateTo },
                    { "@ItemGroup",       selectedItemGroup?.ID ?? 0 },
                    { "@CatID",           selectedCategory?.CatID.ToString() ?? "" },
                    { "@AdditionalGroup", selectedMainGroup?.MainGroupID ?? 0 },
                    { "@Customers",       selectedCustomer?.CustCode ?? "" }
                };

                var req = new ReportRequest
                {
                    ReportName = "ExportPerformanceReport.rpt",
                    Parameters = parameters
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                Notification.ShowError("Print Error", ex.Message);
            }
        }

        private void GoBack() => NavigationManager.NavigateTo("/export/statistics");
    }
}
