using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Export;
using DataAccessLibrary.Models.ViewModels.Company;
using Impulse.Services;
using Impulse.Services.Export;
using Microsoft.JSInterop;

namespace Impulse.Pages.Export.Proformas
{
    public partial class ProformaList : ComponentBase
    {
        [Inject] private IProformaListService ProformaListService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private INotificationService NotificationServiceManager { get; set; } = null!;
        [Inject] private IDBHelperService DbHelper { get; set; } = null!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = null!;
        [Inject] private IJSRuntime JS { get; set; } = null!;
        [Inject] private NavigationManager NavManager { get; set; } = null!;

        private bool isLoading = true;
        private string searchText = string.Empty;

        // Filters
        private string selectedDateRange = "180";
        private DateTime fromDate = DateTime.Today.AddDays(-180);
        private DateTime toDate = DateTime.Today;
        private long selectedCompanyId = 0;
        private string selectedCountry = string.Empty;
        private CustomerLookupModel? selectedCustomer = null;

        private List<CompanyLookupModel> companies = new();
        private List<CustomerLookupModel> customers = new();
        private List<string> countries = new();

        private List<ProformaListModel> allItems = new();
        
        // Modals
        private bool showPasswordModal = false;
        private string passwordSettingName = string.Empty;
        private string enteredPassword = string.Empty;
        private string passwordError = string.Empty;
        private Func<Task>? postPasswordAction = null;

        private bool showCompanyModal = false;
        private long overrideCompanyId = 0;
        
        private bool showDiscountModal = false;
        private double overrideDiscount = 0;

        private ProformaListModel? contextMenuRowItem = null;

        protected override async Task OnInitializedAsync()
        {
            await LoadLookups();
            await LoadData();
        }

        private async Task LoadLookups()
        {
            try
            {
                companies = (await ProformaListService.GetCompaniesAsync()).ToList();
                customers = (await ProformaListService.GetCustomersAsync()).ToList();
                countries = (await ProformaListService.GetCountriesAsync(selectedCustomer?.CustCode ?? "0")).ToList();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        private void OnDateRangeChanged()
        {
            toDate = DateTime.Today;
            if (int.TryParse(selectedDateRange, out int days))
            {
                fromDate = DateTime.Today.AddDays(-days);
                _ = LoadData();
            }
        }

        private async Task OnCustomerChanged(CustomerLookupModel? customer)
        {
            selectedCustomer = customer;
            selectedCountry = string.Empty;
            countries = (await ProformaListService.GetCountriesAsync(selectedCustomer?.CustCode ?? "0")).ToList();
            await LoadData();
        }

        private async Task OnCompanyChanged(ChangeEventArgs e)
        {
            if (long.TryParse(e.Value?.ToString(), out long compId))
            {
                selectedCompanyId = compId;
                await LoadData();
            }
        }

        private async Task OnCountryChanged(ChangeEventArgs e)
        {
            selectedCountry = e.Value?.ToString() ?? string.Empty;
            await LoadData();
        }

        private async Task<IEnumerable<CustomerLookupModel>> SearchCustomers(string searchText)
        {
            return FilterCustomers(searchText);
        }

        private IEnumerable<CustomerLookupModel> FilterCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return customers;
            return customers.Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                                        c.CustCode.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        private async Task LoadData()
        {
            try
            {
                isLoading = true;
                StateHasChanged();

                string custCode = selectedCustomer?.CustCode ?? "0";
                var data = await ProformaListService.GetProformaListAsync(fromDate, toDate, custCode, selectedCompanyId, selectedCountry);
                allItems = data.ToList();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error Loading Data", ex.Message);
            }
            finally
            {
                isLoading = false;
                StateHasChanged();
            }
        }

        private IEnumerable<ProformaListModel> FilteredItems => string.IsNullOrWhiteSpace(searchText) 
            ? allItems 
            : allItems.Where(i => 
                (i.PInvoice?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (i.CompanyName?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (i.CustCode?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (i.Orders?.Contains(searchText, StringComparison.OrdinalIgnoreCase) ?? false));

        private string GetRowCssClass(ProformaListModel item)
        {
            if (item.AdvanceAmount > 0) return "text-primary";
            if (item.ChargesAmt > 0) return "text-danger";
            return "";
        }

        // --- Context Menus ---
        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task ShowRowContextMenu(MouseEventArgs e, ProformaListModel item)
        {
            contextMenuRowItem = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task CheckPasswordThenExecute(string settingName, Func<Task> action)
        {
            passwordSettingName = settingName;
            enteredPassword = string.Empty;
            passwordError = string.Empty;
            postPasswordAction = action;
            showPasswordModal = true;
            await Task.CompletedTask;
        }

        private async Task VerifyPasswordAndExecute()
        {
            try
            {
                string correctPassword = await DbHelper.getDatabasePasswordAsync(passwordSettingName);
                if (enteredPassword == correctPassword)
                {
                    showPasswordModal = false;
                    if (postPasswordAction != null)
                    {
                        await postPasswordAction.Invoke();
                    }
                }
                else
                {
                    passwordError = "Incorrect password. Please try again.";
                }
            }
            catch (Exception ex)
            {
                passwordError = $"Error validating password: {ex.Message}";
            }
        }

        // --- Row Actions ---
        private void EditProforma(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            NavManager.NavigateTo($"/export/proformas/new-proforma/{contextMenuRowItem.PInvoice}");
        }

        private async Task DeleteProforma(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;

            bool confirm = await JS.InvokeAsync<bool>("confirm", $"Are you sure you want to delete proforma {contextMenuRowItem.PInvoice}?");
            if (!confirm) return;

            _ = CheckPasswordThenExecute("DeleteExportEntries", async () =>
            {
                try
                {
                    await ProformaListService.DeleteProformaAsync(contextMenuRowItem.PInvoice);
                    NotificationServiceManager.ShowSuccess("Success", "Proforma deleted successfully.");
                    await LoadData();
                }
                catch (Exception ex)
                {
                    NotificationServiceManager.ShowError("Delete Error", ex.Message);
                }
            });
        }

        private async Task PrintProforma()
        {
            if (contextMenuRowItem == null) return;
            var request = new ReportRequest
            {
                ReportName = "rptFProforma.rpt",
                SelectionFormula = $"{{VrptProforma.PInvoice}}='{contextMenuRowItem.PInvoice}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintUSAProforma(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            var request = new ReportRequest
            {
                ReportName = "Proforma_USA.rpt",
                SelectionFormula = $"{{VrptProforma.PInvoice}}='{contextMenuRowItem.PInvoice}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        private void ShowCompanyOverrideModal(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            overrideCompanyId = 0;
            showCompanyModal = true;
        }

        private async Task PrintProformaCompanySelection()
        {
            showCompanyModal = false;
            if (contextMenuRowItem == null || overrideCompanyId == 0) return;
            
            // Subreport logic implies overriding {Companies.EntryID} or passing it somehow.
            // Normally it prints rptFProforma.rpt with CompanyRefID. 
            // In Blazor, we can pass it via report parameters or subreport formula.
            var request = new ReportRequest
            {
                ReportName = "rptFProforma.rpt",
                SelectionFormula = $"{{VrptProforma.PInvoice}}='{contextMenuRowItem.PInvoice}'"
            };
            request.Parameters.Add("CompanyRefID", overrideCompanyId.ToString());
            
            await ReportNavigation.PrintReportAsync(request);
        }

        private void ShowDiscountModal(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            overrideDiscount = 0;
            showDiscountModal = true;
        }

        private async Task PrintProformaDiscount()
        {
            showDiscountModal = false;
            if (contextMenuRowItem == null) return;

            var request = new ReportRequest
            {
                ReportName = "rptFProforma.rpt",
                SelectionFormula = $"{{VrptProforma.PInvoice}}='{contextMenuRowItem.PInvoice}'"
            };
            request.Parameters.Add("Discount", overrideDiscount.ToString());

            await ReportNavigation.PrintReportAsync(request);
        }

        private async Task PrintProformaWD(ItemClickEventArgs e)
        {
            if (contextMenuRowItem == null) return;
            var request = new ReportRequest
            {
                ReportName = "rptFProformaWD.rpt",
                SelectionFormula = $"{{VrptProforma.PInvoice}}='{contextMenuRowItem.PInvoice}'"
            };
            await ReportNavigation.PrintReportAsync(request);
        }

        // --- General Reports ---
        private async Task PrintList(ItemClickEventArgs e)
        {
            string sel = $"{{VFProformaList.DT}}=#{fromDate:yyyy-MM-dd}# TO #{toDate:yyyy-MM-dd}#";
            
            if (selectedCustomer != null && selectedCustomer.CustCode != "0")
            {
                sel += $" AND {{VFProformaList.CustCode}}='{selectedCustomer.CustCode}'";
            }
            if (!string.IsNullOrEmpty(selectedCountry))
            {
                sel += $" AND {{VFProformaList.Country}}='{selectedCountry}'";
            }

            var request = new ReportRequest
            {
                ReportName = "ProformaList.rpt",
                SelectionFormula = sel
            };
            
            string customersStr = $"{(selectedCustomer?.Name ?? "<All Customers>")} {(string.IsNullOrEmpty(selectedCountry) ? "" : selectedCountry)}";
            string dateRangeStr = $"{fromDate:dd-MMM-yyyy} to {toDate:dd-MMM-yyyy}";
            
            request.Parameters.Add("Customers", customersStr);
            request.Parameters.Add("DateRange", dateRangeStr);

            await ReportNavigation.PrintReportAsync(request);
        }
    }
}
