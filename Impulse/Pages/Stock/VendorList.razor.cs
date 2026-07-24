using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using Impulse.Services;

namespace Impulse.Pages.Stock
{
    public partial class VendorList : ComponentBase
    {
        [Inject] public INewVendorService VendorService { get; set; } = default!;
        [Inject] public NavigationManager NavigationManager { get; set; } = default!;
        [Inject] public IJSRuntime JSRuntime { get; set; } = default!;
        [Inject] public INotificationService NotificationService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigation { get; set; } = default!;

        private bool isLoading = true;
        private List<NewVendorViewModel> allVendors = new();
        private List<NewVendorViewModel> filteredVendors = new();
        private NewVendorViewModel? highlightedVendor;

        private string searchText = "";
        private bool showInactive = false;

        protected override async Task OnInitializedAsync()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            isLoading = true;
            try
            {
                allVendors = await VendorService.GetVendorsAsync();
                ApplyFilters();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error loading vendors", ex.Message);
            }
            finally
            {
                isLoading = false;
            }
        }

        private void ApplyFilters()
        {
            var query = allVendors.AsEnumerable();

            // Filter out inactive unless shown
            if (!showInactive)
            {
                query = query.Where(v => !v.Inactive);
            }

            // Text search
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                var lowerSearch = searchText.ToLower();
                query = query.Where(v =>
                    (v.VendorName != null && v.VendorName.ToLower().Contains(lowerSearch)) ||
                    (v.MakerNo != null && v.MakerNo.ToLower().Contains(lowerSearch)) ||
                    (v.ContactPerson != null && v.ContactPerson.ToLower().Contains(lowerSearch))
                );
            }

            filteredVendors = query.ToList();
            StateHasChanged();
        }

        private void OnSearchInput(ChangeEventArgs e)
        {
            searchText = e.Value?.ToString() ?? "";
            ApplyFilters();
        }

        private void OnShowInactiveChanged(ChangeEventArgs e)
        {
            showInactive = (bool)(e.Value ?? false);
            ApplyFilters();
        }

        private void SelectVendor(NewVendorViewModel vendor)
        {
            highlightedVendor = vendor;
        }

        private void EditVendor(NewVendorViewModel vendor)
        {
            NavigationManager.NavigateTo($"/newvendor/{vendor.VendID}?returnUrl=/vendorlist");
        }

        private void AddNewVendor()
        {
            NavigationManager.NavigateTo("/newvendor?returnUrl=/vendorlist");
        }

        // Context Menu Handlers
        private void OnEditVendorClick(ItemClickEventArgs e)
        {
            if (e.Data is NewVendorViewModel vendor)
            {
                EditVendor(vendor);
            }
        }

        private async Task OnToggleActiveClick(ItemClickEventArgs e)
        {
            if (e.Data is NewVendorViewModel vendor)
            {
                if (vendor.CurrentBalance != 0 && !vendor.Inactive)
                {
                    NotificationService.ShowWarning("Cannot Deactivate", "This Vendor Cannot Be Inactive Unless Balance Becomes Zero.");
                    return;
                }

                bool newActiveStatus = vendor.Inactive; // If it's currently inactive, make it active
                string actionText = vendor.Inactive ? "active" : "inactive";
                
                bool confirmed = await NotificationService.ShowQuestionNotification("Confirm Status Change", $"Are you sure you want to make '{vendor.VendorName}' {actionText}?");
                if (confirmed)
                {
                    try
                    {
                        await VendorService.ToggleVendorStatusAsync(vendor.AccNo, newActiveStatus);
                        NotificationService.ShowSuccess("Success", $"Vendor '{vendor.VendorName}' is now {actionText}.");
                        await LoadDataAsync(); // Refresh data
                    }
                    catch (Exception ex)
                    {
                        NotificationService.ShowError("Status Update Failed", ex.Message);
                    }
                }
            }
        }

        private NewVendorViewModel? GetItemContext(ItemClickEventArgs? e)
        {
            return e?.Data as NewVendorViewModel ?? highlightedVendor;
        }

        private async Task PrintVendorList(ItemClickEventArgs? e = null)
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = "VenderList.rpt",
                    SelectionFormula = ""
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintVendorReport(string reportName, ItemClickEventArgs? e)
        {
            var item = GetItemContext(e);
            if (item == null)
            {
                NotificationService.ShowWarning("No vendor selected", "Please select a vendor to print.");
                return;
            }
            try
            {
                var request = new ReportRequest
                {
                    ReportName = reportName,
                    SelectionFormula = $"{{VMakers.VendID}}={item.VendID}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Print Error", ex.Message);
            }
        }

        private async Task PrintVendorFileLabel(ItemClickEventArgs? e = null) => await PrintVendorReport("Print_Vender_File_Label.rpt", e);
        private async Task PrintVendorProfile(ItemClickEventArgs? e = null) => await PrintVendorReport("Print_Vender_Profile.rpt", e);
        private async Task PrintVendorDetails(ItemClickEventArgs? e = null) => await PrintVendorReport("Print_Vender_Details.rpt", e);
        private async Task PrintVendorLabel(ItemClickEventArgs? e = null) => await PrintVendorReport("Print_Vender_Lable.rpt", e);
        private async Task PrintVendorShippingLabel(ItemClickEventArgs? e = null) => await PrintVendorReport("Print_Vender_Shipping_Lable.rpt", e);
    }
}
