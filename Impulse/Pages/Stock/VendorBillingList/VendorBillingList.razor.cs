using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Stock.VendorBillingList
{
    public partial class VendorBillingList : ComponentBase
    {
        private List<VendorBillingListViewModel> BillingList;
        private bool IsLoading = false;
        
        public bool IsSampleBilling { get; set; } = false;

        private IEnumerable<object> Vendors = new List<object>();
        private object SelectedVendor;
        private VendorBillingListViewModel contextMenuRowItem;

        // Filters
        private class FilterModel
        {
            public int DateRangeIndex { get; set; } = 0;
            public DateTime DateFrom { get; set; } = DateTime.Today;
            public DateTime DateTo { get; set; } = DateTime.Today;
        }
        private FilterModel Filter = new();

        protected override async Task OnInitializedAsync()
        {
            var rawVendors = await BillingListService.GetVendorsAsync();
            Vendors = rawVendors.ToList();
        }

        private void OnDateRangeChanged()
        {
            Filter.DateTo = DateTime.Today;
            switch (Filter.DateRangeIndex)
            {
                case 0: Filter.DateFrom = DateTime.Today; break;
                case 1: Filter.DateFrom = DateTime.Today.AddDays(-15); break;
                case 2: Filter.DateFrom = DateTime.Today.AddDays(-30); break;
                case 3: Filter.DateFrom = DateTime.Today.AddDays(-60); break;
                case 4: Filter.DateFrom = DateTime.Today.AddDays(-90); break;
                case 5: break; // Custom
            }
        }

        private Task<IEnumerable<object>> SearchVendors(string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
            {
                return Task.FromResult(Vendors);
            }
            var filtered = Vendors.Where(x => ((dynamic)x).AccTitle.ToString().Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
            return Task.FromResult((IEnumerable<object>)filtered);
        }

        private async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                string vendorId = SelectedVendor != null ? ((dynamic)SelectedVendor).AccNo.ToString() : "0";
                
                var data = await BillingListService.GetVendorBillingListAsync(
                    Filter.DateFrom, 
                    Filter.DateTo, 
                    vendorId, 
                    IsSampleBilling);
                    
                BillingList = data.ToList();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(Radzen.NotificationSeverity.Error, "Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task OnPrintList()
        {
            string vendorId = SelectedVendor != null ? ((dynamic)SelectedVendor).AccNo.ToString() : "0";
            string selection = $"{{VVenderPostedBillsList.DT}}>=Date({Filter.DateFrom.Year},{Filter.DateFrom.Month},{Filter.DateFrom.Day}) AND {{VVenderPostedBillsList.DT}}<=Date({Filter.DateTo.Year},{Filter.DateTo.Month},{Filter.DateTo.Day})";
            
            if (vendorId != "0")
            {
                selection += $" AND {{VVenderPostedBillsList.AccNo}}='{vendorId}'";
            }
            
            selection += $" AND {{VVenderPostedBillsList.SampleBilling}}={(IsSampleBilling ? "True" : "False")}";

            var req = new ReportRequest
            {
                ReportName = "VendorPostedBillsList.rpt",
                SelectionFormula = selection
            };
            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task OpenRowContextMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e, VendorBillingListViewModel item)
        {
            contextMenuRowItem = item;
            await BlazorContextMenuService.ShowMenu("rowContextMenu", (int)e.ClientX, (int)e.ClientY);
        }

        private async Task OnPrintBill(ItemClickEventArgs e)
        {
            if (contextMenuRowItem != null)
            {
                var req = new ReportRequest
                {
                    ReportName = "VenderPostedBill.rpt",
                    SelectionFormula = $"{{VenderPostedBills.VchrNo}}='{contextMenuRowItem.VchrNo}'"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
        }

        private async Task OnDeleteBill(ItemClickEventArgs e)
        {
            if (contextMenuRowItem != null)
            {
                // Verify Password via Global Security Service
                bool isAuthorized = await SecurityService.VerifyActionAsync("Delete_Vendor_Bill");
                
                if (!isAuthorized)
                {
                    NotificationService.Notify(Radzen.NotificationSeverity.Error, "Unauthorized", "Incorrect password or action cancelled.");
                    return;
                }

                try
                {
                    await BillingListService.DeleteVendorBillAsync(contextMenuRowItem.EntryID);
                    NotificationService.Notify(Radzen.NotificationSeverity.Success, "Success", "Bill deleted successfully.");
                    await LoadDataAsync();
                }
                catch (Exception ex)
                {
                    NotificationService.Notify(Radzen.NotificationSeverity.Error, "Error", $"Failed to delete bill: {ex.Message}");
                }
            }
        }
    }
}
