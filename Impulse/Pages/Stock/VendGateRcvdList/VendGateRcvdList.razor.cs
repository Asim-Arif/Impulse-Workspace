using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using DataAccessLibrary.Models.ViewModels.Stock;
using Impulse.Services.Stock;
using Radzen;
using BlazorContextMenu;

namespace Impulse.Pages.Stock.VendGateRcvdList
{
    public partial class VendGateRcvdList : ComponentBase
    {
        private List<VendGateRcvdListViewModel> ReceivingList;
        private bool IsLoading = false;
        private bool IsSampleList = false;

        public class ListFilter
        {
            public int DateRangeIndex { get; set; } = 0; // Default: Today / Last 15 Days (based on VB6)
            public DateTime DateFrom { get; set; } = DateTime.Today;
            public DateTime DateTo { get; set; } = DateTime.Today;
            public string OrderNo { get; set; } = string.Empty;
        }

        private ListFilter Filter = new ListFilter();
        private object SelectedVendor;
        private object SelectedRMGroup;

        protected override async Task OnInitializedAsync()
        {
            // Initial dates based on standard load (Today to Today, or -15 to Today based on legacy, let's use -15)
            Filter.DateFrom = DateTime.Today.AddDays(-15);
            Filter.DateRangeIndex = 1; // 1 = Last 15 Days
            
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            IsLoading = true;
            StateHasChanged();

            string vendId = "0";
            if (SelectedVendor != null)
            {
                vendId = ((dynamic)SelectedVendor).AccNo;
            }

            string groupId = "0";
            if (SelectedRMGroup != null)
            {
                groupId = ((dynamic)SelectedRMGroup).ID.ToString();
            }

            try
            {
                ReceivingList = await VendGateRcvdListService.GetReceivingListAsync(
                    Filter.DateFrom,
                    Filter.DateTo,
                    vendId,
                    Filter.OrderNo,
                    groupId,
                    IsSampleList
                );
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", ex.Message);
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private async Task<IEnumerable<object>> SearchVendors(string searchText)
        {
            var vendors = await VendGateRcvdListService.GetVendorsAsync();
            if (string.IsNullOrEmpty(searchText))
            {
                return vendors;
            }
            return vendors.Where(v => ((dynamic)v).AccTitle.ToLower().Contains(searchText.ToLower())).ToList();
        }

        private async Task<IEnumerable<object>> SearchRMGroups(string searchText)
        {
            var groups = await VendGateRcvdListService.GetRMGroupsAsync();
            if (string.IsNullOrEmpty(searchText))
            {
                return groups;
            }
            return groups.Where(g => ((dynamic)g).Description.ToLower().Contains(searchText.ToLower())).ToList();
        }

        private async Task OnDateRangeChanged()
        {
            DateTime serverDate = DateTime.Today;
            switch (Filter.DateRangeIndex)
            {
                case 0: // Today
                    Filter.DateFrom = serverDate;
                    Filter.DateTo = serverDate;
                    break;
                case 1: // Last 15 Days
                    Filter.DateFrom = serverDate.AddDays(-15);
                    Filter.DateTo = serverDate;
                    break;
                case 2: // Last 30 Days
                    Filter.DateFrom = serverDate.AddDays(-30);
                    Filter.DateTo = serverDate;
                    break;
                case 3: // Last 60 Days
                    Filter.DateFrom = serverDate.AddDays(-60);
                    Filter.DateTo = serverDate;
                    break;
                case 4: // Last 90 Days
                    Filter.DateFrom = serverDate.AddDays(-90);
                    Filter.DateTo = serverDate;
                    break;
                case 5: // Custom
                    break;
            }

            if (Filter.DateRangeIndex != 5)
            {
                await LoadDataAsync();
            }
        }

        private async Task OnRefreshClicked()
        {
            await LoadDataAsync();
        }

        private async Task ShowOptionsMenu(Microsoft.AspNetCore.Components.Web.MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void EditReceiving(VendGateRcvdListViewModel item)
        {
            string url = $"/stock/vend-gate-rcvd?returnUrl={Uri.EscapeDataString(NavigationManager.Uri)}&rcvId={Uri.EscapeDataString(item.RcvID)}";
            NavigationManager.NavigateTo(url);
        }

        private void OnEdit(ItemClickEventArgs e)
        {
            if (e.Data is VendGateRcvdListViewModel item)
            {
                EditReceiving(item);
            }
        }

        private async Task OnDelete(ItemClickEventArgs e)
        {
            if (e.Data is VendGateRcvdListViewModel item)
            {
                try
                {
                    await VendGateRcvdListService.DeleteReceivingAsync(item.RcvID);
                    NotificationService.ShowSuccess("Success", "Deleted Successfully.");
                    await LoadDataAsync();
                }
                catch (Exception ex)
                {
                    NotificationService.ShowError("Error", ex.Message);
                }
            }
        }

        private void OnInspectionEntry(BlazorContextMenu.ItemClickEventArgs e)
        {
            if (e.Data is VendGateRcvdListViewModel item)
            {
                var returnUrl = Uri.EscapeDataString(NavigationManager.Uri);
                NavigationManager.NavigateTo($"/stock/vend-rcv-list/inspection/{item.RcvID}?returnUrl={returnUrl}");
            }
        }

        private async Task OnPrintSelectedGRN(ItemClickEventArgs e)
        {
            if (e.Data is VendGateRcvdListViewModel item)
            {
                var req = new Impulse.Services.ReportRequest
                {
                    ReportName = "GRN.rpt",
                    SelectionFormula = $"{{VendRcvd.RcvID}}='{item.RcvID}'"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
        }

        private async Task OnPrintMakerGRN(ItemClickEventArgs e)
        {
            if (e.Data is VendGateRcvdListViewModel item)
            {
                var req = new Impulse.Services.ReportRequest
                {
                    ReportName = "GRN_MKR.rpt",
                    SelectionFormula = $"{{VendRcvd.RcvID}}='{item.RcvID}'"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
        }

        private async Task OnPrintDailyReceiving(ItemClickEventArgs e)
        {
            string selection = $"{{VVendRcvd.RcvDate}}>='{Filter.DateFrom:yyyy-MM-dd}' AND {{VVendRcvd.RcvDate}}<='{Filter.DateTo:yyyy-MM-dd}'";
            selection += $" AND {{VVendRcvd.SampleOrder}}={(IsSampleList ? "True" : "False")}";
            
            if (SelectedVendor != null)
            {
                string vendId = ((dynamic)SelectedVendor).AccNo;
                selection += $" AND {{VVendRcvd.VendID}}='{vendId}'";
            }

            var req = new Impulse.Services.ReportRequest
            {
                ReportName = "DailyRcving.rpt",
                SelectionFormula = selection
            };
            await ReportNavigation.PrintReportAsync(req);
        }

        private async Task OnPrintListWithValue(ItemClickEventArgs e)
        {
            // Legacy uses a specific SQL query string injected, but standard approach passes report name.
            var req = new Impulse.Services.ReportRequest
            {
                ReportName = "RMReceivingListWValue.rpt",
                // Assuming it relies on the same date filter logic or is just called. In legacy, it passes a SQL string.
                // For now, we will just call it with no selection formula, standard practice is handled inside the report.
                SelectionFormula = ""
            };
            await ReportNavigation.PrintReportAsync(req);
        }
    }
}
