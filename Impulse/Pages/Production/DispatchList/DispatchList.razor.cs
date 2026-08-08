using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Radzen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Production.DispatchList
{
    public partial class DispatchList : ComponentBase
    {
        [Inject] private IDispatchListService DispatchService { get; set; } = default!;
        [Inject] private NavigationManager Navigation { get; set; } = default!;
        [Inject] private Radzen.NotificationService NotificationService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private AuthenticationStateProvider AuthStateProvider { get; set; } = default!;

        // Lookups & Grid
        public List<LookupItemString> Customers { get; set; } = new List<LookupItemString>();
        public List<DispatchListSummaryItem> DispatchLists { get; set; } = new List<DispatchListSummaryItem>();

        // Filter State
        public LookupItemString? SelectedCustomer { get; set; } = null;
        public DispatchListSearchFilter Filter { get; set; } = new DispatchListSearchFilter();

        public bool IsLoading { get; set; } = false;
        public DispatchListSummaryItem? ContextRowItem { get; set; } = null;

        // Computed Properties
        public int TotalCartonsSum => DispatchLists.Sum(x => x.TotalCartons);
        public decimal TotalPiecesSum => DispatchLists.Sum(x => x.TotalQty);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Customers = await DispatchService.GetCustomersAsync();

                Filter.DateRangeIndex = 0;
                Filter.DtFrom = DateTime.Today;
                Filter.DtTo = DateTime.Today;

                await LoadDispatchListsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Initialization Error",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
        }

        public async Task<IEnumerable<LookupItemString>> SearchCustomers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return await Task.FromResult(Customers);
            }

            return await Task.FromResult(Customers.Where(c =>
                c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
        }

        public async Task OnCustomerChanged(LookupItemString? cust)
        {
            SelectedCustomer = cust;
            Filter.CustCode = cust?.Id ?? string.Empty;
            await LoadDispatchListsAsync();
        }

        public async Task OnDateRangeChanged()
        {
            switch (Filter.DateRangeIndex)
            {
                case 0:
                    Filter.DtFrom = DateTime.Today;
                    Filter.DtTo = DateTime.Today;
                    break;
                case 1:
                    Filter.DtFrom = DateTime.Today.AddDays(-7);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 2:
                    Filter.DtFrom = DateTime.Today.AddDays(-15);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 3:
                    Filter.DtFrom = DateTime.Today.AddDays(-30);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 4:
                    Filter.DtFrom = DateTime.Today.AddDays(-60);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 5:
                    Filter.DtFrom = DateTime.Today.AddDays(-90);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 6:
                    Filter.DtFrom = DateTime.Today.AddYears(-1);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 7:
                    Filter.DtFrom = DateTime.Today.AddYears(-5);
                    Filter.DtTo = DateTime.Today;
                    break;
                case 8:
                    // Keep user custom range
                    break;
            }

            await LoadDispatchListsAsync();
        }

        public async Task LoadDispatchListsAsync()
        {
            IsLoading = true;
            try
            {
                DispatchLists = await DispatchService.GetDispatchListsAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Failed to Load Dispatches",
                    Detail = ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void EditDispatch(DispatchListSummaryItem item)
        {
            Navigation.NavigateTo($"/production/create-dispatch-list?entryId={item.EntryID}");
        }

        public void ContextEdit(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem != null)
            {
                EditDispatch(ContextRowItem);
            }
        }

        public async Task ContextFinalize(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            if (ContextRowItem.Finalyzed)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Info,
                    Summary = "Already Finalized",
                    Detail = "This dispatch list is already finalized.",
                    Duration = 3000
                });
                return;
            }

            try
            {
                string userName = await GetCurrentUserName();
                string machineName = Environment.MachineName;

                var result = await DispatchService.FinalizeDispatchListAsync(ContextRowItem.EntryID, userName, machineName);

                if (result.Success)
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Success,
                        Summary = "Dispatch Finalized",
                        Detail = result.Message,
                        Duration = 4000
                    });

                    await LoadDispatchListsAsync();
                }
                else
                {
                    NotificationService.Notify(new Radzen.NotificationMessage
                    {
                        Severity = Radzen.NotificationSeverity.Warning,
                        Summary = "Finalization Blocked",
                        Detail = result.Message,
                        Duration = 5000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new Radzen.NotificationMessage
                {
                    Severity = Radzen.NotificationSeverity.Error,
                    Summary = "Error Finalizing",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Reporting Handlers
        // ─────────────────────────────────────────────────────────────

        public async Task PrintInnerwise(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "DispatchListInnerwise.rpt",
                SelectionFormula = $"{{DispatchList.EntryID}} = {ContextRowItem.EntryID}"
            });
        }

        public async Task PrintFinal(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "DispatchListFinal.rpt",
                SelectionFormula = $"{{DispatchList.EntryID}} = {ContextRowItem.EntryID}"
            });
        }

        public async Task PrintWithValue(ItemClickEventArgs args)
        {
            ResolveRowItem(args);
            if (ContextRowItem == null) return;

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "DispatchListInnerwiseWithValue.rpt",
                SelectionFormula = $"{{DispatchList.DispatchListEntryID}} = {ContextRowItem.EntryID}"
            });
        }

        private void ResolveRowItem(ItemClickEventArgs args)
        {
            if (args.Data is DispatchListSummaryItem item)
            {
                ContextRowItem = item;
            }
            else
            {
                ContextRowItem = null;
            }
        }

        private async Task<string> GetCurrentUserName()
        {
            try
            {
                var authState = await AuthStateProvider.GetAuthenticationStateAsync();
                var user = authState.User;
                return user.Identity?.IsAuthenticated == true
                    ? user.Identity.Name ?? "System"
                    : "System";
            }
            catch
            {
                return "System";
            }
        }
    }
}
