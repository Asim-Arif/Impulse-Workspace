using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ReportModels;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.FixedAssets
{
    public partial class FixedAssetRegister : ComponentBase
    {
        [Inject] private IFixedAssetRegisterDataAccess DataAccess { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = null!;

        [SupplyParameterFromQuery(Name = "from")]
        public string? QueryFromDate { get; set; }

        [SupplyParameterFromQuery(Name = "to")]
        public string? QueryToDate { get; set; }

        private DateTime dateFrom = new(DateTime.Today.Year, 7, 1);
        private DateTime dateTo = new(DateTime.Today.Year + 1, 6, 30);

        private FixedAssetRegisterFilterOptions options = new();
        private List<FixedAssetRegisterRow> allRows = new();
        private string selectedCategory = "";
        private string selectedDeptId = "";
        private bool newAssetsOnly;
        private string? selectedPeriodText;
        private string searchText = "";

        private bool isSaved;
        private bool isBusy;
        private bool showDeleteConfirmModal;

        // Paging
        private int pageSize = 50;
        private int currentPage = 1;
        private bool showAll => pageSize <= 0;

        private List<FixedAssetRegisterRow> filteredRows
        {
            get
            {
                var q = allRows.AsEnumerable();

                if (!string.IsNullOrWhiteSpace(selectedCategory))
                {
                    q = q.Where(x => string.Equals(x.MasterDescription, selectedCategory, StringComparison.OrdinalIgnoreCase));
                }

                if (!string.IsNullOrWhiteSpace(selectedDeptId))
                {
                    q = q.Where(x => x.DeptId == selectedDeptId);
                }

                if (newAssetsOnly)
                {
                    q = q.Where(x => x.AcquiredDate.HasValue && x.AcquiredDate.Value >= dateFrom && x.AcquiredDate.Value <= dateTo);
                }

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    var term = searchText.Trim();
                    q = q.Where(x =>
                        (!string.IsNullOrEmpty(x.AssetId) && x.AssetId.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(x.Particulars) && x.Particulars.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(x.DeptName) && x.DeptName.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(x.Location) && x.Location.Contains(term, StringComparison.OrdinalIgnoreCase)));
                }

                return q.ToList();
            }
        }

        private List<FixedAssetRegisterRow> pagedRows => showAll
            ? filteredRows
            : filteredRows.Skip((currentPage - 1) * pageSize).Take(pageSize).ToList();

        private int totalPages => showAll ? 1 : Math.Max(1, (int)Math.Ceiling(filteredRows.Count / (double)pageSize));
        private int pagerStart => filteredRows.Count == 0 || showAll ? Math.Min(1, filteredRows.Count) : (currentPage - 1) * pageSize + 1;
        private int pagerEnd => showAll ? filteredRows.Count : Math.Min(currentPage * pageSize, filteredRows.Count);

        // Totals
        private decimal totalCost => filteredRows.Sum(x => x.Cost);
        private decimal totalBookValue => filteredRows.Sum(x => x.BookValue);
        private decimal totalDisposedValue => filteredRows.Sum(x => x.DisposedValue);
        private decimal totalOpeningBalance => filteredRows.Sum(x => x.OpeningBalance);
        private decimal totalAccountingYear => filteredRows.Sum(x => x.AccountingYear);
        private decimal totalAccumulatedBalance => filteredRows.Sum(x => x.AccumulatedBalance);
        private decimal totalWrittenDownValue => filteredRows.Sum(x => x.WrittenDownValue);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                isBusy = true;
                options = await DataAccess.GetFilterOptionsAsync();

                if (!string.IsNullOrWhiteSpace(QueryFromDate) && DateTime.TryParse(QueryFromDate, out var qf))
                {
                    dateFrom = qf;
                }
                else
                {
                    dateFrom = options.DefaultDateFrom;
                }

                if (!string.IsNullOrWhiteSpace(QueryToDate) && DateTime.TryParse(QueryToDate, out var qt))
                {
                    dateTo = qt;
                }
                else
                {
                    dateTo = options.DefaultDateTo;
                }

                await RefreshDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Initialization Error", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task RefreshDataAsync()
        {
            try
            {
                isBusy = true;
                var res = await DataAccess.LoadRegisterDataAsync(dateFrom, dateTo);
                isSaved = res.isSaved;
                allRows = res.rows ?? new List<FixedAssetRegisterRow>();
                currentPage = 1;
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Calculation Error", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task HandleFromDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                dateFrom = dt;
                selectedPeriodText = null;
                await RefreshDataAsync();
            }
        }

        private async Task HandleToDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out var dt))
            {
                dateTo = dt;
                selectedPeriodText = null;
                await RefreshDataAsync();
            }
        }

        private async Task HandlePeriodSelected(ChangeEventArgs e)
        {
            selectedPeriodText = e.Value?.ToString();
            var period = options.Periods.FirstOrDefault(x => x.DisplayText == selectedPeriodText);
            if (period != null)
            {
                dateFrom = period.DateFrom;
                dateTo = period.DateTo;
                await RefreshDataAsync();
            }
        }

        private void HandleCategoryFilterChanged(ChangeEventArgs e)
        {
            selectedCategory = e.Value?.ToString() ?? "";
            currentPage = 1;
        }

        private void HandleDeptFilterChanged(ChangeEventArgs e)
        {
            selectedDeptId = e.Value?.ToString() ?? "";
            currentPage = 1;
        }

        private void HandleNewAssetsOnlyChanged(ChangeEventArgs e)
        {
            newAssetsOnly = (bool)(e.Value ?? false);
            currentPage = 1;
        }

        private void GoToPage(int p)
        {
            currentPage = Math.Clamp(p, 1, totalPages);
        }

        private void HandlePageSizeChanged(ChangeEventArgs e)
        {
            pageSize = int.TryParse(e.Value?.ToString(), out var sz) ? sz : 50;
            GoToPage(1);
        }

        private async Task OpenReportsMenu(MouseEventArgs e)
        {
            await ContextMenuService.ShowMenu("registerReportsMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task HandleSaveRegisterAsync()
        {
            if (allRows.Count == 0)
            {
                NotificationService.Notify(NotificationSeverity.Warning, "Notice", "No asset register rows to save.", 4000);
                return;
            }

            try
            {
                isBusy = true;
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var userName = authState.User.Identity?.Name ?? authState.User.FindFirstValue(ClaimTypes.Name) ?? "Administrator";

                var res = await DataAccess.SaveRegisterAsync(dateFrom, dateTo, allRows, userName, Environment.MachineName);
                if (res.Success)
                {
                    isSaved = true;
                    NotificationService.Notify(NotificationSeverity.Success, "Saved", res.Message ?? "Depreciation register run saved successfully.", 4000);
                    await RefreshDataAsync();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Save Failed", res.Message ?? "Failed to save register run.", 5000);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task HandleDeleteRegisterAsync()
        {
            try
            {
                isBusy = true;
                var res = await DataAccess.DeleteRegisterAsync(dateFrom);
                showDeleteConfirmModal = false;
                if (res.success)
                {
                    isSaved = false;
                    NotificationService.Notify(NotificationSeverity.Success, "Deleted", res.message, 4000);
                    await RefreshDataAsync();
                }
                else
                {
                    NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", res.message, 5000);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private async Task PrintRegisterReport(string reportName)
        {
            try
            {
                var req = new ReportRequest
                {
                    ReportName = reportName,
                    Parameters = new Dictionary<string, object>
                    {
                        { "@DTFrom", dateFrom },
                        { "@DTTo", dateTo }
                    },
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "DateRange", $"'{dateFrom:dd-MMM-yyyy} to {dateTo:dd-MMM-yyyy}'" }
                    }
                };

                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }
    }
}

