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
    public partial class FixedAssetsList : ComponentBase
    {
        [Inject] private IFixedAssetProfileDataAccess DataAccess { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = null!;
        [Inject] private NotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;
        [Inject] private IBlazorContextMenuService ContextMenuService { get; set; } = null!;
        [Inject] private AuthenticationStateProvider AuthenticationStateProvider { get; set; } = null!;

        [SupplyParameterFromQuery(Name = "categoryId")]
        public string? CategoryId { get; set; }

        [SupplyParameterFromQuery(Name = "filter")]
        public string? FilterParam { get; set; }

        private FixedAssetLookupOptions lookup = new();
        private List<FixedAssetListRow> rows = new();
        private string? selectedMasterId;
        private string searchText = "";
        private bool includeInactive;
        private bool isBusy;
        private FixedAssetListRow? selectedRow;
        private FixedAssetListRow? modalTarget;

        // Paging
        private int pageSize = 25;
        private int currentPage = 1;
        private bool showAll => pageSize <= 0;
        private List<FixedAssetListRow> pagedRows => showAll
            ? rows
            : rows.Skip((currentPage - 1) * pageSize).Take(pageSize).ToList();
        private int totalPages => showAll ? 1 : Math.Max(1, (int)Math.Ceiling(rows.Count / (double)pageSize));
        private int pagerStart => rows.Count == 0 || showAll ? Math.Min(1, rows.Count) : (currentPage - 1) * pageSize + 1;
        private int pagerEnd => showAll ? rows.Count : Math.Min(currentPage * pageSize, rows.Count);

        // Modals
        private bool showCategoryModal;
        private string newCategoryName = "";
        private decimal newCategoryRate;

        private bool showDeleteModal;

        private bool showActiveModal;
        private DateTime statusEffectiveDate = DateTime.Today;
        private string statusRemarks = "";

        private bool showDateReportModal;
        private DateTime reportFromDate = new DateTime(DateTime.Today.Year, 1, 1);
        private DateTime reportToDate = DateTime.Today;

        private bool initialized;

        protected override async Task OnInitializedAsync()
        {
            lookup = await DataAccess.GetLookupOptionsAsync();

            if (!string.IsNullOrWhiteSpace(CategoryId))
            {
                selectedMasterId = CategoryId;
            }

            if (string.Equals(FilterParam, "inactive", StringComparison.OrdinalIgnoreCase))
            {
                includeInactive = true;
            }

            await LoadAsync();
            initialized = true;
        }

        protected override async Task OnParametersSetAsync()
        {
            if (initialized && CategoryId != selectedMasterId)
            {
                selectedMasterId = CategoryId;
                await LoadAsync();
            }
        }

        private async Task HandleCategoryChanged(ChangeEventArgs args)
        {
            selectedMasterId = args.Value?.ToString();
            CategoryId = selectedMasterId;
            UpdateUrlWithCategory();
            await LoadAsync();
        }

        private async Task HandleSearchKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                await LoadAsync();
            }
        }

        private async Task ToggleInactiveFilter(ChangeEventArgs args)
        {
            includeInactive = (bool)(args.Value ?? false);
            await LoadAsync();
        }

        private void UpdateUrlWithCategory()
        {
            var url = !string.IsNullOrWhiteSpace(selectedMasterId)
                ? $"/fixedassets/list?categoryId={Uri.EscapeDataString(selectedMasterId)}"
                : "/fixedassets/list";
            NavigationManager.NavigateTo(url, replace: true);
        }

        private async Task LoadAsync()
        {
            try
            {
                isBusy = true;
                int? masterId = int.TryParse(selectedMasterId, out var id) ? id : null;
                var list = await DataAccess.GetListAsync(masterId, searchText, includeInactive);

                if (!string.IsNullOrWhiteSpace(FilterParam))
                {
                    if (string.Equals(FilterParam, "active", StringComparison.OrdinalIgnoreCase))
                    {
                        list = list.Where(x => !x.InActive && !x.Disposed).ToList();
                    }
                    else if (string.Equals(FilterParam, "disposed", StringComparison.OrdinalIgnoreCase))
                    {
                        list = list.Where(x => x.Disposed && x.DisposedType == 1).ToList();
                    }
                    else if (string.Equals(FilterParam, "writtenoff", StringComparison.OrdinalIgnoreCase))
                    {
                        list = list.Where(x => x.Disposed && x.DisposedType == 0).ToList();
                    }
                }

                rows = list.ToList();
                currentPage = 1;

                if (selectedRow != null)
                {
                    selectedRow = rows.FirstOrDefault(x => x.EntryId == selectedRow.EntryId);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error Loading Assets", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        private void SelectRow(FixedAssetListRow row)
        {
            selectedRow = row;
        }

        private void GoToPage(int page)
        {
            currentPage = Math.Clamp(page, 1, totalPages);
        }

        private void HandlePageSizeChanged(ChangeEventArgs args)
        {
            pageSize = int.TryParse(args.Value?.ToString(), out var size) ? size : 25;
            GoToPage(1);
        }

        private async Task OpenOptionsBarMenu(MouseEventArgs e)
        {
            await ContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private void OpenNewAsset()
        {
            var url = !string.IsNullOrWhiteSpace(selectedMasterId)
                ? $"/fixedassets/profile?categoryId={Uri.EscapeDataString(selectedMasterId)}&returnUrl={Uri.EscapeDataString("/fixedassets/list")}"
                : $"/fixedassets/profile?returnUrl={Uri.EscapeDataString("/fixedassets/list")}";
            NavigationManager.NavigateTo(url);
        }

        private void OpenAsset(int entryId)
        {
            var url = $"/fixedassets/profile/{entryId}?returnUrl={Uri.EscapeDataString("/fixedassets/list")}";
            NavigationManager.NavigateTo(url);
        }

        private void CopyAsset(int entryId)
        {
            var url = $"/fixedassets/profile?copyFrom={entryId}&returnUrl={Uri.EscapeDataString("/fixedassets/list")}";
            NavigationManager.NavigateTo(url);
        }

        private void OpenRegister()
        {
            NavigationManager.NavigateTo("/fixedassets/register");
        }

        private void OpenAdjustment(string assetId)
        {
            NavigationManager.NavigateTo($"/fixedassets/adjustment/{Uri.EscapeDataString(assetId)}?returnUrl={Uri.EscapeDataString("/fixedassets/list")}");
        }

        private void OpenMove(string assetId)
        {
            NavigationManager.NavigateTo($"/fixedassets/move/{Uri.EscapeDataString(assetId)}?returnUrl={Uri.EscapeDataString("/fixedassets/list")}");
        }

        private void OpenDisposal(string assetId)
        {
            NavigationManager.NavigateTo($"/fixedassets/disposal/{Uri.EscapeDataString(assetId)}?returnUrl={Uri.EscapeDataString("/fixedassets/list")}");
        }

        // Category Modal
        private void OpenAddCategoryModal()
        {
            newCategoryName = "";
            newCategoryRate = 0;
            showCategoryModal = true;
        }

        private async Task SaveCategoryAsync()
        {
            if (string.IsNullOrWhiteSpace(newCategoryName)) return;

            try
            {
                isBusy = true;
                await DataAccess.SaveMasterCategoryAsync(newCategoryName.Trim(), newCategoryRate);
                showCategoryModal = false;
                lookup = await DataAccess.GetLookupOptionsAsync();
                NotificationService.Notify(NotificationSeverity.Success, "Success", $"Category '{newCategoryName}' added successfully.", 4000);
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

        // Delete Modal
        private void OpenDeleteModal(FixedAssetListRow row)
        {
            modalTarget = row;
            showDeleteModal = true;
        }

        private async Task ConfirmDeleteAsync()
        {
            if (modalTarget == null) return;

            try
            {
                isBusy = true;
                var assetId = modalTarget.AssetId;
                await DataAccess.DeleteAssetAsync(modalTarget.EntryId, assetId);
                showDeleteModal = false;
                if (selectedRow?.EntryId == modalTarget.EntryId)
                {
                    selectedRow = null;
                }
                NotificationService.Notify(NotificationSeverity.Success, "Deleted", $"Asset '{assetId}' was permanently deleted.", 4000);
                await LoadAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Delete Failed", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        // Active Status Modal
        private void OpenActiveStatusModal(FixedAssetListRow row)
        {
            modalTarget = row;
            statusEffectiveDate = DateTime.Today;
            statusRemarks = "";
            showActiveModal = true;
        }

        private async Task ConfirmActiveStatusAsync()
        {
            if (modalTarget == null || string.IsNullOrWhiteSpace(statusRemarks)) return;

            try
            {
                isBusy = true;
                var authState = await AuthenticationStateProvider.GetAuthenticationStateAsync();
                var userName = authState.User.Identity?.Name ?? authState.User.FindFirstValue(ClaimTypes.Name) ?? "Administrator";
                bool newInactive = !modalTarget.InActive;

                await DataAccess.SetActiveStatusAsync(modalTarget.EntryId, modalTarget.AssetId, newInactive, statusEffectiveDate, statusRemarks, userName, Environment.MachineName);
                showActiveModal = false;
                NotificationService.Notify(NotificationSeverity.Success, "Status Changed", $"Asset '{modalTarget.AssetId}' marked as {(newInactive ? "Inactive" : "Active")}.", 4000);
                await LoadAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Status Update Failed", ex.Message, 5000);
            }
            finally
            {
                isBusy = false;
            }
        }

        // Reports
        private async Task PrintListAsync()
        {
            try
            {
                var req = new ReportRequest
                {
                    ReportName = ReportNames.FixedAssets.FixedAssetsReport,
                    SelectionFormula = string.IsNullOrWhiteSpace(selectedMasterId) ? "" : $"{{FixedAssets_Master.EntryID}} = {selectedMasterId}"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }

        private async Task PrintRegisterReportAsync()
        {
            try
            {
                var req = new ReportRequest
                {
                    ReportName = ReportNames.FixedAssets.FixedAssetsRegister,
                    SelectionFormula = string.IsNullOrWhiteSpace(selectedMasterId) ? "" : $"{{FixedAssets_Master.EntryID}} = {selectedMasterId}"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }

        private void OpenDatewiseReportModal()
        {
            showDateReportModal = true;
        }

        private async Task PrintDatewiseReportAsync()
        {
            try
            {
                showDateReportModal = false;
                var formula = $"{{FixedAssets.AcquiredDT}} >= Date({reportFromDate.Year},{reportFromDate.Month},{reportFromDate.Day}) and {{FixedAssets.AcquiredDT}} <= Date({reportToDate.Year},{reportToDate.Month},{reportToDate.Day})";
                if (!string.IsNullOrWhiteSpace(selectedMasterId))
                {
                    formula += $" and {{FixedAssets_Master.EntryID}} = {selectedMasterId}";
                }

                var req = new ReportRequest
                {
                    ReportName = ReportNames.FixedAssets.FixedAssetsDatewise,
                    SelectionFormula = formula,
                    FormulaValues = new Dictionary<string, object?>
                    {
                        ["DateRange"] = $"{reportFromDate:dd-MMM-yyyy} to {reportToDate:dd-MMM-yyyy}"
                    }
                };
                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }

        private async Task PrintReceiptSlip(string assetId)
        {
            try
            {
                var req = new ReportRequest
                {
                    ReportName = ReportNames.FixedAssets.FixedAssetsRepairSlip,
                    SelectionFormula = $"{{VFixedAssets.AssetID}} = '{assetId.Replace("'", "''")}'"
                };
                await ReportNavigation.PrintReportAsync(req);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }

        private async Task PrintReportByName(string reportFileName)
        {
            try
            {
                await ReportNavigation.PrintReportAsync(new ReportRequest { ReportName = reportFileName });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", ex.Message, 5000);
            }
        }

        private static string GetGatePassBadgeClass(string? status)
        {
            return status?.Trim().ToLowerInvariant() switch
            {
                "issued" => "bg-warning-subtle text-warning-emphasis border border-warning-subtle",
                "available" => "bg-success-subtle text-success border border-success-subtle",
                _ => "bg-light text-muted border"
            };
        }

        private static string GetReceiptNoteBadgeClass(string? status)
        {
            return status?.Trim().ToLowerInvariant() switch
            {
                "received" => "bg-success-subtle text-success border border-success-subtle",
                "rejected" => "bg-danger-subtle text-danger border border-danger-subtle",
                "case closed" => "bg-secondary-subtle text-secondary border border-secondary-subtle",
                "under process" => "bg-info-subtle text-info-emphasis border border-info-subtle",
                _ => "bg-light text-muted border"
            };
        }
    }
}

