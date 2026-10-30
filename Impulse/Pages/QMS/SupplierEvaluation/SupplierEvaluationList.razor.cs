using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BlazorContextMenu;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Impulse.Components.Common;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Common;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.QMS.SupplierEvaluation
{
    public partial class SupplierEvaluationList : ComponentBase
    {
        [Inject] private ISupplierEvaluationDataAccess SupplierEvaluationDataAccess { get; set; } = default!;
        [Inject] private IEntityAttachmentService AttachmentService { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        private List<SupplierEvaluationListItemDto> evaluations = new();
        private List<SupplierDropdownDto> suppliers = new();
        private SupplierDropdownDto? selectedSupplier;
        private SupplierEvaluationFilterDto filter = new();

        private int selectedRangePreset = 5; // Default: Last 180 Days
        private bool isLoading = true;

        private AttachmentModal? attachmentModalRef;

        protected override async Task OnInitializedAsync()
        {
            ApplyDateRangePreset(selectedRangePreset);
            await LoadSuppliersAsync();
            await LoadDataAsync();
        }

        private async Task LoadSuppliersAsync()
        {
            try
            {
                suppliers = await SupplierEvaluationDataAccess.GetSuppliersListAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load suppliers: {ex.Message}");
            }
        }

        private async Task LoadDataAsync()
        {
            try
            {
                isLoading = true;
                evaluations = await SupplierEvaluationDataAccess.GetSupplierEvaluationsAsync(filter);

                // Populate attachment counts asynchronously
                foreach (var item in evaluations)
                {
                    try
                    {
                        item.AttachmentCount = await AttachmentService.GetCountAsync("SupplierEvaluation", item.AccNo);
                    }
                    catch
                    {
                        item.AttachmentCount = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Error", $"Failed to load evaluations: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private async Task RefreshDataAsync()
        {
            await LoadDataAsync();
        }

        private void OnRangePresetChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int preset))
            {
                selectedRangePreset = preset;
                ApplyDateRangePreset(preset);
                _ = LoadDataAsync();
            }
        }

        private void ApplyDateRangePreset(int preset)
        {
            DateTime today = DateTime.Today;
            filter.DateTo = today;

            switch (preset)
            {
                case 0: // Today
                    filter.DateFrom = today;
                    break;
                case 1: // Last 15 Days
                    filter.DateFrom = today.AddDays(-15);
                    break;
                case 2: // Last 30 Days
                    filter.DateFrom = today.AddDays(-30);
                    break;
                case 3: // Last 60 Days
                    filter.DateFrom = today.AddDays(-60);
                    break;
                case 4: // Last 90 Days
                    filter.DateFrom = today.AddDays(-90);
                    break;
                case 5: // Last 180 Days
                    filter.DateFrom = today.AddDays(-180);
                    break;
                case 6: // Custom
                    // Keep existing dates
                    break;
            }
        }

        private void OnDateChanged(ChangeEventArgs e, bool isFrom)
        {
            if (DateTime.TryParse(e.Value?.ToString(), out DateTime dt))
            {
                if (isFrom)
                    filter.DateFrom = dt;
                else
                    filter.DateTo = dt;

                _ = LoadDataAsync();
            }
        }

        private Task<IEnumerable<SupplierDropdownDto>> SearchSuppliers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Task.FromResult<IEnumerable<SupplierDropdownDto>>(suppliers.Take(100));
            }

            var query = searchText.Trim();
            var matches = suppliers.Where(s =>
                (!string.IsNullOrEmpty(s.AccNo) && s.AccNo.Contains(query, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(s.AccTitle) && s.AccTitle.Contains(query, StringComparison.OrdinalIgnoreCase))
            );

            return Task.FromResult<IEnumerable<SupplierDropdownDto>>(matches);
        }

        private async Task OnSupplierFilterChanged(SupplierDropdownDto? sup)
        {
            selectedSupplier = sup;
            filter.AccNo = sup?.AccNo;
            await LoadDataAsync();
        }

        private void HandleSearchKeyDown(KeyboardEventArgs e)
        {
            if (e.Key == "Enter")
            {
                _ = LoadDataAsync();
            }
        }

        private void ClearSearch()
        {
            filter.SearchText = string.Empty;
            _ = LoadDataAsync();
        }

        private void NavigateToHub()
        {
            NavigationManager.NavigateTo("/qms");
        }

        private void CreateNewEvaluation()
        {
            string returnUrl = Uri.EscapeDataString(NavigationManager.Uri);
            NavigationManager.NavigateTo($"/qms/supplier-evaluations/new?returnUrl={returnUrl}");
        }

        private void EditEvaluation(string accNo)
        {
            string returnUrl = Uri.EscapeDataString(NavigationManager.Uri);
            NavigationManager.NavigateTo($"/qms/supplier-evaluations/edit/{Uri.EscapeDataString(accNo)}?returnUrl={returnUrl}");
        }

        private void OpenAttachments(SupplierEvaluationListItemDto item)
        {
            attachmentModalRef?.Open(
                entityType: "SupplierEvaluation",
                entityId: item.AccNo,
                secondaryId: item.EntryID.ToString(),
                displayName: $"Supplier: {item.AccTitle} ({item.AccNo})"
            );
        }

        private async Task ShowOptionsMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("optionsBarMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private async Task PrintEvaluation(string accNo)
        {
            if (string.IsNullOrWhiteSpace(accNo)) return;

            try
            {
                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.QMS.SupplierEvaluation,
                    SelectionFormula = $"{{SupplierEvaluation.AccNo}} = '{accNo}'"
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print supplier evaluation: {ex.Message}");
            }
        }

        private async Task PrintEvaluationList()
        {
            try
            {
                var f = filter.DateFrom ?? DateTime.Today.AddDays(-180);
                var t = filter.DateTo ?? DateTime.Today;
                string sel = $"{{SupplierEvaluation.EntryDT}} in Date({f.Year}, {f.Month}, {f.Day}) to Date({t.Year}, {t.Month}, {t.Day})";
                if (!string.IsNullOrWhiteSpace(filter.AccNo))
                {
                    sel += $" AND {{SupplierEvaluation.AccNo}} = '{filter.AccNo}'";
                }

                await ReportNavigationService.PrintReportAsync(new ReportRequest
                {
                    ReportName = ReportNames.QMS.SupplierEvaluation,
                    SelectionFormula = sel
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(NotificationSeverity.Error, "Print Failed", $"Could not print evaluation list: {ex.Message}");
            }
        }
    }
}

