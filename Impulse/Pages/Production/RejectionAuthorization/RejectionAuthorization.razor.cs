using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Production;
using Impulse.Services;
using Impulse.Services.Production;
using Microsoft.AspNetCore.Components;

namespace Impulse.Pages.Production.RejectionAuthorization
{
    public partial class RejectionAuthorization : ComponentBase
    {
        [Inject] private IRejectionAuthorizationService AuthService { get; set; } = null!;
        [Inject] private IReportNavigationService ReportNavigationService { get; set; } = null!;
        [Inject] private INotificationService NotificationService { get; set; } = null!;
        [Inject] private NavigationManager NavigationManager { get; set; } = null!;

        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private RejectionAuthFilterModel Filter { get; set; } = new();
        private int FactoryMakerId { get; set; }
        private int ActiveTab { get; set; }
        private bool IsLoading { get; set; }
        private bool IsProcessing { get; set; }
        private bool IsCategoryDropdownOpen { get; set; }

        private List<RejectionAuthMakerLookup> Makers { get; set; } = new();
        private List<RejectionAuthEmployeeLookup> Employees { get; set; } = new();
        private List<RejectionAuthCategoryLookup> Categories { get; set; } = new();

        private RejectionAuthMakerLookup? SelectedMaker { get; set; }
        private RejectionAuthEmployeeLookup? SelectedEmployee { get; set; }

        private List<RejectionAuthItemModel> PendingItems { get; set; } = new();
        private List<RejectionAuthItemModel> HistoryItems { get; set; } = new();
        private HashSet<int> CheckedIds { get; set; } = new();

        private string CategorySummaryText => !Filter.CategoryIds.Any()
            ? "<All Categories>"
            : string.Join(", ", Categories.Where(c => Filter.CategoryIds.Contains(c.CatID)).Select(c => c.Description));

        protected override async Task OnInitializedAsync()
        {
            ApplyDatePreset(Filter.DateRangeIndex);
            await LoadLookupsAsync();
            await RefreshAsync();
        }

        private async Task LoadLookupsAsync()
        {
            try
            {
                FactoryMakerId = await AuthService.GetFactoryMakerIdAsync();
                var makersList = await AuthService.GetMakersAsync(FactoryMakerId);
                Makers = new List<RejectionAuthMakerLookup>
                {
                    new() { VendID = 0, VenderName = "<Show All>" }
                };
                Makers.AddRange(makersList);
                SelectedMaker = Makers.First();

                var empList = await AuthService.GetEmployeesAsync();
                Employees = new List<RejectionAuthEmployeeLookup>
                {
                    new() { EmpID = "0", Name = "<Show All>" }
                };
                Employees.AddRange(empList);
                SelectedEmployee = Employees.First();

                Categories = await AuthService.GetCategoriesAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error", $"Failed to load lookups: {ex.Message}");
            }
        }

        private async Task RefreshAsync()
        {
            IsLoading = true;
            IsCategoryDropdownOpen = false;
            try
            {
                CheckedIds.Clear();
                PendingItems = await AuthService.GetPendingRejectionsAsync(Filter);
                HistoryItems = await AuthService.GetHistoryAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Load Error", $"Failed to load rejections: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task OnShowEmployeesChanged(ChangeEventArgs e)
        {
            Filter.ShowEmployees = e.Value is bool b && b;
            Filter.VendID = 0;
            Filter.EmpID = "0";
            SelectedMaker = Makers.FirstOrDefault();
            SelectedEmployee = Employees.FirstOrDefault();
            await RefreshAsync();
        }

        private Task<IEnumerable<RejectionAuthMakerLookup>> SearchMakers(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<RejectionAuthMakerLookup>>(Makers);

            return Task.FromResult<IEnumerable<RejectionAuthMakerLookup>>(
                Makers.Where(m => m.DisplayName.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnMakerChanged(RejectionAuthMakerLookup? maker)
        {
            SelectedMaker = maker ?? Makers.FirstOrDefault();
            Filter.VendID = SelectedMaker?.VendID ?? 0;
            await RefreshAsync();
        }

        private Task<IEnumerable<RejectionAuthEmployeeLookup>> SearchEmployees(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Task.FromResult<IEnumerable<RejectionAuthEmployeeLookup>>(Employees);

            return Task.FromResult<IEnumerable<RejectionAuthEmployeeLookup>>(
                Employees.Where(e =>
                    e.EmpID.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.FName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Designation.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList());
        }

        private async Task OnEmployeeChanged(RejectionAuthEmployeeLookup? emp)
        {
            SelectedEmployee = emp ?? Employees.FirstOrDefault();
            Filter.EmpID = SelectedEmployee?.EmpID ?? "0";
            await RefreshAsync();
        }

        private async Task OnDateRangeChanged(ChangeEventArgs e)
        {
            if (int.TryParse(e.Value?.ToString(), out int idx))
            {
                Filter.DateRangeIndex = idx;
                ApplyDatePreset(idx);
                if (idx != 5)
                {
                    await RefreshAsync();
                }
            }
        }

        private void ApplyDatePreset(int idx)
        {
            DateTime today = DateTime.Today;
            switch (idx)
            {
                case 0:
                    Filter.DtFrom = today;
                    Filter.DtTo = today;
                    break;
                case 1:
                    Filter.DtFrom = today.AddDays(-15);
                    Filter.DtTo = today;
                    break;
                case 2:
                    Filter.DtFrom = today.AddDays(-30);
                    Filter.DtTo = today;
                    break;
                case 3:
                    Filter.DtFrom = today.AddDays(-60);
                    Filter.DtTo = today;
                    break;
                case 4:
                    Filter.DtFrom = today.AddDays(-90);
                    Filter.DtTo = today;
                    break;
            }
        }

        private void ToggleCategoryDropdown() => IsCategoryDropdownOpen = !IsCategoryDropdownOpen;

        private void ToggleAllCategories(ChangeEventArgs e)
        {
            if (e.Value is bool b && b)
            {
                Filter.CategoryIds.Clear();
            }
        }

        private void ToggleCategory(string catId, bool isChecked)
        {
            if (isChecked)
            {
                if (!Filter.CategoryIds.Contains(catId))
                    Filter.CategoryIds.Add(catId);
            }
            else
            {
                Filter.CategoryIds.Remove(catId);
            }
        }

        private void SelectAllRows()
        {
            CheckedIds = PendingItems.Select(i => i.EntryID).ToHashSet();
        }

        private void ClearAllRows()
        {
            CheckedIds.Clear();
        }

        private void ToggleRowCheck((int EntryId, bool IsChecked) args)
        {
            if (args.IsChecked)
                CheckedIds.Add(args.EntryId);
            else
                CheckedIds.Remove(args.EntryId);
        }

        private async Task DeductSelectedAsync() => await ApplyDecisionAsync(deduct: true);

        private async Task IgnoreSelectedAsync() => await ApplyDecisionAsync(deduct: false);

        private async Task ApplyDecisionAsync(bool deduct)
        {
            if (!CheckedIds.Any())
            {
                NotificationService.ShowWarning("No Selection", "Please select at least one rejection entry.");
                return;
            }

            IsProcessing = true;
            try
            {
                int count = await AuthService.UpdateDeductFromBillingAsync(CheckedIds.ToList(), deduct);
                NotificationService.ShowSuccess("Updated", $"{count} entry(s) marked as {(deduct ? "Deduct" : "Ignored")}.");
                await RefreshAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Update Error", ex.Message);
            }
            finally
            {
                IsProcessing = false;
            }
        }

        private async Task UpdateRateAsync((int EntryId, int NewRate) args)
        {
            if (args.NewRate < 0)
            {
                NotificationService.ShowWarning("Invalid Rate", "Rate cannot be negative.");
                return;
            }

            try
            {
                bool ok = await AuthService.UpdateRateAsync(args.EntryId, args.NewRate);
                if (ok)
                {
                    var item = PendingItems.FirstOrDefault(i => i.EntryID == args.EntryId);
                    if (item != null)
                    {
                        item.Rate = args.NewRate;
                    }
                    NotificationService.ShowSuccess("Rate Updated", "Rejection rate updated successfully.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Rate Update Error", ex.Message);
            }
        }

        private async Task PrintReportAsync()
        {
            var conditions = new List<string>
            {
                "IsNull({VMakerRepair.VchrNo})",
                "{VMakerRepair.ShowInBilling} = true",
                "IsNull({VMakerRepair.DeductFromBilling})",
                $"{{VMakerRepair.DT}} in Date({Filter.DtFrom.Year}, {Filter.DtFrom.Month}, {Filter.DtFrom.Day}) to Date({Filter.DtTo.Year}, {Filter.DtTo.Month}, {Filter.DtTo.Day})"
            };

            if (Filter.ShowEmployees)
            {
                if (!string.IsNullOrWhiteSpace(Filter.EmpID) && Filter.EmpID != "0")
                {
                    conditions.Add($"{{VMakerRepair.EmpID}} = '{Filter.EmpID.Replace("'", "''")}'");
                }
            }
            else if (Filter.VendID > 0)
            {
                conditions.Add($"{{VMakerRepair.VendID}} = {Filter.VendID}");
            }

            if (Filter.CategoryIds.Any())
            {
                var cats = Filter.CategoryIds.Select(c => $"'{c.Replace("'", "''")}'");
                conditions.Add($"{{VMakerRepair.CatID}} in [{string.Join(",", cats)}]");
            }

            if (!string.IsNullOrWhiteSpace(Filter.LotNo))
            {
                conditions.Add($"{{VMakerRepair.LotNo}} = '{Filter.LotNo.Trim().Replace("'", "''")}'");
            }

            await ReportNavigationService.PrintReportAsync(new ReportRequest
            {
                ReportName = "Rejection_Authorization.rpt",
                SelectionFormula = string.Join(" and ", conditions)
            });
        }

        private void GoBack()
        {
            NavigationManager.NavigateTo(string.IsNullOrWhiteSpace(ReturnUrl) ? "/production" : ReturnUrl);
        }
    }
}
