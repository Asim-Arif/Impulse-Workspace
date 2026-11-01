using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Budgeting;
using DataAccessLibrary.Models.ViewModels.Budgeting;
using Impulse.Constants;
using Impulse.Models;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Budgeting
{
    public partial class IncomeBudgetEntry : ComponentBase
    {
        [Inject] public IBudgetDataAccess BudgetData { get; set; } = default!;
        [Inject] public NavigationManager Navigation { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigation { get; set; } = default!;

        protected bool isRevision => Navigation.Uri.Contains("entry-revision", StringComparison.OrdinalIgnoreCase);
        protected bool isLoading = true;
        protected bool showRevisionModal = false;

        protected List<FinancialYearOption> financialYears = new();
        protected string selectedFinYear = "";

        protected List<MasterHeadOption> masterHeads = new();
        protected MasterHeadOption? SelectedMasterHead;
        protected string selectedHeadCode => SelectedMasterHead?.Code ?? "0";

        protected string searchQuery = "";

        protected List<MonthwiseBudgetRow> rows = new();
        protected MonthwiseBudgetTotals totals = new();

        protected override async Task OnInitializedAsync()
        {
            financialYears = BudgetData.GetFinancialYears(2010);
            if (financialYears.Count > 0)
            {
                selectedFinYear = financialYears[0].Code;
            }

            masterHeads = await BudgetData.GetMasterHeadsAsync(isExpenses: false);
            SelectedMasterHead = masterHeads.FirstOrDefault(h => h.Code == "0") ?? masterHeads.FirstOrDefault();

            await LoadBudgetDataAsync();
        }

        protected override async Task OnParametersSetAsync()
        {
            if (rows.Count > 0)
            {
                await LoadBudgetDataAsync();
            }
        }

        protected async Task<IEnumerable<MasterHeadOption>> SearchMasterHeads(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return masterHeads;

            var q = searchText.Trim();
            return await Task.FromResult(
                masterHeads.Where(h =>
                    h.Code.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    h.Head.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    h.DisplayText.Contains(q, StringComparison.OrdinalIgnoreCase))
            );
        }

        protected async Task OnMasterHeadChanged(MasterHeadOption? selected)
        {
            SelectedMasterHead = selected ?? masterHeads.FirstOrDefault(h => h.Code == "0");
            await LoadBudgetDataAsync();
        }

        protected async Task OnFinYearChanged(ChangeEventArgs e)
        {
            selectedFinYear = e.Value?.ToString() ?? "";
            await LoadBudgetDataAsync();
        }

        protected void OnSearchChanged(ChangeEventArgs e)
        {
            searchQuery = e.Value?.ToString() ?? "";
            CalculateTotals();
        }

        protected IEnumerable<MonthwiseBudgetRow> FilteredRows
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchQuery))
                    return rows;

                var q = searchQuery.Trim();
                return rows.Where(r =>
                    r.AccNo.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    r.AccTitle.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    r.HeadName.Contains(q, StringComparison.OrdinalIgnoreCase));
            }
        }

        protected async Task LoadBudgetDataAsync()
        {
            try
            {
                isLoading = true;
                rows = await BudgetData.GetMonthwiseBudgetAsync(selectedFinYear, selectedHeadCode, isRevision);
                CalculateTotals();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Failed",
                    Detail = "Error loading income budget data: " + ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        protected void CalculateTotals()
        {
            var activeRows = FilteredRows.ToList();
            totals = new MonthwiseBudgetTotals
            {
                July = activeRows.Sum(r => r.MonthJuly),
                August = activeRows.Sum(r => r.MonthAugust),
                September = activeRows.Sum(r => r.MonthSeptember),
                October = activeRows.Sum(r => r.MonthOctober),
                November = activeRows.Sum(r => r.MonthNovember),
                December = activeRows.Sum(r => r.MonthDecember),
                January = activeRows.Sum(r => r.MonthJanuary),
                February = activeRows.Sum(r => r.MonthFebruary),
                March = activeRows.Sum(r => r.MonthMarch),
                April = activeRows.Sum(r => r.MonthApril),
                May = activeRows.Sum(r => r.MonthMay),
                June = activeRows.Sum(r => r.MonthJune)
            };
        }

        protected async Task OnCellChangedAsync(MonthwiseBudgetRow row, string column, object? rawVal)
        {
            decimal newVal = decimal.TryParse(rawVal?.ToString(), out var parsed) ? parsed : 0;

            switch (column.ToLowerInvariant())
            {
                case "month_july": row.MonthJuly = newVal; break;
                case "month_august": row.MonthAugust = newVal; break;
                case "month_september": row.MonthSeptember = newVal; break;
                case "month_october": row.MonthOctober = newVal; break;
                case "month_november": row.MonthNovember = newVal; break;
                case "month_december": row.MonthDecember = newVal; break;
                case "month_january": row.MonthJanuary = newVal; break;
                case "month_february": row.MonthFebruary = newVal; break;
                case "month_march": row.MonthMarch = newVal; break;
                case "month_april": row.MonthApril = newVal; break;
                case "month_may": row.MonthMay = newVal; break;
                case "month_june": row.MonthJune = newVal; break;
            }

            CalculateTotals();

            try
            {
                await BudgetData.SaveMonthwiseCellAsync(selectedFinYear, row.AccNo, column, newVal, isRevision);
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved",
                    Detail = $"Updated {row.AccNo} ({column.Replace("Month_", "")}: {newVal:N0})",
                    Duration = 2000
                });
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = $"Failed to save {row.AccNo}: {ex.Message}",
                    Duration = 4000
                });
            }
        }

        protected void PromptUpdateRevision()
        {
            showRevisionModal = true;
        }

        protected void CancelUpdateRevision()
        {
            showRevisionModal = false;
        }

        protected async Task ConfirmUpdateRevisionAsync()
        {
            showRevisionModal = false;
            try
            {
                isLoading = true;
                await BudgetData.UpdateMonthwiseRevisionsAsync(selectedFinYear);
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Revisions Updated",
                    Detail = $"Income revisions updated from Proposed amounts for FY {selectedFinYear}.",
                    Duration = 3500
                });
                await LoadBudgetDataAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = "Failed to update revisions: " + ex.Message,
                    Duration = 4000
                });
            }
            finally
            {
                isLoading = false;
            }
        }

        protected async Task PrintReportAsync()
        {
            try
            {
                var headParam = selectedHeadCode == "0" ? "" : $" AND {{VAccounts.Type}}='{selectedHeadCode}'";
                var request = new ReportRequest
                {
                    ReportName = isRevision ? ReportNames.Budgeting.BudgetMonthwiseRevision : ReportNames.Budgeting.BudgetMonthwise,
                    SelectionFormula = $"{{FinancialBudgeting.FinancialYear}}='{selectedFinYear}'{headParam}"
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Print Failed",
                    Detail = "Could not print report: " + ex.Message,
                    Duration = 4000
                });
            }
        }
    }
}

