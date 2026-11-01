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
    public partial class ExpenseBudgetEntry : ComponentBase
    {
        [Inject] public IBudgetDataAccess BudgetData { get; set; } = default!;
        [Inject] public NavigationManager Navigation { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigation { get; set; } = default!;

        protected bool isLoading = true;
        protected bool showRevisionModal = false;

        protected List<FinancialYearOption> financialYears = new();
        protected string selectedFinYear = "";

        protected List<MasterHeadOption> masterHeads = new();
        protected MasterHeadOption? SelectedMasterHead;
        protected string selectedHeadCode => SelectedMasterHead?.Code ?? "0";

        protected string searchQuery = "";

        protected List<ExpenseBudgetRow> rows = new();
        protected ExpenseBudgetTotals totals = new();

        protected override async Task OnInitializedAsync()
        {
            financialYears = BudgetData.GetFinancialYears(2010);
            if (financialYears.Count > 0)
            {
                selectedFinYear = financialYears[0].Code;
            }

            masterHeads = await BudgetData.GetMasterHeadsAsync(isExpenses: true);
            SelectedMasterHead = masterHeads.FirstOrDefault(h => h.Code == "0") ?? masterHeads.FirstOrDefault();

            await LoadBudgetDataAsync();
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

        protected IEnumerable<ExpenseBudgetRow> FilteredRows
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
                rows = await BudgetData.GetExpenseBudgetAsync(selectedFinYear, selectedHeadCode);
                CalculateTotals();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Load Failed",
                    Detail = "Error loading expense budget data: " + ex.Message,
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
            totals = new ExpenseBudgetTotals
            {
                TotalProposed = activeRows.Sum(r => r.Proposed),
                TotalRevised = activeRows.Sum(r => r.Revised)
            };
        }

        protected async Task OnCellChangedAsync(ExpenseBudgetRow row, string column, decimal newVal)
        {
            if (column.Equals("Revised", StringComparison.OrdinalIgnoreCase))
            {
                row.Revised = newVal;
            }
            else
            {
                row.Proposed = newVal;
            }

            CalculateTotals();

            try
            {
                await BudgetData.SaveExpenseCellAsync(selectedFinYear, row.AccNo, column, newVal);
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Saved",
                    Detail = $"Updated {row.AccNo} ({column}: {newVal:N0})",
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
                await BudgetData.UpdateExpenseRevisionsAsync(selectedFinYear);
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Success,
                    Summary = "Revisions Updated",
                    Detail = $"Expense revisions updated from Proposed amounts for FY {selectedFinYear}.",
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
                    ReportName = ReportNames.Budgeting.BudgetDateRange,
                    SelectionFormula = $"{{FinancialBudgeting_Expenses.FinancialYear}}='{selectedFinYear}'{headParam}"
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

