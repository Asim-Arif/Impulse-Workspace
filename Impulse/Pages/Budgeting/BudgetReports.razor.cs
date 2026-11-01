using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Budgeting;
using DataAccessLibrary.Models.ViewModels.Budgeting;
using Impulse.Constants;
using Impulse.Services;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.WebUtilities;
using Radzen;

namespace Impulse.Pages.Budgeting
{
    public partial class BudgetReports : ComponentBase
    {
        [Inject] public IBudgetDataAccess BudgetData { get; set; } = default!;
        [Inject] public NavigationManager Navigation { get; set; } = default!;
        [Inject] public NotificationService NotificationService { get; set; } = default!;
        [Inject] public IReportNavigationService ReportNavigation { get; set; } = default!;

        protected bool isExpenses;
        protected string reportType = "quarter"; // "quarter", "date-range", "glance"
        protected int quarterNo = 1;
        protected string pageTitle = "Budget Report";

        protected bool isLoading = true;

        protected List<FinancialYearOption> financialYears = new();
        protected string selectedFinYear = "";

        protected List<MasterHeadOption> masterHeads = new();
        protected MasterHeadOption? SelectedMasterHead;
        protected string selectedHeadCode => SelectedMasterHead?.Code ?? "0";

        protected DateTime fromDate = DateTime.Today;
        protected DateTime toDate = DateTime.Today;

        protected string searchQuery = "";
        protected List<BudgetReportRow> reportRows = new();

        protected decimal TotalProposed => reportRows.Sum(r => r.ProposedAmount);
        protected decimal TotalRevised => reportRows.Sum(r => r.RevisedAmount);
        protected decimal TotalActual => reportRows.Sum(r => r.ActualAmount);
        protected decimal TotalVariance => TotalActual - TotalProposed;

        protected IEnumerable<BudgetReportRow> FilteredRows
        {
            get
            {
                if (string.IsNullOrWhiteSpace(searchQuery))
                    return reportRows;

                var q = searchQuery.Trim();
                return reportRows.Where(r =>
                    r.AccNo.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    r.AccTitle.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                    r.HeadName.Contains(q, StringComparison.OrdinalIgnoreCase));
            }
        }

        protected override async Task OnInitializedAsync()
        {
            financialYears = BudgetData.GetFinancialYears(2010);
            if (financialYears.Count > 0)
            {
                selectedFinYear = financialYears[0].Code;
                fromDate = financialYears[0].StartDate;
                toDate = financialYears[0].EndDate;
            }

            ParseRouteAndQuery();
            await LoadMasterHeadsAsync();
            await LoadReportAsync();
        }

        protected override async Task OnParametersSetAsync()
        {
            ParseRouteAndQuery();
            await LoadMasterHeadsAsync();
            await LoadReportAsync();
        }

        private void ParseRouteAndQuery()
        {
            var uri = new Uri(Navigation.Uri);
            var path = uri.AbsolutePath.ToLowerInvariant();
            isExpenses = path.Contains("/expense");

            if (path.Contains("budget-at-glance"))
            {
                reportType = "glance";
                pageTitle = isExpenses ? "Budget at Glance (Expenses)" : "Budget at Glance";
            }
            else if (path.Contains("date-range"))
            {
                reportType = "date-range";
                pageTitle = isExpenses ? "Date Range Budget Report (Expenses)" : "Date Range Budget Report (Income)";
            }
            else if (path.Contains("/q1"))
            {
                reportType = "quarter";
                quarterNo = 1;
                pageTitle = isExpenses ? "1st Quarter Report (Expenses) [Jul - Sep]" : "1st Quarter Report (Income) [Jul - Sep]";
            }
            else if (path.Contains("/q2"))
            {
                reportType = "quarter";
                quarterNo = 2;
                pageTitle = isExpenses ? "2nd Quarter Report (Expenses) [Oct - Dec]" : "2nd Quarter Report (Income) [Oct - Dec]";
            }
            else if (path.Contains("/q3"))
            {
                reportType = "quarter";
                quarterNo = 3;
                pageTitle = isExpenses ? "3rd Quarter Report (Expenses) [Jan - Mar]" : "3rd Quarter Report (Income) [Jan - Mar]";
            }
            else if (path.Contains("/q4"))
            {
                reportType = "quarter";
                quarterNo = 4;
                pageTitle = isExpenses ? "4th Quarter Report (Expenses) [Apr - Jun]" : "4th Quarter Report (Income) [Apr - Jun]";
            }
            else
            {
                reportType = "quarter";
                quarterNo = 1;
                pageTitle = isExpenses ? "Quarterly Budget Report (Expenses)" : "Quarterly Budget Report (Income)";
            }

            var query = QueryHelpers.ParseQuery(uri.Query);
            if (query.TryGetValue("year", out var yr) && !string.IsNullOrWhiteSpace(yr))
            {
                selectedFinYear = yr.ToString();
                var fy = financialYears.FirstOrDefault(f => f.Code == selectedFinYear);
                if (fy != null)
                {
                    fromDate = fy.StartDate;
                    toDate = fy.EndDate;
                }
            }
        }

        private async Task LoadMasterHeadsAsync()
        {
            masterHeads = await BudgetData.GetMasterHeadsAsync(isExpenses);
            SelectedMasterHead = masterHeads.FirstOrDefault(h => h.Code == "0") ?? masterHeads.FirstOrDefault();
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
            await LoadReportAsync();
        }

        protected async Task OnFinYearChanged(ChangeEventArgs e)
        {
            selectedFinYear = e.Value?.ToString() ?? "";
            var fy = financialYears.FirstOrDefault(f => f.Code == selectedFinYear);
            if (fy != null)
            {
                fromDate = fy.StartDate;
                toDate = fy.EndDate;
            }
            await LoadReportAsync();
        }

        protected void OnSearchChanged(ChangeEventArgs e)
        {
            searchQuery = e.Value?.ToString() ?? "";
        }

        protected async Task SwitchModule(bool expenses)
        {
            isExpenses = expenses;
            await LoadMasterHeadsAsync();
            UpdatePageTitle();
            await LoadReportAsync();
        }

        protected async Task SwitchReportType(int quarter, bool isDateRange = false, bool isGlance = false)
        {
            if (isGlance)
            {
                reportType = "glance";
            }
            else if (isDateRange)
            {
                reportType = "date-range";
            }
            else
            {
                reportType = "quarter";
                quarterNo = quarter;
            }

            UpdatePageTitle();
            await LoadReportAsync();
        }

        private void UpdatePageTitle()
        {
            if (reportType == "glance")
            {
                pageTitle = isExpenses ? "Budget at Glance (Expenses)" : "Budget at Glance";
            }
            else if (reportType == "date-range")
            {
                pageTitle = isExpenses ? "Date Range Budget Report (Expenses)" : "Date Range Budget Report (Income)";
            }
            else
            {
                var qName = quarterNo == 1 ? "1st Quarter [Jul - Sep]" :
                            quarterNo == 2 ? "2nd Quarter [Oct - Dec]" :
                            quarterNo == 3 ? "3rd Quarter [Jan - Mar]" : "4th Quarter [Apr - Jun]";
                pageTitle = isExpenses ? $"{qName} (Expenses)" : $"{qName} (Income)";
            }
        }

        protected async Task LoadReportAsync()
        {
            try
            {
                isLoading = true;

                if (reportType == "glance")
                {
                    reportRows = await BudgetData.GetBudgetAtGlanceDataAsync(selectedFinYear, selectedHeadCode);
                }
                else if (reportType == "date-range")
                {
                    reportRows = await BudgetData.GetDateRangeReportDataAsync(selectedFinYear, fromDate, toDate, selectedHeadCode, isExpenses);
                }
                else
                {
                    reportRows = await BudgetData.GetQuarterlyReportDataAsync(selectedFinYear, quarterNo, selectedHeadCode, isExpenses);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Load Error",
                    Detail = "Error generating report data: " + ex.Message,
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
                string rptName;
                var parameters = new Dictionary<string, object>
                {
                    { "@FinancialYear", selectedFinYear }
                };

                if (reportType == "glance")
                {
                    rptName = ReportNames.Budgeting.Financial_Budgeting_Report_Glance;
                }
                else if (reportType == "date-range")
                {
                    rptName = isExpenses
                        ? ReportNames.Budgeting.Financial_Budgeting_Expenses_Report_Quarterwise
                        : ReportNames.Budgeting.Financial_Budgeting_Report_Quarterwise;

                    parameters.Add("@HeadCode", selectedHeadCode);
                    parameters.Add("@FromDate", fromDate.ToString("yyyy-MM-dd"));
                    parameters.Add("@ToDate", toDate.ToString("yyyy-MM-dd"));
                }
                else
                {
                    rptName = isExpenses
                        ? ReportNames.Budgeting.Financial_Budgeting_Expenses_Report_Quarterwise
                        : ReportNames.Budgeting.Financial_Budgeting_Report_Quarterwise;

                    parameters.Add("@QuarterNo", quarterNo);
                    parameters.Add("@HeadCode", selectedHeadCode);
                }

                var request = new ReportRequest
                {
                    ReportName = rptName,
                    Parameters = parameters
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

