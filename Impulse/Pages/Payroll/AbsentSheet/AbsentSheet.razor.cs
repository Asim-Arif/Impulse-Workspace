using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Payroll.AbsentSheet
{
    public partial class AbsentSheet : ComponentBase
    {
        [Inject] private IAbsentSheetService AbsentSheetService { get; set; } = default!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private AbsentSheetFilterDto Filter { get; set; } = new();
        private AbsentSheetResultModel Result { get; set; } = new();

        private List<DepartmentListItemModel> Departments { get; set; } = new();
        private DepartmentListItemModel? SelectedDepartment { get; set; }

        private string SearchText { get; set; } = string.Empty;
        private bool IsLoading { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Departments = await EmployeeService.GetDepartmentsAsync(false, null);
                await OnRangeOptionChangedAsync(0); // Default to Today
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Initializing Page",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task OnRangeOptionChangedAsync(int option)
        {
            Filter.RangeOption = option;
            if (option == 0) // Today
            {
                Filter.FromDate = DateTime.Today;
                Filter.ToDate = DateTime.Today;
            }
            else if (option == 1) // Current Month
            {
                Filter.FromDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                Filter.ToDate = DateTime.Today;
            }
            // option == 2 is Custom Range; retain user selected dates

            await LoadAbsentSheetAsync();
        }

        private async Task LoadAbsentSheetAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Filter.DeptID = SelectedDepartment?.DeptID;
                Result = await AbsentSheetService.GetAbsentSheetAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Absent Sheet",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsLoading = false;
                StateHasChanged();
            }
        }

        private async Task OnDepartmentChangedAsync(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadAbsentSheetAsync();
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            await Task.CompletedTask;
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Departments;
            }
            return Departments
                .Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private IEnumerable<AbsentSheetRowModel> FilteredRows
        {
            get
            {
                if (Result.Rows == null) return Enumerable.Empty<AbsentSheetRowModel>();
                if (string.IsNullOrWhiteSpace(SearchText)) return Result.Rows;

                return Result.Rows.Where(r =>
                    r.EmpID.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.EmpName.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.Designation.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.DeptName.Contains(SearchText, StringComparison.OrdinalIgnoreCase));
            }
        }

        private async Task PrintSummaryReport()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.AbsentLeaveSummary,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "FromTo", $"'From {Filter.FromDate:dd-MMM-yyyy} To {Filter.ToDate:dd-MMM-yyyy}'" }
                    }
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task PrintDetailedReport()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DailyAttendance,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "FromTo", $"'From {Filter.FromDate:dd-MMM-yyyy} To {Filter.ToDate:dd-MMM-yyyy}'" }
                    }
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Report Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private void GoBack()
        {
            if (!string.IsNullOrEmpty(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll/dashboard");
            }
        }
    }
}
