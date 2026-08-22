using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Impulse.Pages.Payroll.OverTimeAuth
{
    public partial class OverTimeAuth : ComponentBase
    {
        [Parameter]
        [SupplyParameterFromQuery(Name = "returnUrl")]
        public string? ReturnUrl { get; set; }

        private bool IsLoading = false;
        private bool IsSaving = false;

        private DateTime AttendanceDate { get; set; } = DateTime.Today;

        private List<DepartmentListItemModel> Departments = new();
        private DepartmentListItemModel? SelectedDepartment { get; set; }

        private List<OverTimeAuthRowDto> Rows = new();
        private string SearchText { get; set; } = string.Empty;

        private List<OverTimeAuthRowDto> FilteredRows
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SearchText))
                    return Rows;

                return Rows.Where(r =>
                    r.EmpID.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.EmpName.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    (r.FName != null && r.FName.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }
        }

        private int AuthorizedCount => FilteredRows.Count(r => r.IsAuthorized);
        private double TotalOtHours => FilteredRows.Where(r => r.IsAuthorized).Sum(r => r.OtHrs);
        private decimal TotalAmount => FilteredRows.Where(r => r.IsAuthorized).Sum(r => r.Amount);

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Departments = await OverTimeAuthService.GetDepartmentsAsync();
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Initialization Error", ex.Message);
            }
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
                return Departments;

            return await Task.FromResult(Departments.Where(d =>
                d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
            ).ToList());
        }

        private async Task OnDepartmentChanged(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadOverTimeAuthListAsync();
        }

        private async Task OnAttendanceDateChanged(ChangeEventArgs e)
        {
            if (DateTime.TryParse(e?.Value?.ToString(), out DateTime parsed))
            {
                AttendanceDate = parsed.Date;
                await LoadOverTimeAuthListAsync();
            }
        }

        private async Task LoadOverTimeAuthListAsync()
        {
            Rows.Clear();

            if (SelectedDepartment == null)
            {
                await InvokeAsync(StateHasChanged);
                return;
            }

            IsLoading = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                Rows = await OverTimeAuthService.GetOverTimeAuthListAsync(SelectedDepartment.DeptID, AttendanceDate);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Error Loading Overtime List", ex.Message);
            }
            finally
            {
                IsLoading = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private void AuthorizeAll()
        {
            foreach (var row in FilteredRows)
            {
                row.IsAuthorized = true;
                if (row.OtHrs <= 0)
                {
                    row.OtHrs = 2; // Default 2 hrs if 0
                }
            }
        }

        private void UnauthorizeAll()
        {
            foreach (var row in FilteredRows)
            {
                row.IsAuthorized = false;
            }
        }

        private async Task SaveOverTimeAuthAsync()
        {
            if (SelectedDepartment == null)
            {
                NotificationService.ShowWarning("Warning", "Please select a department.");
                return;
            }

            IsSaving = true;
            await InvokeAsync(StateHasChanged);

            try
            {
                bool success = await OverTimeAuthService.SaveOverTimeAuthAsync(SelectedDepartment.DeptID, AttendanceDate, Rows);
                if (success)
                {
                    NotificationService.ShowSuccess("Success", $"Overtime authorizations saved successfully for {SelectedDepartment.Name}.");
                    await LoadOverTimeAuthListAsync();
                }
                else
                {
                    NotificationService.ShowError("Error", "Failed to save overtime authorizations.");
                }
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Save Error", ex.Message);
            }
            finally
            {
                IsSaving = false;
                await InvokeAsync(StateHasChanged);
            }
        }

        private async Task PrintDailyReport()
        {
            try
            {
                string formula = $"{{EmpOTHrs.DT}}=Date({AttendanceDate.Year}, {AttendanceDate.Month}, {AttendanceDate.Day})";
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DailyOverTime,
                    SelectionFormula = formula
                };
                await ReportNavigationService.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationService.ShowError("Report Error", ex.Message);
            }
        }

        private void NavigateBack()
        {
            if (!string.IsNullOrWhiteSpace(ReturnUrl))
            {
                NavigationManager.NavigateTo(ReturnUrl);
            }
            else
            {
                NavigationManager.NavigateTo("/payroll");
            }
        }
    }
}
