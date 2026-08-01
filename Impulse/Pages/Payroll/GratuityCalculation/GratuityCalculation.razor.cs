using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace Impulse.Pages.Payroll.GratuityCalculation
{
    public partial class GratuityCalculation : ComponentBase
    {
        [Inject] private IGratuityService GratuityService { get; set; } = default!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;
        [Inject] private IJSRuntime JSRuntime { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private GratuitySheetModel Sheet { get; set; } = new();
        private List<DepartmentListItemModel> Departments { get; set; } = new();
        private DepartmentListItemModel? SelectedDepartment { get; set; }

        private int SelectedYear { get; set; } = DateTime.Today.Year;
        private string SearchText { get; set; } = string.Empty;

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Departments = await EmployeeService.GetDepartmentsAsync(false, null);
                await LoadGratuitySheetAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Gratuity",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task LoadGratuitySheetAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                string? deptId = SelectedDepartment?.DeptID;
                Sheet = await GratuityService.GetGratuitySheetAsync(SelectedYear, deptId);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Gratuity Sheet",
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

        private async Task OnYearChangedAsync()
        {
            await LoadGratuitySheetAsync();
        }

        private async Task OnDepartmentChangedAsync(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadGratuitySheetAsync();
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

        private IEnumerable<GratuityRowModel> FilteredRows
        {
            get
            {
                if (Sheet.Rows == null) return Enumerable.Empty<GratuityRowModel>();
                if (string.IsNullOrWhiteSpace(SearchText)) return Sheet.Rows;

                return Sheet.Rows.Where(r =>
                    r.EmpID.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.EmpName.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.Designation.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.DeptName.Contains(SearchText, StringComparison.OrdinalIgnoreCase));
            }
        }

        private async Task SaveGratuitySheetAsync()
        {
            if (Sheet.IsPosted)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Locked",
                    Detail = "Vouchers have been posted for this year. Editing is locked.",
                    Duration = 4000
                });
                return;
            }

            if (Sheet.Rows == null || !Sheet.Rows.Any())
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Data",
                    Detail = "There are no employees to save.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            StateHasChanged();
            try
            {
                var dto = new SaveGratuityDto
                {
                    Year = Sheet.Year,
                    Rows = Sheet.Rows
                };

                bool success = await GratuityService.SaveGratuitySheetAsync(dto);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = $"Gratuity calculation for {Sheet.Year} saved successfully.",
                        Duration = 4000
                    });
                    await LoadGratuitySheetAsync();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Save Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsSaving = false;
                StateHasChanged();
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
