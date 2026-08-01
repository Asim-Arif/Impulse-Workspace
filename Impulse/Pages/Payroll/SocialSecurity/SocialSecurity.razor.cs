using BlazorContextMenu;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Constants;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Radzen;

namespace Impulse.Pages.Payroll.SocialSecurity
{
    public partial class SocialSecurity : ComponentBase
    {
        [Inject] private ISocialSecurityService SocialSecurityService { get; set; } = default!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = default!;
        [Inject] private IReportNavigationService ReportNavigation { get; set; } = default!;
        [Inject] private IBlazorContextMenuService BlazorContextMenuService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private SocialSecurityFilterDto Filter { get; set; } = new();
        private SocialSecuritySheetModel Sheet { get; set; } = new();

        private List<DepartmentListItemModel> Departments { get; set; } = new();
        private DepartmentListItemModel? SelectedDepartment { get; set; }

        private string SearchText { get; set; } = string.Empty;
        private bool IsLoading { get; set; } = false;
        private bool IsGenerating { get; set; } = false;
        private bool IsSaving { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Departments = await EmployeeService.GetDepartmentsAsync(false, null);
                await LoadSheetAsync();
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Page",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
        }

        private async Task LoadSheetAsync()
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Filter.DeptID = SelectedDepartment?.DeptID;
                Sheet = await SocialSecurityService.GetSocialSecuritySheetAsync(Filter);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Social Security Sheet",
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

        private async Task GenerateSheetAsync()
        {
            IsGenerating = true;
            StateHasChanged();
            try
            {
                Filter.DeptID = SelectedDepartment?.DeptID;
                Sheet = await SocialSecurityService.GenerateSocialSecuritySheetAsync(Filter);

                if (Sheet.Rows.Count == 0)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Warning,
                        Summary = "No Data",
                        Detail = "No social security eligible employees found for selected month.",
                        Duration = 4000
                    });
                }
                else
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Info,
                        Summary = "Calculated",
                        Detail = $"Calculated figures for {Sheet.Rows.Count} employees. Click 'Save Sheet' to persist.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Calculation Error",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsGenerating = false;
                StateHasChanged();
            }
        }

        private async Task SaveSheetAsync()
        {
            if (Sheet.Rows == null || !Sheet.Rows.Any(r => r.IsSelected))
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "No Selection",
                    Detail = "Please select at least one employee row to save.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            StateHasChanged();
            try
            {
                var dto = new SaveSocialSecurityDto
                {
                    Month = Filter.Month,
                    Year = Filter.Year,
                    DeptID = SelectedDepartment?.DeptID,
                    Rows = Sheet.Rows.Where(r => r.IsSelected).ToList()
                };

                bool success = await SocialSecurityService.SaveSocialSecuritySheetAsync(dto);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Saved",
                        Detail = $"Social Security Sheet for {GetMonthName(Filter.Month)} {Filter.Year} saved successfully.",
                        Duration = 4000
                    });
                    await LoadSheetAsync();
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

        private async Task OnMonthYearChangedAsync()
        {
            await LoadSheetAsync();
        }

        private async Task OnDepartmentChangedAsync(DepartmentListItemModel? dept)
        {
            SelectedDepartment = dept;
            await LoadSheetAsync();
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

        private IEnumerable<SocialSecurityRowModel> FilteredRows
        {
            get
            {
                if (Sheet.Rows == null) return Enumerable.Empty<SocialSecurityRowModel>();
                if (string.IsNullOrWhiteSpace(SearchText)) return Sheet.Rows;

                return Sheet.Rows.Where(r =>
                    r.EmpID.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.EmpName.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.SocialNo.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    r.Designation.Contains(SearchText, StringComparison.OrdinalIgnoreCase));
            }
        }

        private void ToggleAllRows(bool select)
        {
            foreach (var r in FilteredRows)
            {
                r.IsSelected = select;
            }
        }

        private async Task PrintSSSheet()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.SocialSecuritySheet,
                    SelectionFormula = $"Month({{VSocialSecurity.DT}})={Filter.Month} and Year({{VSocialSecurity.DT}})={Filter.Year}"
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

        private async Task PrintSSForm()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.SocialSecurityForm,
                    SelectionFormula = $"Month({{VSocialSecurity.DT}})={Filter.Month} and Year({{VSocialSecurity.DT}})={Filter.Year}"
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

        private async Task PrintNonSS()
        {
            try
            {
                string monthName = GetMonthName(Filter.Month);
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.NonSocialSecurity,
                    FormulaValues = new Dictionary<string, object>
                    {
                        { "MonthYear", $"'{monthName} - {Filter.Year}'" }
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

        private async Task ShowPrintMenu(MouseEventArgs e)
        {
            await BlazorContextMenuService.ShowMenu("printOptionsMenu", (int)e.ClientX, (int)e.ClientY + 15);
        }

        private static string GetMonthName(int month)
        {
            return new DateTime(2000, month, 1).ToString("MMMM");
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
