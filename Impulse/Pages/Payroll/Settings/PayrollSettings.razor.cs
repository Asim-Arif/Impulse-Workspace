using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataAccessLibrary.Models.ViewModels.Payroll;
using Impulse.Services;
using Impulse.Services.Payroll;
using Microsoft.AspNetCore.Components;
using Radzen;

namespace Impulse.Pages.Payroll.Settings
{
    public partial class PayrollSettings : ComponentBase
    {
        [Inject] private IPayrollSettingsService SettingsService { get; set; } = default!;
        [Inject] private IEmployeeService EmployeeService { get; set; } = default!;
        [Inject] private NotificationService NotificationService { get; set; } = default!;
        [Inject] private NavigationManager NavigationManager { get; set; } = default!;

        [Parameter, SupplyParameterFromQuery]
        public string? ReturnUrl { get; set; }

        private List<DepartmentListItemModel> Departments { get; set; } = new();
        private DepartmentListItemModel? SelectedDepartment { get; set; }

        private List<EmpAttCategoryModel> Categories { get; set; } = new();
        private EmpAttCategoryModel? SelectedCategory { get; set; }

        private List<CategoryRuleItem> CategoryRules { get; set; } = new();
        private GeneralSettingsModel Model { get; set; } = new();

        private TimeOnly? EnterTimeOnly
        {
            get => TimeOnly.TryParse(Model.EnterTime, out var t) ? t : new TimeOnly(8, 0);
            set => Model.EnterTime = value?.ToString("HH:mm") ?? "08:00";
        }

        private TimeOnly? ExitTimeOnly
        {
            get => TimeOnly.TryParse(Model.ExitTime, out var t) ? t : new TimeOnly(17, 0);
            set => Model.ExitTime = value?.ToString("HH:mm") ?? "17:00";
        }

        private TimeOnly? LunchInTimeOnly
        {
            get => TimeOnly.TryParse(Model.LunchInTime, out var t) ? t : new TimeOnly(13, 0);
            set => Model.LunchInTime = value?.ToString("HH:mm") ?? "13:00";
        }

        private TimeOnly? LunchOutTimeOnly
        {
            get => TimeOnly.TryParse(Model.LunchOutTime, out var t) ? t : new TimeOnly(14, 0);
            set => Model.LunchOutTime = value?.ToString("HH:mm") ?? "14:00";
        }

        private TimeOnly? FridayBreakStartOnly
        {
            get => TimeOnly.TryParse(Model.FridayBreakStartTime, out var t) ? t : new TimeOnly(12, 30);
            set => Model.FridayBreakStartTime = value?.ToString("HH:mm") ?? "12:30";
        }

        private TimeOnly? FridayBreakEndOnly
        {
            get => TimeOnly.TryParse(Model.FridayBreakEndTime, out var t) ? t : new TimeOnly(14, 0);
            set => Model.FridayBreakEndTime = value?.ToString("HH:mm") ?? "14:00";
        }

        private TimeOnly? FridayCheckOutOnly
        {
            get => TimeOnly.TryParse(Model.FridayCheckOut, out var t) ? t : new TimeOnly(17, 30);
            set => Model.FridayCheckOut = value?.ToString("HH:mm") ?? "17:30";
        }

        private bool IsLoading { get; set; } = false;
        private bool IsSaving { get; set; } = false;
        private bool IsUpdatingAll { get; set; } = false;

        protected override async Task OnInitializedAsync()
        {
            IsLoading = true;
            try
            {
                Departments = await EmployeeService.GetDepartmentsAsync(false, null);
                Categories = await SettingsService.GetAttendanceCategoriesAsync();

                if (Departments.Any())
                {
                    SelectedDepartment = Departments.First();
                    await LoadDeptSettingsAsync(SelectedDepartment.DeptID);
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Initializing Settings",
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

        private async Task LoadDeptSettingsAsync(string deptId)
        {
            IsLoading = true;
            StateHasChanged();
            try
            {
                Model = await SettingsService.GetSettingsByDeptAsync(deptId);
                SelectedCategory = Categories.FirstOrDefault(c => c.EntryID == Model.EmpAttCat_RefID)
                                  ?? Categories.FirstOrDefault();

                if (SelectedCategory != null)
                {
                    Model.EmpAttCat_RefID = SelectedCategory.EntryID;
                    CategoryRules = await SettingsService.GetCategoryRulesByIdAsync(SelectedCategory.EntryID);
                }
                else
                {
                    CategoryRules = new List<CategoryRuleItem>();
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Error Loading Department Settings",
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
            if (dept != null)
            {
                await LoadDeptSettingsAsync(dept.DeptID);
            }
        }

        private async Task OnCategoryChangedAsync(ChangeEventArgs e)
        {
            if (int.TryParse(Convert.ToString(e.Value), out int catId))
            {
                Model.EmpAttCat_RefID = catId;
                SelectedCategory = Categories.FirstOrDefault(c => c.EntryID == catId);
                CategoryRules = await SettingsService.GetCategoryRulesByIdAsync(catId);
                StateHasChanged();
            }
        }

        private async Task SaveSettingsAsync(bool closeAfterSave)
        {
            if (SelectedDepartment == null)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Warning,
                    Summary = "Select Department",
                    Detail = "Please select a department to save settings.",
                    Duration = 3000
                });
                return;
            }

            IsSaving = true;
            StateHasChanged();
            try
            {
                Model.DeptID = SelectedDepartment.DeptID;
                bool success = await SettingsService.SaveSettingsAsync(Model);
                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Settings Saved",
                        Detail = $"Settings for '{SelectedDepartment.Name}' saved successfully.",
                        Duration = 4000
                    });

                    if (closeAfterSave)
                    {
                        GoBack();
                    }
                    else
                    {
                        // Save & New: Move to next department if available
                        int idx = Departments.IndexOf(SelectedDepartment);
                        if (idx >= 0 && idx < Departments.Count - 1)
                        {
                            SelectedDepartment = Departments[idx + 1];
                            await LoadDeptSettingsAsync(SelectedDepartment.DeptID);
                        }
                    }
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

        private async Task UpdateAllDepartmentsAsync()
        {
            IsUpdatingAll = true;
            StateHasChanged();
            try
            {
                bool success = await SettingsService.UpdateAllDepartmentTimingsAsync(
                    Model.EnterTime,
                    Model.ExitTime,
                    Model.LunchInTime,
                    Model.LunchOutTime,
                    Model.FridayBreakStartTime,
                    Model.FridayBreakEndTime,
                    Model.FridayCheckOut
                );

                if (success)
                {
                    NotificationService.Notify(new NotificationMessage
                    {
                        Severity = NotificationSeverity.Success,
                        Summary = "Updated All Departments",
                        Detail = "Shift timings updated successfully across all departments.",
                        Duration = 4000
                    });
                }
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage
                {
                    Severity = NotificationSeverity.Error,
                    Summary = "Update Failed",
                    Detail = ex.Message,
                    Duration = 5000
                });
            }
            finally
            {
                IsUpdatingAll = false;
                StateHasChanged();
            }
        }

        private async Task<IEnumerable<DepartmentListItemModel>> SearchDepartments(string searchText)
        {
            await Task.CompletedTask;
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return Departments;
            }
            return Departments.Where(d => d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));
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
