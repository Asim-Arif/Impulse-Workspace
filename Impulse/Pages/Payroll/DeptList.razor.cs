using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using DataAccessLibrary.Models.ViewModels.Payroll;
using DataAccessLibrary.Models.ViewModels;
using Impulse.Services;
using Impulse.Services.Payroll;
using Impulse.Constants;

namespace Impulse.Pages.Payroll
{
    public partial class DeptList : ComponentBase
    {
        [Inject]
        private IDepartmentService DepartmentService { get; set; } = null!;

        [Inject]
        private IDBHelperService DBHelperService { get; set; } = null!;

        [Inject]
        private IReportNavigationService ReportNavigation { get; set; } = null!;

        [Inject]
        private IJSRuntime JS { get; set; } = null!;

        [Inject]
        private INotificationService NotificationServiceManager { get; set; } = null!;

        private List<DepartmentModel> departments = new();
        private bool showInactive = false;
        private string searchText = string.Empty;
        private bool isLoading = true;

        private NewDepartment editorModal = null!;

        protected override async Task OnInitializedAsync()
        {
            await RefreshList();
        }

        private async Task RefreshList()
        {
            try
            {
                isLoading = true;
                departments = await DepartmentService.GetDepartmentsListAsync(true);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", $"Failed to load departments: {ex.Message}");
            }
            finally
            {
                isLoading = false;
            }
        }

        private List<DepartmentModel> FilteredDepartments
        {
            get
            {
                var list = departments.AsEnumerable();
                if (!showInactive)
                {
                    list = list.Where(d => d.Active);
                }

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    list = list.Where(d => d.DeptID.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                        || d.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase)
                                        || (d.Description != null && d.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase)));
                }

                return list.ToList();
            }
        }

        private async Task OnShowInactiveChanged(ChangeEventArgs e)
        {
            showInactive = (bool)(e.Value ?? false);
            await Task.CompletedTask;
        }

        private async Task OpenAddModal()
        {
            await editorModal.OpenAsync(null);
        }

        private async Task OpenEditModal(string deptId)
        {
            await editorModal.OpenAsync(deptId);
        }

        private async Task ToggleActive(DepartmentModel dept)
        {
            bool newActiveState = !dept.Active;

            // VB6 child employee count check
            int employeeCount = DBHelperService.getSingleIntValue("Count(EmpID)", "Employees", $"Where DeptID = '{dept.DeptID}'");

            bool confirm = await JS.InvokeAsync<bool>("confirm", 
                $"This Will Affect All The Employees Of This Department.\nThere Are {employeeCount} Employees In This Department.\nDo You Want To Continue?");

            if (!confirm)
                return;

            try
            {
                await DepartmentService.SetDepartmentActiveStateAsync(dept.DeptID, newActiveState);
                NotificationServiceManager.ShowSuccess("Success", $"Department status updated successfully.");
                await RefreshList();
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Error", ex.Message);
            }
        }

        private async Task PrintDepartmentList()
        {
            try
            {
                var request = new ReportRequest
                {
                    ReportName = ReportNames.Payroll.DeptList,
                    SelectionFormula = ""
                };
                await ReportNavigation.PrintReportAsync(request);
            }
            catch (Exception ex)
            {
                NotificationServiceManager.ShowError("Print Error", ex.Message);
            }
        }

        private async Task OnSaved()
        {
            await RefreshList();
        }

        private Task OnCancelled()
        {
            return Task.CompletedTask;
        }

        private string ConvertWin32ColorToHex(int? win32Color)
        {
            if (win32Color == null || win32Color == 0)
                return "#000000";

            int color = win32Color.Value;
            int r = color & 0xFF;
            int g = (color >> 8) & 0xFF;
            int b = (color >> 16) & 0xFF;

            return $"#{r:X2}{g:X2}{b:X2}";
        }
    }
}
